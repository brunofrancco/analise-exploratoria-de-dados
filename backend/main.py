"""
Backend da Análise Exploratória de Dados.

API em FastAPI que recebe UMA base de dados (upload único) e a mantém em
memória, compartilhada por sete endpoints — um por biblioteca de análise
automática: ydata-profiling, Sweetviz, AutoViz, D-Tale, Lux, skimpy e
missingno.

Cada biblioteca é importada dentro da própria função de rota (import
"preguiçoso"), de forma que, se uma delas falhar ao instalar/importar no
servidor, apenas aquele endpoint específico fica indisponível — o resto da
API continua funcionando normalmente.
"""

import base64
import contextlib
import io
import os
import time
import uuid
from html import escape
from typing import Dict

import matplotlib
matplotlib.use("Agg")  # backend sem interface gráfica, necessário em servidor

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse

# URL pública deste serviço depois do deploy (ajuste via variável de
# ambiente PUBLIC_BASE_URL no Render). Usada para montar links absolutos,
# por exemplo para a sessão interativa do D-Tale.
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "https://analise-exploratoria-backend.onrender.com")

app = FastAPI(title="Análise Exploratória de Dados — API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------------------------
# Armazenamento em memória do DataFrame compartilhado (por sessão de upload)
# ---------------------------------------------------------------------------
SESSIONS: Dict[str, dict] = {}
SESSION_TTL_SECONDS = 2 * 60 * 60  # 2 horas


def _cleanup_sessions() -> None:
    now = time.time()
    expired = [sid for sid, s in SESSIONS.items() if now - s["created_at"] > SESSION_TTL_SECONDS]
    for sid in expired:
        SESSIONS.pop(sid, None)


def _get_df(session_id: str) -> pd.DataFrame:
    _cleanup_sessions()
    session = SESSIONS.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail="Sessão não encontrada ou expirada. Faça upload da base novamente.")
    return session["df"]


def _read_dataframe(filename: str, content: bytes) -> pd.DataFrame:
    name = (filename or "").lower()
    buf = io.BytesIO(content)
    try:
        if name.endswith(".csv") or name.endswith(".txt"):
            return pd.read_csv(buf, sep=None, engine="python")
        if name.endswith(".tsv"):
            return pd.read_csv(buf, sep="\t")
        if name.endswith(".xlsx") or name.endswith(".xls"):
            return pd.read_excel(buf)
        if name.endswith(".parquet"):
            return pd.read_parquet(buf)
        if name.endswith(".json"):
            return pd.read_json(buf)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"Falha ao ler o arquivo: {exc}") from exc
    raise HTTPException(status_code=400, detail="Formato não suportado. Use CSV, TSV, XLSX, XLS, JSON ou Parquet.")


def _fig_to_data_uri(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=110)
    plt_close(fig)
    buf.seek(0)
    return "data:image/png;base64," + base64.b64encode(buf.read()).decode("ascii")


def plt_close(fig) -> None:
    import matplotlib.pyplot as plt
    plt.close(fig)


# ---------------------------------------------------------------------------
# Rotas gerais
# ---------------------------------------------------------------------------
@app.get("/")
def root():
    return {"status": "ok", "service": "analise-exploratoria-backend"}


@app.get("/health")
def health():
    return {"status": "ok", "sessions_ativas": len(SESSIONS)}


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    content = await file.read()
    df = _read_dataframe(file.filename, content)
    if df.empty:
        raise HTTPException(status_code=400, detail="A base carregada está vazia.")
    session_id = uuid.uuid4().hex
    SESSIONS[session_id] = {"df": df, "filename": file.filename, "created_at": time.time()}
    return {"session_id": session_id, "rows": len(df), "columns": [str(c) for c in df.columns]}


# ---------------------------------------------------------------------------
# ydata-profiling — relatório completo de EDA
# ---------------------------------------------------------------------------
@app.get("/api/ydata-profiling/{session_id}", response_class=HTMLResponse)
def ydata_profiling_report(session_id: str):
    df = _get_df(session_id)
    try:
        from ydata_profiling import ProfileReport
        profile = ProfileReport(
            df,
            title="Relatório ydata-profiling",
            minimal=len(df) > 20000,
            explorative=len(df) <= 20000,
        )
        html = profile.to_html()
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Falha ao gerar relatório ydata-profiling: {exc}") from exc
    return HTMLResponse(content=html)


# ---------------------------------------------------------------------------
# Sweetviz — relatório visual e comparação de variáveis
# ---------------------------------------------------------------------------
@app.get("/api/sweetviz/{session_id}", response_class=HTMLResponse)
def sweetviz_report(session_id: str):
    df = _get_df(session_id)
    try:
        import sweetviz as sv
        report = sv.analyze(df)
        tmp_path = f"/tmp/sweetviz_{session_id}.html"
        report.show_html(filepath=tmp_path, open_browser=False, layout="vertical")
        with open(tmp_path, "r", encoding="utf-8") as f:
            html = f.read()
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Falha ao gerar relatório Sweetviz: {exc}") from exc
    return HTMLResponse(content=html)


# ---------------------------------------------------------------------------
# skimpy — resumo estatístico do DataFrame
# ---------------------------------------------------------------------------
@app.get("/api/skimpy/{session_id}", response_class=HTMLResponse)
def skimpy_report(session_id: str):
    df = _get_df(session_id)
    try:
        from skimpy import skim
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            skim(df)
        text = buf.getvalue()
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Falha ao gerar resumo skimpy: {exc}") from exc
    html = (
        "<!DOCTYPE html><html><head><meta charset='utf-8'>"
        "<style>body{background:#12141A;color:#E8EAF0;margin:0;padding:20px;}"
        "pre{font-family:'JetBrains Mono',Consolas,monospace;font-size:12.5px;white-space:pre-wrap;}</style>"
        f"</head><body><pre>{escape(text)}</pre></body></html>"
    )
    return HTMLResponse(content=html)


# ---------------------------------------------------------------------------
# missingno — análise e visualização de valores ausentes
# ---------------------------------------------------------------------------
@app.get("/api/missingno/{session_id}")
def missingno_report(session_id: str):
    df = _get_df(session_id)
    try:
        import missingno as msno
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Biblioteca missingno indisponível: {exc}") from exc

    result: Dict[str, str] = {}

    def _capture(fn):
        try:
            ax = fn(df)
            fig = ax.get_figure() if hasattr(ax, "get_figure") else ax
            return _fig_to_data_uri(fig)
        except Exception:  # noqa: BLE001
            return None

    matrix = _capture(msno.matrix)
    if matrix:
        result["matrix"] = matrix
    bar = _capture(msno.bar)
    if bar:
        result["bar"] = bar
    if df.isna().sum().sum() > 0 and df.shape[1] > 1:
        heatmap = _capture(msno.heatmap)
        if heatmap:
            result["heatmap"] = heatmap

    if not result:
        raise HTTPException(status_code=500, detail="Não foi possível gerar nenhuma visualização de ausentes para esta base.")
    return JSONResponse(result)


# ---------------------------------------------------------------------------
# AutoViz — geração automática de gráficos
# ---------------------------------------------------------------------------
@app.get("/api/autoviz/{session_id}")
def autoviz_report(session_id: str):
    df = _get_df(session_id)
    import glob
    import shutil
    import tempfile

    tmp_dir = tempfile.mkdtemp(prefix="autoviz_")
    try:
        from autoviz.AutoViz_Class import AutoViz_Class

        av = AutoViz_Class()
        av.AutoViz(
            filename="",
            sep=",",
            depVar="",
            dfte=df,
            header=0,
            verbose=0,
            lowess=False,
            chart_format="png",
            max_rows_analyzed=min(len(df), 150000),
            max_cols_analyzed=min(len(df.columns), 30),
            save_plot_dir=tmp_dir,
        )
        images = []
        for path in sorted(glob.glob(os.path.join(tmp_dir, "**", "*.png"), recursive=True)):
            with open(path, "rb") as f:
                images.append("data:image/png;base64," + base64.b64encode(f.read()).decode("ascii"))
        return JSONResponse({"images": images})
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Falha ao gerar gráficos com AutoViz: {exc}") from exc
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Lux — sugestão automática de visualizações
# ---------------------------------------------------------------------------
@app.get("/api/lux/{session_id}")
def lux_report(session_id: str):
    df = _get_df(session_id).copy()
    try:
        import lux  # noqa: F401  (registra o accessor .recommendation em pandas)

        recommendation = df.recommendation or {}
        items = []
        for action, vis_list in recommendation.items():
            for vis in list(vis_list)[:4]:
                entry = {"name": f"{action}", "description": str(vis)}
                try:
                    chart = vis.to_matplotlib()
                    fig = chart[0] if isinstance(chart, tuple) else chart
                    entry["image"] = _fig_to_data_uri(fig)
                except Exception:  # noqa: BLE001
                    entry["image"] = None
                items.append(entry)
        return JSONResponse({"recommendations": items})
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Falha ao gerar recomendações Lux: {exc}") from exc


# ---------------------------------------------------------------------------
# D-Tale — exploração interativa dos dados
# Embutido via ponte WSGI->ASGI (a2wsgi) no mesmo processo/porta do FastAPI,
# já que serviços como o Render expõem apenas uma porta pública.
# ---------------------------------------------------------------------------
_dtale_mounted = False


def _ensure_dtale_mounted():
    global _dtale_mounted
    if _dtale_mounted:
        return
    import dtale.app as dtale_app
    from a2wsgi import WSGIMiddleware

    flask_app = dtale_app.build_app(reaper_on=False)
    app.mount("/dtale-app", WSGIMiddleware(flask_app))
    _dtale_mounted = True


@app.get("/api/dtale/{session_id}")
def dtale_report(session_id: str):
    df = _get_df(session_id)
    try:
        _ensure_dtale_mounted()
        from dtale.views import startup

        instance = startup(data=df, ignore_duplicate=True)
        data_id = getattr(instance, "_data_id", None) or getattr(instance, "data_id", None)
        url = f"{PUBLIC_BASE_URL}/dtale-app/dtale/main/{data_id}"
        return JSONResponse({"url": url})
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Falha ao iniciar o D-Tale: {exc}") from exc


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), reload=False)
