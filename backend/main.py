"""
Backend da Análise Exploratória de Dados.

API em FastAPI que recebe UMA base de dados (upload único) e a mantém em
memória, compartilhada pelos endpoints de análise automática: ydata-profiling
e Sweetviz.

Cada biblioteca é importada dentro da própria função de rota (import
"preguiçoso"), de forma que, se uma delas falhar ao instalar/importar no
servidor, apenas aquele endpoint específico fica indisponível — o resto da
API continua funcionando normalmente.
"""

import io
import os
import time
import uuid
from typing import Dict

import matplotlib
matplotlib.use("Agg")  # backend sem interface gráfica, necessário em servidor

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse

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


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:app", host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), reload=False)
