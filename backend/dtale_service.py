"""
Serviço dedicado ao D-Tale.

Por que um serviço separado?
-----------------------------
O bundle JS do D-Tale chama sua própria API a partir da RAIZ do domínio
(ex.: /dtale/data/1, /dtale/dtypes/1). Quando o D-Tale é embutido sob um
subcaminho de outra aplicação (ex.: /dtale-app dentro do backend principal
em FastAPI), essas chamadas caem em 404 e a interface fica em branco — o
bundle não tem como saber que foi montado em um prefixo. Rodando o D-Tale
como o próprio app na raiz do seu domínio, esse problema desaparece.

Este serviço expõe apenas:
  POST /load-session  — recebe {columns, rows} (a base ativa da sessão do
                         usuário, repassada pelo backend principal) e inicia
                         uma sessão do D-Tale, devolvendo o caminho da
                         interface interativa.
  GET  /health         — healthcheck usado pelo Render.

Todo o resto das rotas (/dtale/...) é servido pelo próprio D-Tale.
"""

import os

import pandas as pd
from dtale.app import build_app
from dtale.views import startup
from flask import jsonify, request

app = build_app(reaper_on=False)


@app.route("/load-session", methods=["POST"])
def load_session():
    payload = request.get_json(force=True, silent=True) or {}
    columns = payload.get("columns") or []
    rows = payload.get("rows") or []
    try:
        df = pd.DataFrame(rows, columns=columns)
        instance = startup(data=df, ignore_duplicate=True)
        data_id = getattr(instance, "_data_id", None) or getattr(instance, "data_id", None)
        return jsonify({"path": f"/dtale/main/{data_id}"})
    except Exception as exc:  # noqa: BLE001
        return jsonify({"error": str(exc)}), 500


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8001))
    app.run(host="0.0.0.0", port=port)
