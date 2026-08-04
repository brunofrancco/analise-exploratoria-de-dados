# Backend — Análise Exploratória de Dados

API em Python (FastAPI) que alimenta as abas de bibliotecas de análise automática do AnálisePro: **ydata-profiling** e **Sweetviz**.

A base carregada pelo usuário é enviada uma única vez via `POST /api/upload`, mantida em memória (por sessão), e reutilizada por todos os endpoints abaixo — sem necessidade de novo upload ao trocar de aba.

## Rodando localmente

```bash
cd backend
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

## Endpoints

| Método | Rota | Retorno |
| --- | --- | --- |
| POST | `/api/upload` | `{ session_id, rows, columns }` |
| GET | `/api/ydata-profiling/{session_id}` | HTML |
| GET | `/api/sweetviz/{session_id}` | HTML |
| GET | `/health` | status |

## Deploy (Render)

Este backend é publicado separadamente do front-end (que fica no GitHub Pages, e não executa Python). Configuração pronta em `render.yaml` na raiz do projeto:

- Runtime: Python
- Diretório raiz do serviço: `backend`
- Build: `pip install -r requirements.txt`
- Start: `uvicorn main:app --host 0.0.0.0 --port $PORT`
- `PYTHON_VERSION=3.11.9` (a versão mais recente do Python no Render costuma não ter wheels prontos para bibliotecas de data science como o ydata-profiling).

## Observações importantes

- Cada biblioteca é importada dentro da própria rota (import "preguiçoso"): se uma delas falhar ao instalar, apenas aquele endpoint fica indisponível — o restante da API continua funcionando.
- Sessões em memória expiram após 2 horas e são perdidas caso o serviço reinicie (comum no plano gratuito do Render, que hiberna após inatividade).
