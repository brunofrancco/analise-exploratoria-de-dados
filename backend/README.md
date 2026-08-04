# Backend — Análise Exploratória de Dados

API em Python (FastAPI) que alimenta as abas de bibliotecas de análise automática do AnálisePro: **ydata-profiling**, **Sweetviz**, **AutoViz**, **D-Tale**, **Lux**, **skimpy** e **missingno**.

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
| GET | `/api/skimpy/{session_id}` | HTML |
| GET | `/api/missingno/{session_id}` | JSON com imagens base64 |
| GET | `/api/autoviz/{session_id}` | JSON com imagens base64 |
| GET | `/api/lux/{session_id}` | JSON com recomendações |
| GET | `/api/dtale/{session_id}` | JSON `{ url }` (sessão interativa embutida em `/dtale-app`) |
| GET | `/health` | status |

## Deploy (Render)

Este backend é publicado separadamente do front-end (que fica no GitHub Pages, e não executa Python). Configuração pronta em `render.yaml` na raiz do projeto:

- Runtime: Python
- Diretório raiz do serviço: `backend`
- Build: `pip install -r requirements.txt`
- Start: `uvicorn main:app --host 0.0.0.0 --port $PORT`

Depois do primeiro deploy, atualize a constante `API_BASE` em `src/AnalisePro.jsx` (front-end) e a variável de ambiente `PUBLIC_BASE_URL` no Render com a URL pública real do serviço.

## Observações importantes

- Cada biblioteca é importada dentro da própria rota (import "preguiçoso"): se uma delas falhar ao instalar, apenas aquele endpoint fica indisponível — o restante da API continua funcionando.
- `ydata-profiling`, `Sweetviz`, `AutoViz`, `D-Tale` e `Lux` são bibliotecas pesadas com dependências que mudam com frequência; é normal precisar ajustar versões no `requirements.txt` após o primeiro deploy, observando os logs de build/runtime no Render.
- Sessões em memória expiram após 2 horas e são perdidas caso o serviço reinicie (comum no plano gratuito do Render, que hiberna após inatividade).
