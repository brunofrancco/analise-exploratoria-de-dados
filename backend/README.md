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

Este backend é publicado separadamente do front-end (que fica no GitHub Pages, e não executa Python). Configuração pronta em `render.yaml` na raiz do projeto, com **dois serviços**:

1. `analise-exploratoria-backend` — a API principal (este `main.py`).
   - Build: `pip install -r requirements.txt`
   - Start: `uvicorn main:app --host 0.0.0.0 --port $PORT`
2. `analise-exploratoria-dtale` — serviço dedicado só para o D-Tale (`dtale_service.py`).
   - Build: `pip install -r requirements-dtale.txt`
   - Start: `gunicorn -w 1 -b 0.0.0.0:$PORT dtale_service:app`

**Por que dois serviços?** O bundle JS do D-Tale chama sua própria API a partir da raiz do domínio (ex.: `/dtale/data/1`). Embutido sob um subcaminho de outra aplicação, essas chamadas retornam 404 e a tela fica em branco. Rodando o D-Tale como o próprio app na raiz do seu domínio, o problema desaparece.

Depois do primeiro deploy de cada serviço, ajuste no Render:
- `PUBLIC_BASE_URL` no serviço principal, com a URL pública real dele.
- `DTALE_SERVICE_URL` no serviço principal, com a URL pública real do serviço `analise-exploratoria-dtale`.
- `PYTHON_VERSION=3.11.9` em ambos (a versão mais recente do Python no Render costuma não ter wheels prontos para várias dessas bibliotecas de data science).

E atualize a constante `API_BASE` em `src/AnalisePro.jsx` (front-end) com a URL do serviço principal, se ela mudar.

## Observações importantes

- Cada biblioteca é importada dentro da própria rota (import "preguiçoso"): se uma delas falhar ao instalar, apenas aquele endpoint fica indisponível — o restante da API continua funcionando.
- `ydata-profiling`, `Sweetviz`, `AutoViz`, `D-Tale` e `Lux` são bibliotecas pesadas com dependências que mudam com frequência; é normal precisar ajustar versões no `requirements.txt`/`requirements-dtale.txt` após o primeiro deploy, observando os logs de build/runtime no Render.
- Sessões em memória expiram após 2 horas e são perdidas caso o serviço reinicie (comum no plano gratuito do Render, que hiberna após inatividade) — inclusive a sessão do D-Tale, que vive no serviço `analise-exploratoria-dtale`.
