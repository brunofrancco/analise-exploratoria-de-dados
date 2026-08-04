# Análise Exploratória de Dados

Aplicação web (React + Vite) para upload de bases de dados (CSV, TSV, Excel, JSON, Parquet) e geração automática de relatórios de análise exploratória: diagnóstico de qualidade, estatísticas descritivas, correlações e dashboards — tudo processado no navegador, sem enviar dados a um servidor.

Além disso, a aplicação integra duas bibliotecas Python de análise automática (via um backend próprio, em `backend/`), alimentadas pelo mesmo upload único: **ydata-profiling** e **Sweetviz**.

## Rodando localmente

```bash
npm install
npm run dev
```

## Build de produção

```bash
npm run build
```

Gera a pasta `dist/`, publicada automaticamente no GitHub Pages a cada push na branch `main` (veja `.github/workflows/deploy.yml`).

## Backend Python

As abas de ydata-profiling e Sweetviz dependem de um backend em Python (FastAPI), publicado separadamente (GitHub Pages não executa Python). Veja `backend/README.md` para instruções de execução local e deploy.
