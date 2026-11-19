# GST Smart Assistant – Phase 1

A zero-LLM, rules-based assistant for the existing GST 2B reconciliation workflow. It keeps reconciliation decisions deterministic and adds chat, summaries, anomaly queries, matching explanations, Excel exports, history, backup and restore.

## Features
- Upload one `.xlsx` workbook containing named Excel tables `Pur` and `Table2B`/`T2B`/`Portal2B` (or `2B` where accepted).
- Reconciliation uses the GST matching logic carried forward from the supplied 18 Nov source.
- Separate chat sessions with predefined prompts and a free-text input.
- Neon PostgreSQL history plus a local JSON mirror. Render's local filesystem is temporary on the free plan; Neon is the durable store.
- Backup/restore ZIP containing analyses, chat sessions/messages and settings. Uploaded source workbooks are not included.
- Excel filename is based on the most common Purchase invoice month, e.g. `Report of Oct 26.xlsx`. Chat export uses `GST 2B Reco - Oct 2026.html`.

## Required Render environment variables
- `APP_USERNAME` – login username (default `admin`)
- `APP_PASSWORD` – set a strong password
- `SECRET_KEY` – generate a long random secret (Render can generate this)
- `DATABASE_URL` – Neon PostgreSQL connection string; use the pooled connection string and ensure SSL is enabled (`sslmode=require`).
- `COOKIE_SECURE=true`

## Deploy
1. Push this project to a private GitHub repository.
2. Create a Render Web Service from the repository, or use the included `render.yaml`.
3. Set the environment variables above.
4. Deploy. Visit `/health`; it reports whether a database URL is configured.
5. Log in and upload the GST workbook.

## Local run
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
set APP_PASSWORD=change-this
set SECRET_KEY=replace-with-a-long-random-secret
set DATABASE_URL=your-neon-connection-string
python -m flask --app app run
```

If `DATABASE_URL` is absent, the app falls back to local JSON history. This is suitable only for local testing, not durable Render storage.

## Privacy and limits
- No external LLM, Gemini, Ollama or model API is called. The chat answers supported questions using Python calculations over the selected reconciliation report.
- Neon is an external database provider because you selected Neon for persistent history. The app stores reconciliation result JSON, chat history and settings there. It does not store the uploaded original workbook.
- Local JSON is a secondary mirror. On Render Free it can disappear after restart/redeploy; use the backup download regularly.
- Phase 1 chat is intent-based, not a general-purpose generative AI model. Unsupported questions receive suggested supported queries.
