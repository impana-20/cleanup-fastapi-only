# Nivora Customer Chatbot API

FastAPI-only backend for the Nivora customer-support chatbot. The project does not include a Streamlit application or a browser chat interface. Use Swagger UI to explore and test the API.

## Project structure

```text
.
├── main.py
├── rag_engine.py
├── requirements.txt
├── Dockerfile
├── .dockerignore
├── .gitignore
├── .env.example
├── data/
│   └── nivora_knowledge_base.txt
└── tests/
    └── test_rag_engine.py
```

## Knowledge base

`data/nivora_knowledge_base.txt` combines the supplied company information and Nivora FAQ text. The source FAQ identifies itself as a draft for internal review, not approved customer-facing policy. The chatbot must not present unapproved policy proposals as confirmed Nivora policy.

The product catalogue workbook is not included. Product endpoints therefore report an unavailable/empty catalogue until an approved workbook is configured.

## Local setup

Use Python 3.12 or later.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
```

Set `OPENAI_API_KEY` in `.env` if AI-generated responses are required. Do not commit `.env` or publish API keys.

Start the API:

```powershell
python -m uvicorn main:app --reload
```

Open Swagger UI:

- `http://127.0.0.1:8000/docs`
- Alternative documentation: `http://127.0.0.1:8000/redoc`
- Health check: `http://127.0.0.1:8000/api/health`

## API endpoints

| Method | Endpoint | Purpose |
|---|---|---|
| GET | `/` | API service information |
| GET | `/api/health` | Backend and knowledge-base status |
| GET | `/api/categories` | Product category summary |
| GET | `/api/products` | Search the configured product catalogue |
| POST | `/api/chat` | Submit a chatbot message |
| POST | `/api/reload-catalogue` | Reload the configured product workbook |
| GET | `/api/faq-status` | Check loaded FAQ count and question numbers |

Example request for `POST /api/chat`:

```json
{
  "message": "Tell me about Nivora",
  "history": []
}
```

## Run tests

```powershell
python -m pytest -q
```

## Docker

Build and run locally:

```powershell
docker build -t nivora-chatbot-api .
docker run --env-file .env -p 8000:8000 nivora-chatbot-api
```

Then open `http://localhost:8000/docs`.

## Deployment notes

The container listens on port `8000`. Configure the `OPENAI_API_KEY` as a secure secret in Azure Container Apps; do not bake it into the image or commit it to source control. Configure `NIVORA_PRODUCT_FILE` only when an approved catalogue workbook is available inside the container.
