# Excel → Word API

FastAPI service that converts an uploaded Excel file into a Word document using a template.

## Files
- `main.py`: API implementation
- `requirements.txt`: Python dependencies
- `Dockerfile`: Render-ready container
- `template.docx`: Replace this placeholder with your real Word template

## Local run
```bash
cd excel-word-api
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 10000
```

## Endpoint
- `POST /generate` with a `multipart/form-data` file field named `file`

## Deploy
- Render: create a new Web Service and deploy as Docker using this repo

## Netlify note
This repo is **backend-only** and does not include a static `index.html`. If Netlify shows "Page not found", create and deploy a separate frontend folder (or set Netlify's Publish Directory to the folder that contains your `index.html`).
