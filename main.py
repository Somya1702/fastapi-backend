from fastapi import FastAPI
from fastapi.responses import HTMLResponse

app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)  # Disable API docs

@app.get("/", response_class=HTMLResponse)
async def home():
    return "<h1>SCN REPLY</h1>"
