import os
import tempfile
from typing import Optional

from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from config import Config
from db import get_connection

# you will replace these imports with your real api
# i'm guessing the names; adjust to actual functions.
from embedding.minilm import MiniLMEmbedder
from text_extraction.basic_extraction import extract_text_from_path  # TODO: fix name


_embedder = MiniLMEmbedder()


def _embed_text(text: str):
    """Encode the uploaded text with the cached MiniLM embedder."""

    return _embedder.encode([text])[0]


app = FastAPI()
templates = Jinja2Templates(directory="templates")


@app.get("/", response_class=HTMLResponse)
async def index_get(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "error": None,
            "results": [],
            "query_text": None,
            "server_root": "",
            "distance_metric": "cosine",
        },
    )


@app.post("/", response_class=HTMLResponse)
async def index_post(
    request: Request,
    file: Optional[UploadFile] = File(None),
    server_root: str = Form(""),
):
    error = None
    results = []
    query_text = None
    distance_metric = "cosine"  # just for display rn

    if not file or not file.filename:
        error = "please choose a file."
    else:
        # save uploaded file to a temp location
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name

        try:
            # 1) extract text using your existing pipeline
            query_text = extract_text_from_path(tmp_path)

            if not query_text or not query_text.strip():
                error = "no text could be extracted from that file."
            else:
                # 2) embed text using your existing minilm embedding
                query_vec = _embed_text(query_text)

                # make sure it's a plain python list, psycopg wants that
                query_vec = list(map(float, query_vec))

                # 3) run similarity search in postgres
                with get_connection() as conn:
                    with conn.cursor() as cur:
                        # using cosine distance (<=>) on minilm_emb
                        # we also join to files + file_locations for paths
                        sql = """
                            SELECT
                                fc.file_hash,
                                fl.file_server_directories,
                                fl.filename,
                                (fc.minilm_emb <=> %(query_vec)s) AS distance
                            FROM file_contents fc
                            JOIN files f
                              ON f.hash = fc.file_hash
                            LEFT JOIN file_locations fl
                              ON fl.file_id = f.id
                            WHERE fc.minilm_emb IS NOT NULL
                            ORDER BY fc.minilm_emb <=> %(query_vec)s
                            LIMIT %(top_k)s;
                        """

                        cur.execute(
                            sql,
                            {
                                "query_vec": query_vec,
                                "top_k": Config.TOP_K,
                            },
                        )
                        rows = cur.fetchall()

                # 4) post-process rows and build full paths
                for row in rows:
                    directory = row.get("file_server_directories") or ""
                    filename = row.get("filename") or ""
                    distance = row.get("distance")

                    # normalize separators a bit; you can tweak for your env
                    pieces = [p for p in [server_root, directory, filename] if p]
                    full_path = os.path.join(*pieces) if pieces else ""

                    results.append(
                        {
                            "file_hash": row["file_hash"],
                            "directory": directory,
                            "filename": filename,
                            "full_path": full_path,
                            "distance": float(distance) if distance is not None else None,
                        }
                    )

        finally:
            # clean up temp file
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "error": error,
            "results": results,
            "query_text": query_text,
            "server_root": server_root,
            "distance_metric": distance_metric,
        },
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=5000)
