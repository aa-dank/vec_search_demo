import os
import shutil
import tempfile
from typing import Optional

from fastapi import FastAPI, File, Form, UploadFile, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from config import Config
from db import get_connection
from utils import extract_server_dirs

from embedding.minilm import MiniLMEmbedder

from text_extraction.pdf_extraction import PDFTextExtractor
from text_extraction.basic_extraction import TextFileTextExtractor, TikaTextExtractor, get_extractor_for_file
from text_extraction.image_extraction import ImageTextExtractor
from text_extraction.office_doc_extraction import PresentationTextExtractor, SpreadsheetTextExtractor, WordFileTextExtractor
from text_extraction.web_extraction import HtmlTextExtractor, EmailTextExtractor
from text_extraction.extraction_utils import common_char_replacements, strip_diacritics, normalize_unicode, normalize_whitespace

USER_SERVER_MOUNT_PATH = "N:\\PPDO\\Records"

# Initialize extractors and Tika fallback
pdf_extractor = PDFTextExtractor()
txt_extractor = TextFileTextExtractor()
image_extractor = ImageTextExtractor()
presentation_extractor = PresentationTextExtractor()
spreadsheet_extractor = SpreadsheetTextExtractor()
word_extractor = WordFileTextExtractor()
html_extractor = HtmlTextExtractor()
email_extractor = EmailTextExtractor()
tika_extractor = TikaTextExtractor()
extractors_list = [
    pdf_extractor,
    txt_extractor,
    image_extractor,
    presentation_extractor,
    spreadsheet_extractor,
    word_extractor,
    html_extractor,
    email_extractor,
]


def extract_and_normalize_text(file_path: str) -> str:
    """Extract text from a file and apply normalization pipeline.
    
    Equivalent to the text extraction logic in add_files_pipeline.py.
    Uses specialized extractors for different file types, with Tika as fallback.
    Applies text normalization and cleaning steps.
    
    Args:
        file_path: Path to the file to extract text from
        
    Returns:
        Normalized text extracted from the file
    """
    # Select appropriate extractor or fallback to Tika
    extractor = get_extractor_for_file(file_path, extractors_list)
    text = extractor(file_path) if extractor else tika_extractor(file_path)
    
    if text:
        # Apply normalization pipeline from add_files_pipeline
        text = common_char_replacements(text)
        text = strip_diacritics(text)
        text = normalize_unicode(text)
        text = normalize_whitespace(text)
    
    return text or ""


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
            "search_target_path": "",
            "distance_metric": "cosine",
        },
    )


@app.post("/", response_class=HTMLResponse)
async def index_post(
    request: Request,
    file: Optional[UploadFile] = File(None),
    search_target_path: str = Form(""),
):
    error = None
    results = []
    query_text = None
    distance_metric = "cosine"

    # Determine path filter from search_target_path
    path_filter = None
    if search_target_path:
        try:
            # extract_server_dirs returns relative path with forward slashes
            # We use include_filename=True because search_target_path is a directory, 
            # and we don't want to strip the last component.
            relative_dir = extract_server_dirs(search_target_path, USER_SERVER_MOUNT_PATH, include_filename=True)
            
            if relative_dir == ".":
                relative_dir = ""
            
            if relative_dir:
                # We want to match anything starting with this directory
                path_filter = relative_dir + "%"
        except ValueError:
             error = f"Search target path must be under {USER_SERVER_MOUNT_PATH}"

    if not error and (not file or not file.filename):
        error = "please choose a file."
    
    if not error:
        temp_dir = tempfile.mkdtemp(prefix="upload_")
        safe_filename = os.path.basename(file.filename)
        tmp_path = os.path.join(temp_dir, safe_filename)

        # save uploaded file to a temp location that preserves its original name
        content = await file.read()
        with open(tmp_path, "wb") as tmp:
            tmp.write(content)

        try:
            # 1) extract text using your existing pipeline
            query_text = extract_and_normalize_text(tmp_path)

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

                        # With a path filter, postgres first narrows the candidate set via file_locations,
                        # then computes vector similarity on that subset.
                        
                        params = {
                            "query_vec": query_vec,
                            "top_k": Config.TOP_K,
                        }

                        if path_filter:
                            params["path_filter"] = path_filter
                            sql = """
                                WITH candidate_files AS (
                                    SELECT DISTINCT file_id 
                                    FROM file_locations 
                                    WHERE file_server_directories LIKE %(path_filter)s
                                )
                                SELECT
                                    fc.file_hash,
                                    fl.file_server_directories,
                                    fl.filename,
                                    (fc.minilm_emb <=> %(query_vec)s) AS distance
                                FROM file_contents fc
                                JOIN files f ON f.hash = fc.file_hash
                                JOIN candidate_files cf ON cf.file_id = f.id
                                LEFT JOIN file_locations fl ON fl.file_id = f.id
                                WHERE fc.minilm_emb IS NOT NULL
                                ORDER BY distance
                                LIMIT %(top_k)s;
                            """
                        else:
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
                                ORDER BY distance
                                LIMIT %(top_k)s;
                            """

                        cur.execute(sql, params)
                        rows = cur.fetchall()

                # 4) post-process rows and build full paths
                for row in rows:
                    directory = row.get("file_server_directories") or ""
                    filename = row.get("filename") or ""
                    distance = row.get("distance")

                    # normalize separators a bit; you can tweak for your env
                    # Use USER_SERVER_MOUNT_PATH as the root
                    pieces = [p for p in [USER_SERVER_MOUNT_PATH, directory, filename] if p]
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
            # clean up temp file and containing directory
            try:
                os.remove(tmp_path)
            except OSError:
                pass

            try:
                os.rmdir(temp_dir)
            except OSError:
                shutil.rmtree(temp_dir, ignore_errors=True)

    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "error": error,
            "results": results,
            "query_text": query_text,
            "search_target_path": search_target_path,
            "distance_metric": distance_metric,
        },
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=5000)
