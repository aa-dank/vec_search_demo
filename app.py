import os
import tempfile

from flask import Flask, render_template, request

from config import Config
from db import get_connection

# you will replace these imports with your real api
# i’m guessing the names; adjust to actual functions.
from embedding.minilm import MiniLMEmbedder
from text_extraction.basic_extraction import extract_text_from_path  # TODO: fix name


_embedder = MiniLMEmbedder()


def _embed_text(text: str):
    """Encode the uploaded text with the cached MiniLM embedder."""

    return _embedder.encode([text])[0]


def create_app() -> Flask:
    app = Flask(__name__)
    app.config.from_object(Config)

    @app.route("/", methods=["GET", "POST"])
    def index():
        error = None
        results = []
        query_text = None
        server_root = ""
        distance_metric = "cosine"  # just for display rn

        if request.method == "POST":
            file = request.files.get("file")
            server_root = request.form.get("server_root", "").strip()

            if not file or not file.filename:
                error = "please choose a file."
            else:
                # save uploaded file to a temp location
                with tempfile.NamedTemporaryFile(delete=False) as tmp:
                    file.save(tmp.name)
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
                                        "top_k": app.config["TOP_K"],
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

        return render_template(
            "index.html",
            error=error,
            results=results,
            query_text=query_text,
            server_root=server_root,
            distance_metric=distance_metric,
        )

    return app


if __name__ == "__main__":
    # for local demo only
    app = create_app()
    app.run(debug=True)
