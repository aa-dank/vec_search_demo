import psycopg
from psycopg.rows import dict_row
from pgvector.psycopg import register_vector  # pip install pgvector

from config import Config

def get_connection():
    # simple one-shot connection; for real app you'd pool or stick on flask.g
    conn = psycopg.connect(
        Config.DATABASE_URL,
        row_factory=dict_row,
    )
    register_vector(conn)
    return conn