import os
import dotenv

dotenv.load_dotenv()

class Config:
    # postgres dsn, e.g. "postgresql+psycopg://user:pass@host:port/dbname"
    DATABASE_URL = os.environ.get("DATABASE_URL", "postgresql://user:pass@localhost:5432/archives")

    # number of neighbors to fetch
    TOP_K = int(os.environ.get("TOP_K", "20"))