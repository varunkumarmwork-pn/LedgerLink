"""Database connection. SQLite for now; set LEDGERLINK_DB_URL to switch (e.g. to PostgreSQL)."""
import os
from pathlib import Path

from fastapi import Depends
from sqlalchemy import create_engine, inspect, text
from sqlalchemy.orm import DeclarativeBase, sessionmaker

DEFAULT_DB_PATH = Path(__file__).resolve().parent.parent / "ledgerlink.db"
DATABASE_URL = os.environ.get("LEDGERLINK_DB_URL", f"sqlite:///{DEFAULT_DB_PATH.as_posix()}")

# check_same_thread is a SQLite-only option needed because FastAPI serves requests on worker threads.
connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(DATABASE_URL, connect_args=connect_args)
SessionLocal = sessionmaker(bind=engine, autoflush=False)


class Base(DeclarativeBase):
    pass


def get_session_factory() -> sessionmaker:
    """FastAPI dependency: makes sessions. Background jobs use it to open their own session
    after the request has finished; tests override it to use a throwaway database."""
    return SessionLocal


def get_session(factory: sessionmaker = Depends(get_session_factory)):
    """FastAPI dependency: one session per request."""
    with factory() as session:
        yield session


def create_schema():
    """Create missing tables, and add columns that later phases introduced to tables that
    already exist (create_all never alters an existing table)."""
    Base.metadata.create_all(engine)
    existing = inspect(engine)
    with engine.begin() as connection:
        for table in Base.metadata.sorted_tables:
            present = {c["name"] for c in existing.get_columns(table.name)}
            for column in table.columns:
                if column.name in present:
                    continue
                kind = column.type.compile(engine.dialect)
                default = f" DEFAULT '{column.server_default.arg}'" if column.server_default is not None else ""
                connection.execute(text(f'ALTER TABLE {table.name} ADD COLUMN "{column.name}" {kind}{default}'))
