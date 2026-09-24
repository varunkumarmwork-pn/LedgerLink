import pytest
from fastapi.testclient import TestClient
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.pool import StaticPool

from app import main
from app.db import Base, get_session_factory
from app.main import app


@pytest.fixture
def client():
    # Fresh in-memory database per test, so the real ledgerlink.db is never touched.
    engine = create_engine("sqlite://", connect_args={"check_same_thread": False}, poolclass=StaticPool)
    Base.metadata.create_all(engine)
    app.dependency_overrides[get_session_factory] = lambda: sessionmaker(bind=engine)
    yield TestClient(app)
    app.dependency_overrides.clear()


@pytest.fixture
def storage(tmp_path, monkeypatch):
    # Uploaded files go to a temporary folder, not backend/storage.
    monkeypatch.setattr(main, "STORAGE_DIR", tmp_path)
    return tmp_path
