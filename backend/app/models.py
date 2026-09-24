from datetime import date, datetime
from typing import Literal

from sqlalchemy import JSON, Date, DateTime, ForeignKey, Integer, String
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.db import Base

# Same choices as the Add client screen.
EntityType = Literal["Proprietorship", "Partnership firm", "LLP", "Private limited company"]


class Client(Base):
    __tablename__ = "clients"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(200))
    entity_type: Mapped[str] = mapped_column(String(50))
    financial_year: Mapped[str] = mapped_column(String(10))  # e.g. "2025-26"
    tb_updated_on: Mapped[date | None] = mapped_column(Date, nullable=True)
    review_status: Mapped[str] = mapped_column(String(100), default="Not started")  # e.g. "3 to check", "Finalised"

    workbooks: Mapped[list["ClientWorkbook"]] = relationship(back_populates="client", cascade="all, delete-orphan")


# Where a workbook came from, and where it is in its life.
SOURCES = ("learn", "roll-forward", "build")
STATUSES = ("learned", "draft", "finalised")  # learned: last year's finished workbook as uploaded


class ClientWorkbook(Base):
    """One year's workbook for a client: the saved file (also the template for next year),
    the mapping of every cell to TB ledgers by name, and the review report."""
    __tablename__ = "client_workbooks"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    client_id: Mapped[int] = mapped_column(ForeignKey("clients.id", ondelete="CASCADE"), index=True)
    financial_year: Mapped[str | None] = mapped_column(String(10))
    file_path: Mapped[str] = mapped_column(String(500))  # original upload, kept unchanged
    learned_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.now)
    mapping: Mapped[dict] = mapped_column(JSON)  # LearnResult.mapping(): ledgers, cells, manual inputs, graph
    report: Mapped[dict] = mapped_column(JSON)  # LearnResult.report(): unused ledgers, manual cells, issues
    source: Mapped[str] = mapped_column(String(20), default="learn", server_default="learn")
    status: Mapped[str] = mapped_column(String(20), default="learned", server_default="learned")
    finalised_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    client: Mapped[Client] = relationship(back_populates="workbooks")
