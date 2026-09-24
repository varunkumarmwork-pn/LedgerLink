"""Background jobs (learning a workbook) and their progress, kept in memory.

The job runs on a worker thread and writes progress here; the Processing screen reads it
through /api/jobs/{id}/events. Jobs are forgotten when the server restarts.
"""
import threading
import uuid

from app.learn import Progress

STEP_COUNT = 4
LEARN_STEPS = ["Read the trial balance", "Find the sheets", "Trace every link back to the trial balance", "Check totals"]
ROLL_FORWARD_STEPS = ["Read the new trial balance", "Match ledgers to last year", "Rewrite formulas and carry forward",
                      "Check totals"]
BUILD_STEPS = ["Read the trial balance", "Put ledgers on statement lines", "Write the ticked sheets", "Check totals"]
LOG_LIMIT = 500  # log lines kept per job


class Job(Progress):
    def __init__(self, client_id: int, client_name: str, filename: str, action: str = "Reading",
                 titles: list[str] = LEARN_STEPS):
        self.id = uuid.uuid4().hex
        self.action = action  # heading on the Processing screen: "Reading Venus Agencies"
        self.titles = titles
        self.client_id = client_id
        self.client_name = client_name
        self.filename = filename
        self.status = "running"  # running | done | failed
        self.step = 0  # index of the step being worked on; STEP_COUNT when all are done
        self.details = [""] * STEP_COUNT  # what each finished step found
        self.log: list[list[str]] = []  # [cell, formula, ledger names]
        self.error: str | None = None
        self.workbook_id: int | None = None
        self.version = 0  # bumped on every change so watchers know when to send an update
        self._lock = threading.Lock()

    # Progress callbacks from the learn engine
    def step_done(self, index: int, detail: str):
        with self._lock:
            self.details[index] = detail
            self.step = index + 1
            self.version += 1

    def traced(self, cell: str, formula: str, ledgers: str):
        with self._lock:
            self.log = (self.log + [[cell, formula, ledgers]])[-LOG_LIMIT:]
            self.version += 1

    def finish(self, workbook_id: int):
        with self._lock:
            self.status, self.workbook_id, self.step = "done", workbook_id, STEP_COUNT
            self.version += 1

    def fail(self, message: str):
        with self._lock:
            self.status, self.error = "failed", message
            self.version += 1

    def snapshot(self) -> dict:
        with self._lock:
            return {
                "id": self.id, "client_id": self.client_id, "client_name": self.client_name,
                "action": self.action, "titles": list(self.titles),
                "filename": self.filename, "status": self.status, "step": self.step,
                "details": list(self.details), "log": list(self.log), "error": self.error,
                "workbook_id": self.workbook_id, "version": self.version,
            }


_jobs: dict[str, Job] = {}


def create(client_id: int, client_name: str, filename: str, action: str = "Reading",
           titles: list[str] = LEARN_STEPS) -> Job:
    job = Job(client_id, client_name, filename, action, titles)
    _jobs[job.id] = job
    return job


def get(job_id: str) -> Job | None:
    return _jobs.get(job_id)
