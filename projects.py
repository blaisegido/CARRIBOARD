import re
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional


class ProjectError(Exception):
    pass


@dataclass(frozen=True)
class Project:
    id: str
    user_id: int
    name: str
    name_canon: str
    created_at: str
    data_path: str
    source_filename: Optional[str]
    date_min: Optional[str]
    date_max: Optional[str]
    nb_livraisons: Optional[int]
    tonnage_total: Optional[float]
    ca_total: Optional[float]
    theme_idx: int
    order_idx: int


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _normalize_name(name: str) -> tuple[str, str]:
    cleaned = re.sub(r"\s+", " ", (name or "").strip())
    if not cleaned:
        raise ProjectError("Nom de dossier requis.")
    if len(cleaned) > 80:
        cleaned = cleaned[:80].rstrip()
    return cleaned, cleaned.casefold()


def init_db(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as con:
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS projects (
              id TEXT PRIMARY KEY,
              user_id INTEGER NOT NULL,
              name TEXT NOT NULL,
              name_canon TEXT NOT NULL,
              created_at TEXT NOT NULL,
              data_path TEXT NOT NULL,
              source_filename TEXT,
              date_min TEXT,
              date_max TEXT,
              nb_livraisons INTEGER,
              tonnage_total REAL,
              ca_total REAL,
              theme_idx INTEGER NOT NULL DEFAULT 0,
              order_idx INTEGER NOT NULL DEFAULT 0
            )
            """
        )
        con.execute("CREATE INDEX IF NOT EXISTS idx_projects_user_id ON projects(user_id)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_projects_user_order ON projects(user_id, order_idx)")
        con.execute("CREATE INDEX IF NOT EXISTS idx_projects_user_name ON projects(user_id, name_canon)")


def _connect(db_path: Path) -> sqlite3.Connection:
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    return con


def _project_from_row(row: sqlite3.Row) -> Project:
    data = dict(row)
    fields = set(Project.__dataclass_fields__.keys())
    filtered = {k: v for k, v in data.items() if k in fields}
    # Compatibilité ascendante/descendante : autorise l'absence de champs optionnels.
    if "name_canon" in fields and "name_canon" not in filtered:
        filtered["name_canon"] = str(filtered.get("name", "")).casefold()
    return Project(**filtered)


def list_projects(db_path: Path, user_id: int) -> list[Project]:
    init_db(db_path)
    with _connect(db_path) as con:
        rows = con.execute(
            """
            SELECT *
            FROM projects
            WHERE user_id = ?
            ORDER BY order_idx ASC, created_at DESC
            """,
            (int(user_id),),
        ).fetchall()
    return [_project_from_row(r) for r in rows]


def get_project(db_path: Path, user_id: int, project_id: str) -> Optional[Project]:
    init_db(db_path)
    with _connect(db_path) as con:
        row = con.execute(
            """
            SELECT *
            FROM projects
            WHERE user_id = ? AND id = ?
            """,
            (int(user_id), str(project_id)),
        ).fetchone()
    return _project_from_row(row) if row else None


def _unique_name(con: sqlite3.Connection, user_id: int, name: str, *, exclude_id: Optional[str] = None) -> str:
    base, canon = _normalize_name(name)
    suffix = 1
    candidate = base
    candidate_canon = canon
    while True:
        if exclude_id:
            row = con.execute(
                "SELECT 1 FROM projects WHERE user_id = ? AND name_canon = ? AND id <> ? LIMIT 1",
                (int(user_id), candidate_canon, str(exclude_id)),
            ).fetchone()
        else:
            row = con.execute(
                "SELECT 1 FROM projects WHERE user_id = ? AND name_canon = ? LIMIT 1",
                (int(user_id), candidate_canon),
            ).fetchone()
        if not row:
            return candidate
        suffix += 1
        candidate = f"{base} ({suffix})"
        candidate_canon = candidate.casefold()


def create_project(
    db_path: Path,
    *,
    project_id: str,
    user_id: int,
    name: str,
    data_path: str,
    source_filename: Optional[str] = None,
    date_min: Optional[str] = None,
    date_max: Optional[str] = None,
    nb_livraisons: Optional[int] = None,
    tonnage_total: Optional[float] = None,
    ca_total: Optional[float] = None,
    theme_idx: int = 0,
) -> str:
    init_db(db_path)
    with _connect(db_path) as con:
        next_order = con.execute(
            "SELECT COALESCE(MAX(order_idx), -1) + 1 FROM projects WHERE user_id = ?",
            (int(user_id),),
        ).fetchone()[0]

        unique = _unique_name(con, user_id, name)
        con.execute(
            """
            INSERT INTO projects (
              id, user_id, name, name_canon, created_at, data_path, source_filename,
              date_min, date_max, nb_livraisons, tonnage_total, ca_total, theme_idx, order_idx
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                str(project_id),
                int(user_id),
                unique,
                unique.casefold(),
                _now_iso(),
                str(data_path),
                source_filename,
                date_min,
                date_max,
                nb_livraisons,
                tonnage_total,
                ca_total,
                int(theme_idx),
                int(next_order),
            ),
        )
    return str(project_id)


def rename_project(db_path: Path, *, user_id: int, project_id: str, new_name: str) -> str:
    init_db(db_path)
    with _connect(db_path) as con:
        existing = con.execute(
            "SELECT 1 FROM projects WHERE user_id = ? AND id = ?",
            (int(user_id), str(project_id)),
        ).fetchone()
        if not existing:
            raise ProjectError("Dossier introuvable.")

        unique = _unique_name(con, user_id, new_name, exclude_id=str(project_id))
        con.execute(
            "UPDATE projects SET name = ?, name_canon = ? WHERE user_id = ? AND id = ?",
            (unique, unique.casefold(), int(user_id), str(project_id)),
        )
    return unique


def update_project_data(
    db_path: Path,
    *,
    user_id: int,
    project_id: str,
    data_path: str,
    date_min: Optional[str] = None,
    date_max: Optional[str] = None,
    nb_livraisons: Optional[int] = None,
    tonnage_total: Optional[float] = None,
    ca_total: Optional[float] = None,
) -> None:
    init_db(db_path)
    with _connect(db_path) as con:
        con.execute(
            """
            UPDATE projects
            SET data_path = ?,
                date_min = ?,
                date_max = ?,
                nb_livraisons = ?,
                tonnage_total = ?,
                ca_total = ?
            WHERE user_id = ? AND id = ?
            """,
            (
                str(data_path),
                date_min,
                date_max,
                nb_livraisons,
                tonnage_total,
                ca_total,
                int(user_id),
                str(project_id),
            ),
        )


def delete_project(db_path: Path, *, user_id: int, project_id: str) -> None:
    init_db(db_path)
    with _connect(db_path) as con:
        cur = con.execute(
            "DELETE FROM projects WHERE user_id = ? AND id = ?",
            (int(user_id), str(project_id)),
        )
        if cur.rowcount == 0:
            raise ProjectError("Dossier introuvable.")
