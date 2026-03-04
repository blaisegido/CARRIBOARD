import base64
import hashlib
import hmac
import json
import re
import secrets
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional


PBKDF2_ITERATIONS = 210_000
SALT_BYTES = 16


class AuthError(Exception):
    pass


@dataclass(frozen=True)
class User:
    id: int
    username: str
    role: str


def _now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _normalize_username(username: str) -> tuple[str, str]:
    u = (username or "").strip()
    if not u:
        raise AuthError("Nom d'utilisateur requis.")
    if len(u) < 3 or len(u) > 50:
        raise AuthError("Le nom d'utilisateur doit faire entre 3 et 50 caractères.")
    # Autorise lettres/chiffres/espaces + quelques séparateurs.
    if not re.fullmatch(r"[A-Za-z0-9 _.\-]+", u):
        raise AuthError("Caractères invalides dans le nom d'utilisateur.")
    return u, u.casefold()


def _hash_password(password: str, salt: bytes) -> bytes:
    pw = (password or "").encode("utf-8")
    return hashlib.pbkdf2_hmac("sha256", pw, salt, PBKDF2_ITERATIONS)


def init_db(db_path: Path) -> None:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(db_path) as con:
        con.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              username TEXT NOT NULL,
              username_canon TEXT NOT NULL UNIQUE,
              role TEXT NOT NULL DEFAULT 'user',
              password_salt BLOB NOT NULL,
              password_hash BLOB NOT NULL,
              created_at TEXT NOT NULL,
              last_login TEXT
            )
            """
        )
        con.execute("CREATE INDEX IF NOT EXISTS idx_users_username_canon ON users(username_canon)")


def get_user(db_path: Path, user_id: int) -> Optional[User]:
    init_db(db_path)
    with sqlite3.connect(db_path) as con:
        row = con.execute(
            "SELECT id, username, role FROM users WHERE id = ?",
            (int(user_id),),
        ).fetchone()
        if not row:
            return None
        uid, username, role = row
        return User(id=int(uid), username=str(username), role=str(role))


def _b64u_encode(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def _b64u_decode(data: str) -> bytes:
    s = (data or "").strip()
    pad = "=" * ((4 - (len(s) % 4)) % 4)
    return base64.urlsafe_b64decode(s + pad)


def issue_session_token(user_id: int, secret: bytes, *, max_age_seconds: int) -> str:
    if not secret:
        raise ValueError("secret is required")
    now = int(datetime.now(timezone.utc).timestamp())
    payload = {"uid": int(user_id), "iat": now, "exp": now + int(max_age_seconds)}
    payload_raw = json.dumps(payload, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    sig = hmac.new(secret, payload_raw, hashlib.sha256).digest()
    return f"v1.{_b64u_encode(payload_raw)}.{_b64u_encode(sig)}"


def verify_session_token(token: str, secret: bytes) -> Optional[int]:
    if not token or not secret:
        return None
    parts = str(token).split(".")
    if len(parts) != 3 or parts[0] != "v1":
        return None
    try:
        payload_raw = _b64u_decode(parts[1])
        sig = _b64u_decode(parts[2])
    except Exception:
        return None

    expected = hmac.new(secret, payload_raw, hashlib.sha256).digest()
    if not hmac.compare_digest(expected, sig):
        return None

    try:
        payload = json.loads(payload_raw.decode("utf-8"))
        uid = int(payload.get("uid"))
        exp = int(payload.get("exp"))
    except Exception:
        return None

    now = int(datetime.now(timezone.utc).timestamp())
    if exp < now:
        return None
    return uid


def user_count(db_path: Path) -> int:
    with sqlite3.connect(db_path) as con:
        row = con.execute("SELECT COUNT(*) FROM users").fetchone()
        return int(row[0] if row else 0)


def create_user(db_path: Path, username: str, password: str) -> User:
    username_clean, username_canon = _normalize_username(username)
    if not password or len(password) < 8:
        raise AuthError("Le mot de passe doit faire au moins 8 caractères.")

    init_db(db_path)
    salt = secrets.token_bytes(SALT_BYTES)
    pw_hash = _hash_password(password, salt)

    role = "admin" if user_count(db_path) == 0 else "user"

    try:
        with sqlite3.connect(db_path) as con:
            cur = con.execute(
                """
                INSERT INTO users (username, username_canon, role, password_salt, password_hash, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (username_clean, username_canon, role, salt, pw_hash, _now_iso()),
            )
            user_id = int(cur.lastrowid)
    except sqlite3.IntegrityError:
        raise AuthError("Ce nom d'utilisateur existe déjà.")

    return User(id=user_id, username=username_clean, role=role)


def authenticate(db_path: Path, username: str, password: str) -> User:
    _, username_canon = _normalize_username(username)
    init_db(db_path)

    with sqlite3.connect(db_path) as con:
        row = con.execute(
            """
            SELECT id, username, role, password_salt, password_hash
            FROM users
            WHERE username_canon = ?
            """,
            (username_canon,),
        ).fetchone()

        if not row:
            raise AuthError("Identifiants incorrects.")

        user_id, username_db, role, salt, pw_hash = row
        calc = _hash_password(password, salt)
        if not hmac.compare_digest(calc, pw_hash):
            raise AuthError("Identifiants incorrects.")

        con.execute("UPDATE users SET last_login = ? WHERE id = ?", (_now_iso(), int(user_id)))

    return User(id=int(user_id), username=str(username_db), role=str(role))

