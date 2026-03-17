import base64
import hashlib
import hmac
import json
import re
import secrets
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional, Any
from supabase import Client

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

def _token_hash(token: str) -> str:
    t = (token or "").encode("utf-8")
    return hashlib.sha256(t).hexdigest()

def _try_parse_token_payload(token: str) -> Optional[dict]:
    parts = str(token or "").split(".")
    if len(parts) != 3 or parts[0] != "v1":
        return None
    try:
        payload_raw = _b64u_decode(parts[1])
        payload = json.loads(payload_raw.decode("utf-8"))
        if not isinstance(payload, dict):
            return None
        return payload
    except Exception:
        return None

def _normalize_username(username: str) -> tuple[str, str]:
    u = (username or "").strip()
    if not u:
        raise AuthError("Nom d'utilisateur requis.")
    if len(u) < 3 or len(u) > 50:
        raise AuthError("Le nom d'utilisateur doit faire entre 3 et 50 caractères.")
    if not re.fullmatch(r"[A-Za-z0-9 _.\-]+", u):
        raise AuthError("Caractères invalides dans le nom d'utilisateur.")
    return u, u.casefold()

def _hash_password(password: str, salt: bytes) -> bytes:
    pw = (password or "").encode("utf-8")
    return hashlib.pbkdf2_hmac("sha256", pw, salt, PBKDF2_ITERATIONS)

def _b64u_encode(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")

def _b64u_decode(data: str) -> bytes:
    s = (data or "").strip()
    pad = "=" * ((4 - (len(s) % 4)) % 4)
    return base64.urlsafe_b64decode(s + pad)

def get_user(client: Client, user_id: int) -> Optional[User]:
    try:
        res = client.table("users").select("id, username, role").eq("id", user_id).execute()
        if not res.data:
            return None
        row = res.data[0]
        return User(id=int(row["id"]), username=str(row["username"]), role=str(row["role"]))
    except:
        return None

def issue_session_token(user_id: int, secret: bytes, *, max_age_seconds: int) -> str:
    if not secret:
        raise ValueError("secret is required")
    now = int(datetime.now(timezone.utc).timestamp())
    payload = {"uid": int(user_id), "iat": now, "exp": now + int(max_age_seconds)}
    payload_raw = json.dumps(payload, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    sig = hmac.new(secret, payload_raw, hashlib.sha256).digest()
    return f"v1.{_b64u_encode(payload_raw)}.{_b64u_encode(sig)}"

def revoke_session_token(client: Client, token: str) -> None:
    token_hash = _token_hash(token)
    payload = _try_parse_token_payload(token)
    exp = None
    if payload and payload.get("exp") is not None:
        try:
            exp = int(payload.get("exp"))
        except:
            pass
    try:
        client.table("revoked_tokens").upsert({
            "token_hash": token_hash,
            "revoked_at": _now_iso(),
            "exp": exp
        }).execute()
    except:
        pass

def verify_session_token(token: str, secret: bytes, client: Optional[Client] = None) -> Optional[int]:
    if not token or not secret:
        return None
    parts = str(token).split(".")
    if len(parts) != 3 or parts[0] != "v1":
        return None
    try:
        payload_raw = _b64u_decode(parts[1])
        sig = _b64u_decode(parts[2])
    except:
        return None

    expected = hmac.new(secret, payload_raw, hashlib.sha256).digest()
    if not hmac.compare_digest(expected, sig):
        return None

    try:
        payload = json.loads(payload_raw.decode("utf-8"))
        uid = int(payload.get("uid"))
        exp = int(payload.get("exp"))
    except:
        return None

    now = int(datetime.now(timezone.utc).timestamp())
    if exp < now:
        return None

    if client is not None:
        try:
            res = client.table("revoked_tokens").select("token_hash").eq("token_hash", _token_hash(token)).execute()
            if res.data:
                return None
        except:
            pass
    return uid

def user_count(client: Client) -> int:
    try:
        res = client.table("users").select("id", count="exact").execute()
        return res.count or 0
    except:
        return 0

def create_user(client: Client, username: str, password: str) -> User:
    username_clean, username_canon = _normalize_username(username)
    if not password or len(password) < 8:
        raise AuthError("Le mot de passe doit faire au moins 8 caractères.")

    salt = secrets.token_bytes(SALT_BYTES)
    pw_hash = _hash_password(password, salt)

    role = "admin" if user_count(client) == 0 else "user"

    try:
        res = client.table("users").insert({
            "username": username_clean,
            "username_canon": username_canon,
            "role": role,
            "password_salt": base64.b64encode(salt).decode("ascii"),
            "password_hash": base64.b64encode(pw_hash).decode("ascii"),
            "created_at": _now_iso()
        }).execute()
        
        if not res.data:
            raise AuthError("Erreur lors de la création de l'utilisateur.")
        
        row = res.data[0]
        return User(id=int(row["id"]), username=username_clean, role=role)
    except Exception as e:
        if "unique" in str(e).lower() or "already exists" in str(e).lower():
            raise AuthError("Ce nom d'utilisateur existe déjà.")
        raise AuthError(f"Erreur technique : {str(e)}")

def authenticate(client: Client, username: str, password: str) -> User:
    _, username_canon = _normalize_username(username)
    try:
        res = client.table("users").select("*").eq("username_canon", username_canon).execute()
        if not res.data:
            raise AuthError("Identifiants incorrects.")

        row = res.data[0]
        salt = base64.b64decode(row["password_salt"])
        pw_hash = base64.b64decode(row["password_hash"])
        
        calc = _hash_password(password, salt)
        if not hmac.compare_digest(calc, pw_hash):
            raise AuthError("Identifiants incorrects.")

        client.table("users").update({"last_login": _now_iso()}).eq("id", row["id"]).execute()
        return User(id=int(row["id"]), username=str(row["username"]), role=str(row["role"]))
    except AuthError:
        raise
    except Exception as e:
        raise AuthError("Erreur lors de la connexion.")
