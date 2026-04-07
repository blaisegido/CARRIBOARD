import re
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional, List
from supabase import Client

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

def _normalize_path_for_storage(path_str: str) -> str:
    # Convertit les backslashes en slashs pour la compatibilité Windows/Linux en base de données.
    if not path_str:
        return ""
    return str(path_str).replace("\\", "/")

def _project_from_row(row: dict) -> Project:
    raw_path = str(row["data_path"])
    
    # Tentative de migration automatique des anciens chemins absolus
    # On ne garde que la partie après "project_files" si présente.
    if "project_files" in raw_path:
        try:
            # Sépare par project_files (cas Windows ou Linux)
            parts = re.split(r"[\\/]+project_files[\\/]+", raw_path)
            if len(parts) > 1:
                # Reconstruit un chemin relatif propre
                inner = parts[1].replace("\\", "/").strip("/")
                raw_path = f"project_files/{inner}"
        except Exception:
            pass
            
    return Project(
        id=str(row["id"]),
        user_id=int(row["user_id"]),
        name=str(row["name"]),
        name_canon=str(row.get("name_canon", row["name"].casefold())),
        created_at=str(row["created_at"]),
        data_path=_normalize_path_for_storage(raw_path),
        source_filename=row.get("source_filename"),
        date_min=row.get("date_min"),
        date_max=row.get("date_max"),
        nb_livraisons=row.get("nb_livraisons"),
        tonnage_total=row.get("tonnage_total"),
        ca_total=row.get("ca_total"),
        theme_idx=int(row.get("theme_idx", 0)),
        order_idx=int(row.get("order_idx", 0))
    )

def list_projects(client: Client, user_id: int) -> List[Project]:
    try:
        res = client.table("projects").select("*").eq("user_id", user_id).order("order_idx").order("created_at", desc=True).execute()
        return [_project_from_row(r) for r in res.data]
    except:
        return []

def get_project(client: Client, user_id: int, project_id: str) -> Optional[Project]:
    try:
        res = client.table("projects").select("*").eq("user_id", user_id).eq("id", project_id).execute()
        if not res.data:
            return None
        return _project_from_row(res.data[0])
    except:
        return None

def _unique_name(client: Client, user_id: int, name: str, *, exclude_id: Optional[str] = None) -> str:
    base, canon = _normalize_name(name)
    suffix = 1
    candidate = base
    candidate_canon = canon
    while True:
        query = client.table("projects").select("id").eq("user_id", user_id).eq("name_canon", candidate_canon)
        if exclude_id:
            query = query.neq("id", exclude_id)
        res = query.execute()
        if not res.data:
            return candidate
        suffix += 1
        candidate = f"{base} ({suffix})"
        candidate_canon = candidate.casefold()

def create_project(
    client: Client,
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
    try:
        # Get next order_idx
        res_order = client.table("projects").select("order_idx").eq("user_id", user_id).order("order_idx", desc=True).limit(1).execute()
        next_order = 0
        if res_order.data:
            next_order = int(res_order.data[0]["order_idx"]) + 1
            
        unique = _unique_name(client, user_id, name)
        
        client.table("projects").insert({
            "id": str(project_id),
            "user_id": user_id,
            "name": unique,
            "name_canon": unique.casefold(),
            "created_at": _now_iso(),
            "data_path": _normalize_path_for_storage(data_path),
            "source_filename": source_filename,
            "date_min": date_min,
            "date_max": date_max,
            "nb_livraisons": nb_livraisons,
            "tonnage_total": tonnage_total,
            "ca_total": ca_total,
            "theme_idx": int(theme_idx),
            "order_idx": next_order
        }).execute()
        return project_id
    except Exception as e:
        raise ProjectError(f"Erreur lors de la création du projet : {str(e)}")

def rename_project(client: Client, *, user_id: int, project_id: str, new_name: str) -> str:
    try:
        unique = _unique_name(client, user_id, new_name, exclude_id=project_id)
        res = client.table("projects").update({
            "name": unique,
            "name_canon": unique.casefold()
        }).eq("user_id", user_id).eq("id", project_id).execute()
        
        if not res.data:
             raise ProjectError("Dossier introuvable.")
        return unique
    except Exception as e:
        raise ProjectError(str(e))

def update_project_data(
    client: Client,
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
    try:
        client.table("projects").update({
            "data_path": _normalize_path_for_storage(data_path),
            "date_min": date_min,
            "date_max": date_max,
            "nb_livraisons": nb_livraisons,
            "tonnage_total": tonnage_total,
            "ca_total": ca_total
        }).eq("user_id", user_id).eq("id", project_id).execute()
    except Exception as e:
        pass

def delete_project(client: Client, *, user_id: int, project_id: str) -> None:
    try:
        res = client.table("projects").delete().eq("user_id", user_id).eq("id", project_id).execute()
        if not res.data:
            raise ProjectError("Dossier introuvable.")
    except Exception as e:
        raise ProjectError(str(e))
