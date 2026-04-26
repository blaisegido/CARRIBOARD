from typing import Optional
from supabase import Client

BUCKET_NAME = "projects"

def ensure_bucket_exists(client: Client) -> None:
    try:
        buckets = client.storage.list_buckets()
        bucket_names = [b.name for b in buckets]
        if BUCKET_NAME not in bucket_names:
            print(f"Tentative de création du bucket '{BUCKET_NAME}'...")
            client.storage.create_bucket(BUCKET_NAME, options={"public": False})
    except Exception as e:
        print(f"Info: Impossible de vérifier/créer le bucket '{BUCKET_NAME}' ({e}). Assurez-vous qu'il existe dans Supabase.")

def upload_file(client: Client, user_id: int, project_id: str, remote_path: str, data: bytes) -> bool:
    """Uploads a file to Supabase Storage."""
    ensure_bucket_exists(client)
    full_storage_path = f"user_{user_id}/{project_id}/{remote_path}"
    try:
        client.storage.from_(BUCKET_NAME).upload(
            path=full_storage_path,
            file=data,
            file_options={"upsert": True}
        )
        return True
    except Exception as e:
        # Tentative d'update si l'upload échoue (certaines versions de l'API se comportent différemment)
        try:
            client.storage.from_(BUCKET_NAME).update(
                path=full_storage_path,
                file=data,
                file_options={"upsert": True}
            )
            return True
        except Exception as e2:
            print(f"Erreur Supabase Upload/Update ({full_storage_path}): {e2}")
            return False

def download_file(client: Client, user_id: int, project_id: str, remote_path: str) -> Optional[bytes]:
    """Downloads a file from Supabase Storage."""
    try:
        full_storage_path = f"user_{user_id}/{project_id}/{remote_path}"
        res = client.storage.from_(BUCKET_NAME).download(full_storage_path)
        return res
    except Exception as e:
        print(f"Erreur Supabase Download ({full_storage_path}): {e}")
        return None

def delete_project_files(client: Client, user_id: int, project_id: str) -> bool:
    """Deletes all files associated with a project in Supabase Storage."""
    try:
        prefix = f"user_{user_id}/{project_id}/"
        files = client.storage.from_(BUCKET_NAME).list(f"user_{user_id}/{project_id}")
        if not files:
            return True
            
        file_paths = [f"{prefix}{f['name']}" for f in files]
        client.storage.from_(BUCKET_NAME).remove(file_paths)
        return True
    except Exception as e:
        print(f"Erreur Supabase Delete ({prefix}): {e}")
        return False
