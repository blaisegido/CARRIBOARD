import os
import streamlit as st
from supabase import create_client, Client

def test_connection():
    # Use the credentials provided in the context for testing
    url = "https://wbiashsnyftvjxbdqznh.supabase.co"
    key = "sb_publishable_HWr5LyhovG74L_iR7uuBnQ_F6BNDgLy"
    
    try:
        supabase: Client = create_client(url, key)
        # Try to list users table structure (or just ping)
        response = supabase.table("users").select("count", count="exact").limit(1).execute()
        print(f"Connexion réussie ! Table 'users' accessible.")
    except Exception as e:
        print(f"Erreur de connexion : {e}")

if __name__ == '__main__':
    test_connection()
