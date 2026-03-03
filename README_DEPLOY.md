# Mise en ligne (Streamlit) — gratuit + persistant

Ton application Streamlit **stocke des comptes** et des **fichiers d’extractions** (`data/users.sqlite3`, `data/projects.sqlite3`, `data/project_files/...`).
Pour que ça soit **vraiment persistant** en gratuit, il faut un hébergement avec **disque persistant** (sinon tout peut disparaître au redémarrage).

## Option recommandée (gratuit + persistant) : VPS Oracle Cloud “Always Free”

### 1) Préparer le dépôt (GitHub)
1. Crée un repo GitHub (privé recommandé).
2. À la racine du projet :
   - `requirements.txt` est prêt.
   - `.gitignore` ignore `data/` et les fichiers d’extraction.
3. **Ne commit pas** les extractions réelles.
   - Sur le serveur, tu copieras manuellement un fichier par défaut si tu veux.

### 2) Créer le VPS
1. Crée un compte Oracle Cloud (Always Free) et une VM Ubuntu.
2. Ouvre le firewall Oracle + l’OS pour `80` et `443`.

### 3) Installation sur le VPS (Ubuntu)
```bash
sudo apt update
sudo apt install -y python3-venv python3-pip nginx

sudo mkdir -p /opt/carriere-dashboard
sudo chown -R $USER:$USER /opt/carriere-dashboard
cd /opt/carriere-dashboard

git clone <TON_REPO_GIT> .
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 4) Lancer Streamlit en service (systemd)
Crée `/etc/systemd/system/carriere-dashboard.service` :
```ini
[Unit]
Description=Carriere Dashboard (Streamlit)
After=network.target

[Service]
User=ubuntu
WorkingDirectory=/opt/carriere-dashboard
Environment=STREAMLIT_SERVER_HEADLESS=true
Environment=STREAMLIT_BROWSER_GATHERUSAGESTATS=false
ExecStart=/opt/carriere-dashboard/.venv/bin/streamlit run app_carriere.py --server.address 127.0.0.1 --server.port 8501
Restart=always
RestartSec=3

[Install]
WantedBy=multi-user.target
```
Puis :
```bash
sudo systemctl daemon-reload
sudo systemctl enable --now carriere-dashboard
sudo systemctl status carriere-dashboard --no-pager
```

### 5) Reverse proxy Nginx (HTTPS)
Crée `/etc/nginx/sites-available/carriere-dashboard` :
```nginx
server {
  listen 80;
  server_name TON_DOMAINE_OU_IP;

  location / {
    proxy_pass http://127.0.0.1:8501;
    proxy_http_version 1.1;
    proxy_set_header Upgrade $http_upgrade;
    proxy_set_header Connection "upgrade";
    proxy_set_header Host $host;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto $scheme;
  }
}
```
Active :
```bash
sudo ln -sf /etc/nginx/sites-available/carriere-dashboard /etc/nginx/sites-enabled/carriere-dashboard
sudo nginx -t
sudo systemctl reload nginx
```
Ensuite, ajoute HTTPS (Let’s Encrypt) si tu as un domaine :
```bash
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d TON_DOMAINE
```

### 6) Persistance des données
Sur ce VPS :
- `data/` reste sur disque → **comptes + extractions persistants**.
- Pense à faire des sauvegardes (snapshot Oracle, ou export régulier).

## Option “zéro serveur” (Streamlit Community Cloud)
⚠️ Pas recommandé si tu veux **persistant** sans développement supplémentaire :
le disque peut être réinitialisé, donc `data/` et les uploads peuvent être perdus.
Pour du persistant sur Streamlit Cloud, il faut migrer vers une DB/Storage externe (Supabase/Neon + stockage fichiers).

