@echo off
echo Lancement du Tableau de Bord Carriere (port 8502)...
echo Astuce: si 8502 est deja utilise, fermez les anciennes fenetres Streamlit (ou tuez le process) puis relancez.
call streamlit run "%~dp0app_carriere.py" --server.port 8502 --server.address localhost
pause
