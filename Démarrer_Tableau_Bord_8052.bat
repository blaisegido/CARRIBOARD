@echo off
echo Lancement du Tableau de Bord Carriere (port 8052)...
echo Ouvrez ensuite http://127.0.0.1:8052/ (evite les conflits de cookies avec localhost).
call streamlit run "%~dp0app_carriere.py" --server.port 8052 --server.address 127.0.0.1
pause
