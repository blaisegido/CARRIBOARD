@echo off
echo Lancement du Tableau de Bord Carriere...
echo Si plusieurs onglets localhost (8501/8502) apparaissent, fermez les anciennes fenetres Streamlit puis relancez.
call streamlit run "%~dp0app_carriere.py" --server.port 8501 --server.address localhost
pause
