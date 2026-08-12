cd "C:\Users\jarias\Desktop\tractocar-ventas"
python procesar.py
git add plantilla.html procesar.py .gitignore
git commit -m "Actualizacion automatica $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
git push origin main
