Windows (CMD o PowerShell):

pip install -r requirements.txt
pyinstaller --onefile --windowed --name OmonelDispersion app.py

macOS (Terminal):
pip3 install -r requirements.txt
pyinstaller --onefile --windowed --name OmonelDispersion app.py


El ejecutable quedará en la carpeta dist/ que se genera automáticamente.
No requiere que el usuario tenga Python instalado — es completamente autocontenido.