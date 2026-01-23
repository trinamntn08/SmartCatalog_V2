

Step 1: Install Python and create virtual environment

python -m venv venv

venv\\Scripts\\activate



Step 2: Install Python dependencies

pip install -r requirements.txt



PS C:\\SmartCatalog\\SmartCatalog> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

PS C:\\SmartCatalog\\SmartCatalog> .\\venv\\Scripts\\Activate.ps1





Step 3: Install system dependencies

\- Tesseract: https://github.com/UB-Mannheim/tesseract/wiki

\- Poppler: https://github.com/oschwartz10612/poppler-windows/releases

(→ Add both to system PATH)



