

Step 1: Install Python and create virtual environment

python -m venv venv

venv\\Scripts\\activate



Step 2: Install Python dependencies

pip install -r requirements.txt



PS C:\\SmartCatalog\\SmartCatalog> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

PS C:\\SmartCatalog\\SmartCatalog> .\\venv\\Scripts\\Activate.ps1



To generate into .exe application
Install auto py to exe in the virtual environment
pip install auto-py-to-exe

Launch it in the virtual environment
auto-py-to-exe.exe

For the configuration
1) Set Script Location to:
run.py

2) Add search path
Advanced → --paths → add:
d:\dev\projects\SmartCatalog_V2\src

3) Hidden imports
Keep: fitz, PyMuPDF

4) Collect all
Advanced → --collect-all → add:
pymupdf