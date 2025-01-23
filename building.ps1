python -m venv venv
.\venv\Scripts\activate
pip install pyinstaller==6.11.1
pip install -r .\requirements.txt
pyinstaller.exe -F -i .\icon.ico --add-data '.\icon.ico;.' -w .\quality-report-harvester.py