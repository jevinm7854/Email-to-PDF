### Convert an Email into PDF

Create a virtual environment  
-    python -m venv venv  
-    ./venv/Scripts/activate  

pip install -r requirements.txt    

Customize output folder and path to wkhtmltopdf(for pdfkit) in app.py  

Important files- 
- app.py
- email_process.py
- file_utils.py

Provided logs for checking status and errors

This python application will go to over outlook application on local computer, read the emails and convert those to pdf. 
You can then delete those emails from outlook. This will help save space on outlook and save money