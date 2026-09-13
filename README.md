# SAP LMS Bulk Question Automator — Online Web App

## Run locally
Python 3.10+:
    pip install -r requirements.txt
    python app.py

Open:
    http://127.0.0.1:5000

## Deploy online
This is a Flask web app and can be deployed to a Python-capable web host.
Production command:
    gunicorn app:app

The web UI accepts:
- DOCX
- TXT
- AIKEN
- XLSX SAP template

It generates the SAP bulk Excel workbook using the uploaded template.

## AIKEN format
Question text
A. option
B. option
C. option
D. option
ANSWER: B

Multiple correct:
ANSWER: A,C

Optional:
TYPE: MULTI_CHOICE_MULTIPLE_ANSWER
