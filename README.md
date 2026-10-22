# BMS SIH 2026 Team Nomination Generator

Quick hosted Flask application that accepts team/participant details and generates a DOCX, PDF and TXT file inside a ZIP.

## Local test

Requires Docker or Python + LibreOffice.

### Docker

```bash
docker build -t sih-generator .
docker run --rm -p 10000:10000 sih-generator
```

Open http://localhost:10000

### Direct Python

Install LibreOffice, then:

```bash
python -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open http://localhost:10000

## Render deployment

1. Put this project in a GitHub repository.
2. In Render, create **New → Web Service** and connect the repository.
3. Set **Language = Docker**.
4. Render will build the Dockerfile and run the app.
5. Use the generated `onrender.com` URL as the student submission page.

No database is required. Generated files are not intentionally stored as submissions; the server returns the ZIP to the student.
