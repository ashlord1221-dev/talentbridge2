# TalentBridge — Smart Hiring Platform

AI-powered job portal. Companies post jobs, candidates apply, AI ranks resumes.

## Quick Start (Local)

```bash
pip install flask werkzeug pdfminer.six python-docx
python app.py
# Open http://localhost:5000
```

## Demo Accounts

| Role | Email | Password |
|------|-------|----------|
| Candidate | arjun@demo.com | demo123 |
| Candidate | priya@demo.com | demo123 |
| Company | google@demo.com | demo123 |
| Company | zomato@demo.com | demo123 |
| Company | infosys@demo.com | demo123 |

## Deploy

- **Vercel**: `vercel --prod` (uses `vercel.json` + `api/index.py`)
- **Render**: uses `Procfile` + `render.yaml`

## Tech Stack

Flask · SQLite · Werkzeug · pdfminer.six · python-docx · Vanilla JS
