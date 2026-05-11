# Candidate Screener Agent

An AI-powered resume screening agent built with the Groq SDK and Python. Screens multiple resumes against a job description, scores each candidate, saves results to a SQLite database, and generates a ranked shortlist report as a `.docx` file.

## What It Does

1. Reads a job description from `job_description.txt`
2. Loops through all PDF resumes in the `resumes/` folder
3. For each resume — extracts text, evaluates fit via LLM (Llama 4 Scout on Groq)
4. Saves structured results to a local SQLite database
5. Generates a ranked `.docx` report in the `reports/` folder

## Tech Stack

- **LLM**: `meta-llama/llama-4-scout-17b-16e-instruct` via Groq API
- **PDF parsing**: `pdfplumber`
- **Database**: SQLite (via Python `sqlite3`)
- **Report generation**: `python-docx`
- **Framework**: None — pure Groq SDK with a manual agent loop

## Project Structure

```
candidate_screener/
├── screener.py           # Main agent
├── job_description.txt   # Job requirements (edit this)
├── resumes/              # Drop resume PDFs here
├── reports/              # Generated .docx reports saved here
├── screening_results.db  # SQLite DB (auto-created)
├── .env                  # Your API key
└── README.md
```

## Setup

```bash
pip install groq pdfplumber python-docx python-dotenv
```

Create a `.env` file:
```
GROQ_API_KEY=your-key-here
```

## Usage

1. Edit `job_description.txt` with the role requirements
2. Drop candidate resume PDFs into the `resumes/` folder
3. Run:

```bash
python screener.py
```

The agent will print each step as it runs. When complete, open the `.docx` report from `reports/`.

## Output

**Database** (`screening_results.db`) — stores:
- Candidate name
- Fit score (0-10)
- Recommendation (Shortlist / Maybe / Reject)
- Matched skills
- Missing skills
- Summary
- Timestamp

**Report** (`reports/screening_report_TIMESTAMP.docx`) — contains:
- Summary table (total screened, shortlisted, maybe, rejected)
- All candidates ranked by fit score
- Per-candidate: score, recommendation, summary, matched/missing skills

## Scoring Guide

| Score | Recommendation |
|-------|---------------|
| 8-10  | Shortlist     |
| 5-7   | Maybe         |
| 0-4   | Reject        |

## Notes

- Resumes that are image-based PDFs (scanned) cannot be parsed — use text-based PDFs
- Resume text is capped at 6000 characters to stay within LLM token limits
- Each run appends to the same DB — clear `screening_results.db` before a fresh batch