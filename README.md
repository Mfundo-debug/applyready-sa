# ApplyReady SA

ApplyReady SA is a Streamlit application readiness tool for South African job seekers, students, graduates, first-time applicants, internship applicants, learnership applicants, bursary applicants and graduate programme applicants.

It helps applicants check whether their application pack is complete, professional and aligned before they submit.

## Version 2: Application Fit Engine

Version 2 adds a stronger opportunity-matching layer on top of the original readiness checker.

### Core features

- Opportunity details form
- Advert input by paste, document upload or public link
- Applicant profile form
- Formal-experience / first-time-applicant mode
- Document checklist based on opportunity type
- Supporting-document upload with file-name type detection
- CV upload and text extraction from PDF, DOCX or TXT
- CV readiness checks
- Smart advert extraction
- CV-to-advert opportunity fit score
- Skills found, matched and missing report
- Keyword alignment report
- Fit recommendations
- First-time applicant CV guidance
- Readiness score out of 100
- Priority fixes
- Tailored motivation-letter draft
- Professional email draft
- Interview preparation pack
- WhatsApp-friendly application summary
- Suggested professional file names
- Application tracker
- Downloadable readiness report
- Downloadable tracker CSV

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy

Deploy on Streamlit Community Cloud and use `app.py` as the main file.

Make sure the `.streamlit/config.toml` folder is included in your GitHub repository so the light theme loads correctly.

## Notes

This is a privacy-aware, rule-based MVP. It does not call an external AI API and it does not store uploaded documents. Uploaded files are processed in memory only.

Always verify the official advert requirements before submitting an application. Applicants should not include false skills, qualifications or experience.
