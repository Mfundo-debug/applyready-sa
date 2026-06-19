# ApplyReady SA

ApplyReady SA is a Streamlit application-readiness tool for South African job seekers, students, graduates, first-time applicants, internship applicants, learnership applicants, bursary applicants and graduate-programme applicants.

It helps applicants check whether their application pack is complete, professional and aligned before they submit.

## Version 2.5: Application Fit Engine + CV Formaliser

Version 2.5 builds on the Version 2 fit engine by adding a formal CV drafting feature. The app remains privacy-aware and rule-based: it does not call an external AI API and it does not store uploaded documents.

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
- Formal CV draft generated from supplied applicant information
- CV polish checklist
- Downloadable CV draft as Markdown
- Downloadable CV draft as Word document
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

## Important notes

ApplyReady SA is a preparation assistant. It does not guarantee employment, interviews, funding, admission or selection.

The CV Formaliser reorganises and polishes information provided by the applicant. It must not be used to invent skills, qualifications, experience or achievements. Applicants should review every output carefully before submitting an application.

During beta testing, users should avoid uploading sensitive documents such as ID copies, bank statements or proof of residence. The checklist can be used instead.
