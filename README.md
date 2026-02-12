# AI Excel Practice Bot

This app generates role-based Excel practice files for students and auto-evaluates their answers.

## Features
- Collects student details: Name, phone, and optional working information.
- Shows company name and current role fields only when working status is `Yes`.
- Generates large role-based practice data in Excel.
- Creates 20 standard auto-generated questions in serial order (1..20).
- Enforces global no-repeat questions across participants using a persistent registry.
- Adds an `Activities` sheet so students can do extra Excel tasks alongside questions.
- Each activity includes an expected output/answer hint.
- Accepts solved workbook upload and evaluates in percentage view.
- Shows a certificate-style score card and a score breakdown chart.
- Supports basic branding: institute name, logo upload, and colors.
- Shows logo at the top-left header and inside certificate score card.
- Adds LinkedIn-ready PNG score card download with institute branding.

## Run on Mac
```bash
cd "/Users/rashokreddy/Documents/Excel Logics"
python3 -m pip install --user -r requirements.txt
python3 -m streamlit run app.py
```

## Global Uniqueness
- Issued question signatures are stored in:
  - `data/used_question_signatures.txt`
- Any question already used for one student will not be reused for future students.

## Student Flow
1. Fill student details and generate workbook.
2. Download workbook and solve in `Answers` sheet.
3. Complete optional tasks in `Activities` sheet.
4. Upload the same workbook for evaluation.
5. Check certificate card (with student details), chart, and detailed report.
