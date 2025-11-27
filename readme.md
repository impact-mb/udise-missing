CPRF Data Quality Checker – Adolescent Records

Streamlit app to:

- Import raw CPRF delivery data (all columns as text)
- Filter to PROGRAMSUBTYPENAME = "ADOLOSCENT"
- Identify records with missing **School UDISE**
- Apply data-quality checks:
  - UDISE code missing
  - Child school name missing/null
  - Date of birth set to 1 Jan
  - Phone number missing/invalid (India 10-digit mobile)
  - Caste marked as DONT KNOW / DONT WISH
  - Parent consent missing/null
- Export:
  - One combined Excel: `ALL_CPRF_issues.xlsx`
  - One Excel file per **ProgramLaunchName**
- Add footer with report timestamp, checks applied, and app version
- Provide ZIP download of all files plus an on-screen summary table
- Maintain a persistent run counter (how many times the tool has been executed)
