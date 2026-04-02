# DG_Sample_Code_Check

Healthcare benefits correspondence automation system built on Pick/MultiValue (UniData). Manages COBRA notices, HIPAA privacy notices, and disabled dependent letters.

## System Analysis Report

- **HTML Report (GitHub Pages):** [https://sachetcognition.github.io/DG_Sample_Code_Check/](https://sachetcognition.github.io/DG_Sample_Code_Check/)
- **Word Document:** [docs/DG_Analysis_Report.docx](docs/DG_Analysis_Report.docx)

The report covers:
1. Executive Summary
2. Record Types — Full Analysis (10 record types)
3. The 9 Letter Types — Identified from Code
4. Dependency Graph
5. Current Coupling Map
6. Coupling Points Table
7. Decoupling Strategy (8 extracted subroutines)
8. Decoupled Architecture
9. Priority Order for Decoupling
10. Program Flow Diagrams (COBRAN, HIPAA.NPP, PRINT.DISABLED.DEP.LETTERS)

## Source Programs

| Program | Purpose |
|---------|---------|
| `COBRAN` | COBRA initial notice generation via SmartComm |
| `HIPAA.NPP` | HIPAA Notice of Privacy Practices (Self-Funded via SmartComm, Fully-Insured via PCL) |
| `PRINT.DISABLED.DEP.LETTERS` | Disabled dependent letter generation via SmartComm |
