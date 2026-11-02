# PW1 Transistors — Group QMS (GitHub)

University of Bath · Year 2 Project Week · **The secret life of transistors**

This repository is the shared **Quality Management System (QMS)** for our group.  
**English only.** Put real project files here as you create them.

## What this repo is for

- Keep all controlled documents in one place  
- Use clear folders and file names  
- Show version control (branch → pull request → merge to `main`) for the QMS skill  

## Folder map (`qms/`)

| Folder | Put here |
|--------|----------|
| `Quality documents/Quality manual/` | Quality commitment / quality manual |
| `Quality documents/Procedure/` | **File naming procedure** (required) |
| `Quality documents/Forms/` | Quality forms (can stay empty) |
| `Quality control/` | Test plans and measurement / truth-table results |
| `Engineering/` | Schematics, code, **traceability matrix** |
| `Operations/` | Time plan / Gantt / schedule evidence |
| `Purchasing/` | Empty for this project week |
| `Sales_customer/` | Non-technical audience / presentation materials |
| `Human resources/` | Team charter, skills matrix, teamwork records |
| `Documents gallery/` | **Document register only** (list of every file + path) |

Each folder has a short `README.md` explaining what to add.

## How teammates should work

1. **Never push straight to `main` for new work** — open a branch.  
2. Name branches clearly, e.g. `docs/team-charter`, `ops/gantt`, `eng/traceability`.  
3. Add your file in the correct folder.  
4. Update `qms/Documents gallery/Document_Register.csv`.  
5. Open a **Pull Request** → merge to `main`.  

```text
git pull
git checkout -b docs/short-description
# add/edit files
git add -A
git commit -m "Add team charter v1.0"
git push -u origin HEAD
# then open a Pull Request on GitHub
```

## File naming (must follow)

```text
PW1_Transistors_<AREA>_<DocumentName>_v<Major.Minor>.<ext>
```

Examples:

- `PW1_Transistors_HR_TeamCharter_v1.0.docx`  
- `PW1_Transistors_OPS_Gantt_v1.0.png`  
- `PW1_Transistors_PROC_FileNaming_v1.0.docx`  

Area codes: `QM` `PROC` `FORM` `QC` `ENG` `OPS` `HR` `SALES` `GAL`

Write the full naming rules in:  
`qms/Quality documents/Procedure/` (team creates this document).

## Hard deadlines (Week 6)

- **Thu 11:15** — QMS audit (find any named file in ≤30 seconds)  
- **Fri 10:15** — Full adder challenge  
- **Fri 12:00** — Individual video on Moodle  
- **Fri 12:15** — Kit return  

## Do not commit

- Secrets / passwords  
- Huge raw scope captures unless needed  
- Personal notes outside the QMS tree  
