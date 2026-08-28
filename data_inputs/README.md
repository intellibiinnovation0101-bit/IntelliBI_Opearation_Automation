# data_inputs/ — local source workbooks for the Batch Planner

`ops_reports_action/pyBatchPlanner.py` reads four local Excel workbooks (it does
not pull them from the API). Keep the current copies here; `common/_bootstrap.py`
points `BP_BASE_DIR` at this folder automatically.

| File | Sheets read |
|------|-------------|
| `CourseAndTechnologiesMappingDocument.xlsx` | `CourseAndTechnologiesMappingDoc` |
| `TechnologyMappingDocuments.xlsx` | `TechnologyMapping` |
| `IntellBIAttendance.xlsx` | `Sessions` |
| `IntelliBIStudentInfo.xlsx` | `Students`, `ClassLearnerTeacherEnrolled` |

Refresh these workbooks (export them from their Google Sheets) before running
the Batch Planner if you need the very latest enrolment/attendance data. This
folder is git-ignored. Override the location with `batch_planner.base_dir` in
`config/config.yaml` or the `BP_BASE_DIR` environment variable.
