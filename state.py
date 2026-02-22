# ─────────────────────────────────────────────
#  state.py  —  Shared application state
#  All modules import from here instead of
#  declaring their own globals.
# ─────────────────────────────────────────────

import tkinter as tk

# ── Data ──────────────────────────────────────
excel_path_var    = None   # tk.StringVar  (set in main.py after root exists)
url_var           = None   # tk.StringVar  (set in main.py after root exists)

workbook          = None
ws                = None
excel_data        = []     # list[dict]  – one dict per respondent row
excel_headers     = []     # list[str]   – column names from Excel row 1
columns           = []     # excel_headers + ["NULL"]
sheet_container   = None
num_sections      = 1

# ── Form / URL ────────────────────────────────
entry_ids         = []     # Google Form entry IDs parsed from URL
parsed_url        = None
query_dict        = None
mapping_vars      = {}     # {entry_id: tk.StringVar(column_name)}

# ── Submission ────────────────────────────────
is_submitting     = False
final_output_urls = []

# ── Project ───────────────────────────────────
current_project_path = None

# ── Undo / Redo ───────────────────────────────
undo_stack = []
redo_stack = []
MAX_UNDO   = 20

# ── UI widget references (set in main.py) ─────
root          = None
preview_text  = None
progress_var  = None
scrollable_frame = None
canvas        = None
