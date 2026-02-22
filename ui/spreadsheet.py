# ─────────────────────────────────────────────
#  ui/spreadsheet.py  —  Data preview grid
#  Uses tksheet for fast virtual rendering
# ─────────────────────────────────────────────

from tksheet import Sheet
import state


def update_spreadsheet_view():
    if not state.excel_headers or not state.excel_data:
        return

    # Build rows: [row_number, col1, col2, ...]
    headers = ["#"] + list(state.excel_headers)
    rows = [
        [idx + 1] + [row_data.get(col, "") or "" for col in state.excel_headers]
        for idx, row_data in enumerate(state.excel_data)
    ]

    # Clear old sheet if exists
    for widget in state.sheet_container.winfo_children():
        widget.destroy()

    sheet = Sheet(
        state.sheet_container,
        headers=headers,
        data=rows,
        theme="light blue",
        show_row_index=False,
        row_index=False,
        default_column_width=120,
        header_bg="#d0d0d0",
        header_fg="black",
        table_bg="white",
        alternate_color="#f2f2f2",
    )
    sheet.enable_bindings("all")
    sheet.column_width(column=0, width=30)
    sheet.pack(fill="both", expand=True)