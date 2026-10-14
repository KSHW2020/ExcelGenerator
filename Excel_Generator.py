import os
import platform
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

try:
    import win32com.client
except ImportError:
    win32com = None

from openpyxl import Workbook


# ============================================================
# SETTINGS
# ============================================================

EXCEL_FILE_NAME = "output.xlsx"

# Use a portable folder instead of a hardcoded Windows user path.
OUTPUT_FOLDER = Path(
    os.environ.get(
        "EXCEL_GENERATOR_OUTPUT_FOLDER",
        str(Path.home() / "Desktop"),
    )
)
EXCEL_FILE = OUTPUT_FOLDER / EXCEL_FILE_NAME


# Temporary values waiting for YES / NO
pending_row = None
pending_letter = None


# Excel COM objects
excel_app = None
workbook = None


# ============================================================
# CREATE EXCEL FILE IF NEEDED
# ============================================================

def create_excel_file():
    """
    Create output.xlsx if it does not already exist.
    """
    try:
        if not OUTPUT_FOLDER.exists():
            try:
                OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
            except OSError as e:
                messagebox.showerror(
                    "Folder Error",
                    "The output folder cannot be created or accessed:\n\n"
                    f"{OUTPUT_FOLDER}\n\nError:\n{e}",
                )
                return False

        if not EXCEL_FILE.exists():
            new_workbook = Workbook()
            worksheet = new_workbook.active
            worksheet.title = "Sheet1"
            new_workbook.save(EXCEL_FILE)

        return True

    except Exception as e:
        messagebox.showerror(
            "File Error",
            "Unable to create the Excel file.\n\n"
            f"Location:\n{EXCEL_FILE}\n\n"
            f"Error:\n{e}",
        )
        return False


# ============================================================
# CONNECT TO EXCEL
# ============================================================

def get_excel_sheet():
    """
    Connect to Microsoft Excel.
    """
    global excel_app
    global workbook

    if platform.system() != "Windows" or win32com is None:
        messagebox.showerror(
            "Platform Error",
            "This program requires Microsoft Excel on Windows and the pywin32 package.\n\n"
            "Install pywin32 and run on Windows to use Excel automation.",
        )
        return None, None

    try:
        if not create_excel_file():
            return None, None

        try:
            excel_app = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            excel_app = win32com.client.DispatchEx("Excel.Application")

        excel_app.Visible = True

        workbook = None
        target_path = str(EXCEL_FILE.resolve()).lower()

        for wb in excel_app.Workbooks:
            try:
                workbook_path = str(Path(wb.FullName).resolve()).lower()
                if workbook_path == target_path:
                    workbook = wb
                    break
            except Exception:
                pass

        if workbook is None:
            workbook = excel_app.Workbooks.Open(str(EXCEL_FILE))

        worksheet = workbook.ActiveSheet
        return workbook, worksheet

    except Exception as e:
        messagebox.showerror(
            "Excel Error",
            "Unable to connect to Microsoft Excel.\n\n"
            f"Error:\n{e}",
        )
        return None, None


# ============================================================
# FIND NEXT EMPTY ROW
# ============================================================

def find_next_row(worksheet):
    """
    Find the first empty row in Column A.
    """
    row = 1
    while worksheet.Cells(row, 1).Value is not None:
        row += 1
    return row


# ============================================================
# CONVERT NUMBER TO LETTER
# ============================================================

def number_to_letters(number):
    """
    1 -> A
    2 -> B
    ...
    26 -> Z
    27 -> AA
    """
    if number <= 0:
        return ""

    result = ""
    while number > 0:
        number -= 1
        result = chr(65 + number % 26) + result
        number //= 26
    return result


# ============================================================
# GENERATE BUTTON
# ============================================================

def generate_data():
    global pending_row
    global pending_letter

    wb, ws = get_excel_sheet()
    if ws is None:
        return

    pending_row = find_next_row(ws)
    pending_letter = number_to_letters(pending_row)

    confirmation_label.config(
        text=(
            "Confirm this entry?\n\n"
            f"A{pending_row} = {pending_letter}\n"
            f"B{pending_row} = {pending_row}"
        )
    )

    generate_button.config(state="disabled")
    confirmation_frame.pack(pady=10)

    status_label.config(
        text="Waiting for confirmation...",
        fg="#D97706",
    )


# ============================================================
# YES BUTTON
# ============================================================

def confirm_yes():
    global pending_row
    global pending_letter

    if pending_row is None:
        return

    wb, ws = get_excel_sheet()
    if ws is None:
        reset_ui()
        return

    try:
        column_a_value = ws.Cells(pending_row, 1).Value
        column_b_value = ws.Cells(pending_row, 2).Value

        if column_a_value is not None or column_b_value is not None:
            messagebox.showwarning(
                "Row Already Used",
                "The target row is no longer empty.\n\n"
                "Nothing was added.\n"
                "Please click GENERATE again.",
            )
            reset_ui()
            return

        ws.Cells(pending_row, 1).Value = pending_letter
        ws.Cells(pending_row, 2).Value = pending_row

        try:
            wb.Save()
        except Exception as e:
            messagebox.showerror(
                "Save Error",
                "Failed to save the workbook.\n\n"
                f"Error:\n{e}",
            )
            reset_ui()
            return

        status_label.config(
            text=f"Successfully added: {pending_letter} / {pending_row}",
            fg="#107C10",
        )

        reset_ui(keep_status=True)

    except Exception as e:
        messagebox.showerror(
            "Write Error",
            "Unable to update Excel.\n\n"
            f"Error:\n{e}",
        )
        reset_ui()


# ============================================================
# NO BUTTON
# ============================================================

def confirm_no():
    status_label.config(
        text="Cancelled - nothing was added.",
        fg="#C00000",
    )
    reset_ui(keep_status=True)


# ============================================================
# RESET USER INTERFACE
# ============================================================

def reset_ui(keep_status=False):
    global pending_row
    global pending_letter

    pending_row = None
    pending_letter = None

    confirmation_frame.pack_forget()
    confirmation_label.config(text="Click GENERATE to create the next entry.")
    generate_button.config(state="normal")

    if not keep_status:
        status_label.config(text="Ready", fg="#555555")


# ============================================================
# MAIN APPLICATION WINDOW
# ============================================================

root = tk.Tk()
root.title("Excel Generator")
root.geometry("430x380")
root.resizable(False, False)

title_label = tk.Label(
    root,
    text="Excel Data Generator",
    font=("Arial", 18, "bold"),
)
title_label.pack(pady=(25, 10))

file_label = tk.Label(
    root,
    text=f"Excel: {EXCEL_FILE}",
    font=("Arial", 9),
    fg="#666666",
)
file_label.pack()

confirmation_label = tk.Label(
    root,
    text="Click GENERATE to create the next entry.",
    font=("Arial", 11),
    height=4,
)
confirmation_label.pack(pady=(10, 5))

generate_button = tk.Button(
    root,
    text="GENERATE",
    font=("Arial", 14, "bold"),
    width=18,
    height=2,
    command=generate_data,
    bg="#0078D4",
    fg="white",
    activebackground="#005A9E",
    activeforeground="white",
    cursor="hand2",
)
generate_button.pack(pady=5)

confirmation_frame = tk.Frame(root)

yes_button = tk.Button(
    confirmation_frame,
    text="YES",
    font=("Arial", 12, "bold"),
    width=9,
    command=confirm_yes,
    bg="#107C10",
    fg="white",
    activebackground="#0B6A0B",
    activeforeground="white",
    cursor="hand2",
)
yes_button.pack(side="left", padx=10)

no_button = tk.Button(
    confirmation_frame,
    text="NO",
    font=("Arial", 12, "bold"),
    width=9,
    command=confirm_no,
    bg="#D13438",
    fg="white",
    activebackground="#A4262C",
    activeforeground="white",
    cursor="hand2",
)
no_button.pack(side="left", padx=10)

status_label = tk.Label(
    root,
    text="Ready",
    font=("Arial", 10, "bold"),
    fg="#555555",
)
status_label.pack(pady=15)

root.mainloop()
