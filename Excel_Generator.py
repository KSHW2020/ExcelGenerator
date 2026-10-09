import tkinter as tk
from tkinter import messagebox
from pathlib import Path

import win32com.client
from openpyxl import Workbook


# ============================================================
# SETTINGS
# ============================================================

EXCEL_FILE_NAME = "output.xlsx"

# Excel output location
OUTPUT_FOLDER = Path(
    r"C:\Users\al103510\OneDrive - Qorvo\Desktop"
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

    openpyxl is used for the initial creation so Excel COM
    does not need to perform SaveAs() into the OneDrive folder.
    """

    try:

        # Check that the Desktop folder exists
        if not OUTPUT_FOLDER.exists():

            messagebox.showerror(
                "Folder Error",
                "The output folder cannot be found:\n\n"
                f"{OUTPUT_FOLDER}"
            )

            return False

        # Create Excel only if it does not already exist
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
            f"Error:\n{e}"
        )

        return False


# ============================================================
# CONNECT TO EXCEL
# ============================================================

def get_excel_sheet():
    """
    Connect to Microsoft Excel.

    If Excel is already running:
        Connect to the existing Excel application.

    If Excel is not running:
        Start Excel.

    If output.xlsx is already open:
        Use the existing workbook.

    Otherwise:
        Open output.xlsx.
    """

    global excel_app
    global workbook

    try:

        # ----------------------------------------------------
        # Make sure output.xlsx exists first
        # ----------------------------------------------------

        if not create_excel_file():
            return None, None


        # ----------------------------------------------------
        # Connect to Excel
        # ----------------------------------------------------

        try:

            excel_app = win32com.client.GetActiveObject(
                "Excel.Application"
            )

        except Exception:

            excel_app = win32com.client.DispatchEx(
                "Excel.Application"
            )


        # Make Excel visible
        excel_app.Visible = True


        # ----------------------------------------------------
        # Find output.xlsx if it is already open
        # ----------------------------------------------------

        workbook = None

        target_path = str(
            EXCEL_FILE.resolve()
        ).lower()


        for wb in excel_app.Workbooks:

            try:

                workbook_path = str(
                    Path(wb.FullName).resolve()
                ).lower()

                if workbook_path == target_path:

                    workbook = wb
                    break

            except Exception:

                pass


        # ----------------------------------------------------
        # Open workbook if it is not already open
        # ----------------------------------------------------

        if workbook is None:

            workbook = excel_app.Workbooks.Open(
                str(EXCEL_FILE)
            )


        # Use active worksheet
        worksheet = workbook.ActiveSheet


        return workbook, worksheet


    except Exception as e:

        messagebox.showerror(
            "Excel Error",
            "Unable to connect to Microsoft Excel.\n\n"
            f"Error:\n{e}"
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
    Convert:

    1  -> A
    2  -> B
    3  -> C

    ...

    26 -> Z
    27 -> AA
    28 -> AB
    29 -> AC
    """

    result = ""

    while number > 0:

        number -= 1

        result = (
            chr(65 + number % 26)
            + result
        )

        number //= 26

    return result


# ============================================================
# GENERATE BUTTON
# ============================================================

def generate_data():

    global pending_row
    global pending_letter


    # --------------------------------------------------------
    # Connect to Excel
    # --------------------------------------------------------

    wb, ws = get_excel_sheet()

    if ws is None:
        return


    # --------------------------------------------------------
    # Find next empty row
    # --------------------------------------------------------

    pending_row = find_next_row(ws)


    # --------------------------------------------------------
    # Generate corresponding letter
    # --------------------------------------------------------

    pending_letter = number_to_letters(
        pending_row
    )


    # --------------------------------------------------------
    # Display proposed entry
    # --------------------------------------------------------

    confirmation_label.config(
        text=(
            "Confirm this entry?\n\n"
            f"A{pending_row} = {pending_letter}\n"
            f"B{pending_row} = {pending_row}"
        )
    )


    # --------------------------------------------------------
    # Disable Generate button
    #
    # Prevents accidental multiple clicking
    # --------------------------------------------------------

    generate_button.config(
        state="disabled"
    )


    # --------------------------------------------------------
    # Show YES / NO
    # --------------------------------------------------------

    confirmation_frame.pack(
        pady=10
    )


    status_label.config(
        text="Waiting for confirmation...",
        fg="#D97706"
    )


# ============================================================
# YES BUTTON
# ============================================================

def confirm_yes():

    global pending_row
    global pending_letter


    if pending_row is None:
        return


    # --------------------------------------------------------
    # Connect / verify Excel
    # --------------------------------------------------------

    wb, ws = get_excel_sheet()

    if ws is None:

        reset_ui()
        return


    try:

        # ----------------------------------------------------
        # Safety check
        #
        # Make sure the row has not been changed after the
        # Generate button was clicked.
        # ----------------------------------------------------

        column_a_value = ws.Cells(
            pending_row,
            1
        ).Value

        column_b_value = ws.Cells(
            pending_row,
            2
        ).Value


        if (
            column_a_value is not None
            or column_b_value is not None
        ):

            messagebox.showwarning(
                "Row Already Used",
                "The target row is no longer empty.\n\n"
                "Nothing was added.\n"
                "Please click GENERATE again."
            )

            reset_ui()

            return


        # ----------------------------------------------------
        # Write Column A
        # ----------------------------------------------------

        ws.Cells(
            pending_row,
            1
        ).Value = pending_letter


        # ----------------------------------------------------
        # Write Column B
        # ----------------------------------------------------

        ws.Cells(
            pending_row,
            2
        ).Value = pending_row


        # ----------------------------------------------------
        # Save workbook
        # ----------------------------------------------------

        wb.Save()


        # ----------------------------------------------------
        # Success message
        # ----------------------------------------------------

        status_label.config(
            text=f"Successfully added: {pending_letter} / {pending_row}",
            fg="#107C10"
        )


        # Reset interface but preserve success message
        reset_ui(
            keep_status=True
        )


    except Exception as e:

        messagebox.showerror(
            "Write Error",
            "Unable to update Excel.\n\n"
            f"Error:\n{e}"
        )

        reset_ui()


# ============================================================
# NO BUTTON
# ============================================================

def confirm_no():

    # Nothing is written to Excel

    status_label.config(
        text="Cancelled - nothing was added.",
        fg="#C00000"
    )

    reset_ui(
        keep_status=True
    )


# ============================================================
# RESET USER INTERFACE
# ============================================================

def reset_ui(keep_status=False):

    global pending_row
    global pending_letter


    # Clear pending data
    pending_row = None
    pending_letter = None


    # --------------------------------------------------------
    # Hide YES / NO buttons
    # --------------------------------------------------------

    confirmation_frame.pack_forget()


    # --------------------------------------------------------
    # Restore instruction
    # --------------------------------------------------------

    confirmation_label.config(
        text="Click GENERATE to create the next entry."
    )


    # --------------------------------------------------------
    # Enable Generate again
    # --------------------------------------------------------

    generate_button.config(
        state="normal"
    )


    # --------------------------------------------------------
    # Reset status if required
    # --------------------------------------------------------

    if not keep_status:

        status_label.config(
            text="Ready",
            fg="#555555"
        )


# ============================================================
# MAIN APPLICATION WINDOW
# ============================================================

root = tk.Tk()


root.title(
    "Excel Generator"
)


root.geometry(
    "430x380"
)


root.resizable(
    False,
    False
)


# ============================================================
# TITLE
# ============================================================

title_label = tk.Label(
    root,
    text="Excel Data Generator",
    font=(
        "Arial",
        18,
        "bold"
    )
)


title_label.pack(
    pady=(25, 10)
)


# ============================================================
# EXCEL FILE INFORMATION
# ============================================================

file_label = tk.Label(
    root,
    text="Excel: Desktop\\output.xlsx",
    font=(
        "Arial",
        9
    ),
    fg="#666666"
)


file_label.pack()


# ============================================================
# CONFIRMATION INFORMATION
# ============================================================

confirmation_label = tk.Label(
    root,
    text="Click GENERATE to create the next entry.",
    font=(
        "Arial",
        11
    ),
    height=4
)


confirmation_label.pack(
    pady=(10, 5)
)


# ============================================================
# GENERATE BUTTON
# ============================================================

generate_button = tk.Button(
    root,
    text="GENERATE",
    font=(
        "Arial",
        14,
        "bold"
    ),
    width=18,
    height=2,
    command=generate_data,
    bg="#0078D4",
    fg="white",
    activebackground="#005A9E",
    activeforeground="white",
    cursor="hand2"
)


generate_button.pack(
    pady=5
)


# ============================================================
# YES / NO FRAME
# ============================================================

confirmation_frame = tk.Frame(
    root
)


# ============================================================
# YES BUTTON
# ============================================================

yes_button = tk.Button(
    confirmation_frame,
    text="YES",
    font=(
        "Arial",
        12,
        "bold"
    ),
    width=9,
    command=confirm_yes,
    bg="#107C10",
    fg="white",
    activebackground="#0B6A0B",
    activeforeground="white",
    cursor="hand2"
)


yes_button.pack(
    side="left",
    padx=10
)


# ============================================================
# NO BUTTON
# ============================================================

no_button = tk.Button(
    confirmation_frame,
    text="NO",
    font=(
        "Arial",
        12,
        "bold"
    ),
    width=9,
    command=confirm_no,
    bg="#D13438",
    fg="white",
    activebackground="#A4262C",
    activeforeground="white",
    cursor="hand2"
)


no_button.pack(
    side="left",
    padx=10
)


# IMPORTANT:
#
# confirmation_frame is intentionally NOT packed here.
#
# Therefore:
#
# YES / NO are hidden when the application starts.
#
# They appear only after GENERATE is clicked.


# ============================================================
# STATUS
# ============================================================

status_label = tk.Label(
    root,
    text="Ready",
    font=(
        "Arial",
        10,
        "bold"
    ),
    fg="#555555"
)


status_label.pack(
    pady=15
)


# ============================================================
# START APPLICATION
# ============================================================

root.mainloop()