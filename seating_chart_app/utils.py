import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox

# Constants
SEAT_POSITIONS = ["Left", "Middle", "Right"]

def prompt_for_new_roll_numbers(position):
    """
    Prompt the user to select a new file for additional roll numbers.
    This function opens a file dialog allowing the user to select an Excel file.
    """
    root = Tk()
    root.withdraw()  # Hide the root window
    messagebox.showinfo("Roll Numbers Exhausted", f"Please select a new file for {position} roll numbers.")
    
    input_path = filedialog.askopenfilename(title=f"Select {position} roll numbers file", filetypes=[("Excel files", "*.xlsx *.xls")])
    
    if os.path.exists(input_path):
        print(f"Selected file for {position} roll numbers: {input_path}")
        return load_roll_numbers_from_file(input_path, position)
    else:
        print(f"Invalid file path for {position}.")
        return None


def load_roll_numbers_from_file(filepath, position):
    """
    Load roll numbers from the specified Excel file.
    """
    try:
        df = pd.read_excel(filepath)
        df.columns = df.columns.str.strip()

        if 'Roll Number' not in df.columns:
            messagebox.showerror("Error", f"'Roll Number' column not found in the Excel file for {position} roll numbers.")
            return None

        return df['Roll Number'].dropna().tolist()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read roll numbers file for {position} roll numbers: {e}")
        return None


def load_room_details(filepath):
    """
    Load room details from an Excel file. Checks for required columns.
    """
    try:
        df = pd.read_excel(filepath)
        df.columns = df.columns.str.strip()

        # Check required columns in the room details file
        required_columns = ['Room Number', 'Number of Rows', 'Number of Bench', 'Number of Student per Bench',
                            'Left Path', 'Middle Path', 'Right Path', 'Left Name', 'Middle Name', 'Right Name']
        if not all(col in df.columns for col in required_columns):
            messagebox.showerror("Error", f"Excel file must contain the following columns: {', '.join(required_columns)}.")
            return None

        return df
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read room details file: {e}")
        return None


def save_seating_chart_to_excel(seating_chart_list, output_filename):
    """
    Save the seating chart data to an Excel file.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Side, Font

    # Initialize workbook and worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Seating Chart"

    # Set the header font style
    bold_font = Font(bold=True)
    medium_border = Border(left=Side(style='medium'), right=Side(style='medium'), top=Side(style='medium'), bottom=Side(style='medium'))

    # Write the seating chart data to Excel
    for seating_df in seating_chart_list:
        for row in seating_df.itertuples(index=False):
            ws.append(row)

    # Save the workbook to a file
    try:
        if not output_filename.endswith(".xlsx"):
            output_filename += ".xlsx"
        wb.save(output_filename)
        messagebox.showinfo("Success", f"Seating chart saved to {output_filename}")
    except PermissionError:
        messagebox.showerror("Error", f"Permission denied: Cannot save to {output_filename}. Please ensure the file is not open and you have write permissions.")


def prompt_for_output_file():
    """
    Prompt the user to select a location for saving the output Excel file.
    """
    root = Tk()
    root.withdraw()  # Hide the root window
    output_path = filedialog.asksaveasfilename(title="Save Seating Chart As", filetypes=[("Excel files", "*.xlsx")])
    
    return output_path


def show_progress_window(max_value):
    """
    Create and return a progress window for the seating chart generation process.
    """
    from tkinter import Toplevel, Label, ttk, StringVar

    progress_window = Toplevel()
    progress_window.title("Progress")

    label = Label(progress_window, text="Generating seating chart...")
    label.pack()

    progress_var = StringVar()
    progress_bar = ttk.Progressbar(progress_window, variable=progress_var, maximum=max_value)
    progress_bar.pack()

    return progress_window, progress_var
