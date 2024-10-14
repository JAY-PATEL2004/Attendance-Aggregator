from math import floor
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import PyPDF2
import pandas as pd
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import threading
from openpyxl.workbook import Workbook

# Function to extract table data from PDF
def extract_table_from_pdf(pdf_path):
    pdf_reader = PyPDF2.PdfReader(pdf_path)
    text = ""
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text += page.extract_text()
    return text

# Function to parse extracted text into a DataFrame
# Function to parse extracted text into a DataFrame with "Total Classes" and "Attendance" columns
# Function to parse extracted text into a DataFrame with "Total Classes", "Attendance", and "Attendance %" columns
def parse_text_to_dataframe(text, subject):
    lines = text.split('\n')
    data = []
    fg = False
    for line in lines:
        if fg == False:
            fg = True
            continue
        parts = line.split()
        if len(parts) >= 3:
            if parts[0].isdigit():
                parts = parts[1:]
        if len(parts) >= 2:
            roll_no = parts[0].upper()
            if parts[-1].isdigit() and parts[-2].isdigit():
                name = " ".join(parts[1:-2])
                total_classes = int(parts[-2])
                attendance = int(parts[-1])
                attendance_percentage = round((attendance / total_classes) * 100,2) if total_classes > 0 else 0
            elif parts[-1].isdigit() and not parts[-2].isdigit():
                name = " ".join(parts[1:-1])
                total_classes = int(parts[-1])
                attendance = "N/A"
                attendance_percentage = "N/A"
            else:
                name = " ".join(parts[1:])
                total_classes = "N/A"
                attendance = "N/A"
                attendance_percentage = "N/A"
        else:
            continue  # Skip rows that don't fit the expected pattern

        data.append([roll_no, name, total_classes, attendance, attendance_percentage])

    # Create DataFrame with "Total Classes", "Attendance", and "Attendance %" columns for each subject
    df = pd.DataFrame(data, columns=['Roll No', 'Name', f'{subject} Total Classes', f'{subject} Attendance', f'{subject} Attendance %'])
    return df



# Function to check for conflicts (same roll no with different names)
def check_conflicts(combined_data):
    if 'Name' not in combined_data.columns:
        raise KeyError("The 'Name' column is missing in the dataset.")
    
    conflict_rows = combined_data.groupby('Roll No').filter(lambda x: len(x['Name'].unique()) > 1)
    return conflict_rows

# Function to combine data from multiple PDFs
def combine_data_from_pdfs(pdf_files, root):
    combined_data = pd.DataFrame()
    for pdf_file in pdf_files:
        subject = os.path.splitext(os.path.basename(pdf_file))[0]
        text = extract_table_from_pdf(pdf_file)
        df = parse_text_to_dataframe(text, subject)
        print(df)
        if combined_data.empty:
            combined_data = df
        else:
            combined_data = pd.merge(combined_data, df, on=['Roll No', 'Name'], how='outer')
            combined_data.fillna('N/A', inplace=True)  # Assuming 'N/A' for missing attendance
            combined_data.sort_values(axis=0, inplace=True, by="Roll No")
    #combined_data = combined_data.merge(df, on='Roll No', how='outer')
    # # for index, row in combined_data.iterrows():
    # #     if pd.notna(row['Name_x']) and pd.notna(row['Name_y']):
    # #         if row['Name_x'].lower() != row['Name_y'].lower():
    # #                 conflicts.append((index, row['Roll No'], row['Name_x'], row['Name_y']))
    print(combined_data)
    total_classes_columns = [col for col in combined_data.columns if 'Total Classes' in col]
    attendance_columns = [col for col in combined_data.columns if 'Attendance' in col and '%' not in col]

    combined_data['Total Classes'] = combined_data[total_classes_columns].apply(pd.to_numeric, errors='coerce').sum(axis=1, min_count=1)
    combined_data['Total Classes Attended'] = combined_data[attendance_columns].apply(pd.to_numeric, errors='coerce').sum(axis=1, min_count=1)
    
    # Calculate 'Total Attendance %' based on the totals
    combined_data['Total Attendance %'] = (combined_data['Total Classes Attended'] / combined_data['Total Classes']) * 100
    combined_data['Total Attendance %'] = combined_data['Total Attendance %'].fillna(0).round(2)
    print(combined_data)
    # for i in range(len(combined_data)):
    #     if i==0:
    #         i+=1
    #         continue
    #     for j in range(len(combined_data[i])):
    #         if j==0 or j==1:
    #             j+=1
    #             continue
    #         if combined_data[i][j]!="N/A":
    #             combined_data[i][j] = int(floor(combined_data[i][j]))
    conflict_rows = check_conflicts(combined_data)
    print(combined_data)
    print(conflict_rows)
    if not conflict_rows.empty:
        combined_data = resolve_conflicts(conflict_rows, combined_data, root, df)
    combined_data.insert(0, 'S No.', range(1, len(combined_data) + 1))

    return combined_data

def resolve_conflicts(conflict_rows, combined_data, root, df):
    for roll_no in conflict_rows['Roll No'].unique():
        conflicting_names = combined_data[combined_data['Roll No'] == roll_no]['Name'].unique()
        chosen_name = simpledialog.askstring(
            "Name Conflict",
            f"Roll No {roll_no} has multiple names: {', '.join(conflicting_names)}. Please choose one name:",
            initialvalue=conflicting_names[0], parent=root
        )
        combined_data.loc[combined_data['Roll No'] == roll_no, 'Name'] = chosen_name
        i = 2
        for subject in combined_data.columns:
            if i==2 or i==1:
                i-=1
                continue
            combined_data.loc[:,subject] = combined_data.loc[:,subject].astype(str)
            a = combined_data.loc[combined_data['Roll No']==roll_no, subject]
            for b in a:
                print(str(b))
            a = str(a.min())
            combined_data.loc[combined_data['Roll No']==roll_no, subject] = a
        print(combined_data)
            #combined_data.loc[combined_data['Roll No']==roll_no, subject] = combined_data.loc[combined_data['Roll No']==roll_no & combined_data[subject]!='N/A', subject]

        #print(df)
    # attendance_columns = combined_data.columns.difference(['Roll No', 'Name'])  # Subject columns
    # merged_data = combined_data.groupby(['Roll No', 'Name'], as_index=False).agg(lambda x: x.ffill().bfill().iloc[0] if not x.empty else 'N/A')

            # Get all attendance data associated with the conflicting names
    #     attendance_data = combined_data[combined_data['Roll No'] == roll_no].iloc[:, 1:]
    #     print(attendance_data)

    #         # For each subject, we should take the attendance where the name is the chosen one
    #     for subject in attendance_data.columns:
    #         if chosen_name in conflicting_names:
    #                 # If the chosen name is found, keep its attendance; otherwise, use the first non-null value
    #             attendance_value = attendance_data[attendance_data.index[attendance_data['Name'] == chosen_name]][subject]
    #             if not attendance_value.empty:
    #                 combined_data.loc[combined_data['Roll No'] == roll_no, subject] = attendance_value.values[0]
    #             else:
    #                 combined_data.loc[combined_data['Roll No'] == roll_no, subject] = attendance_data[subject].dropna().iloc[0] if not attendance_data[subject].dropna().empty else 'N/A'
    # print(combined_data)

    # # Step 2: Remove duplicates after resolving conflicts
    # combined_data = combined_data.drop_duplicates()

    # #print(combined_data)
    # #     attendance_data = combined_data[combined_data['Roll No'] == roll_no].iloc[:, 2:].copy()
    # #     attendance_data['Name'] = chosen_name

    # #         # Update the attendance in the main DataFrame
    # #     for subject in attendance_data.columns[1:]:
    # #             # Use the first non-null value for attendance
    # #         attendance_value = attendance_data[subject].dropna().iloc[0] if not attendance_data[subject].dropna().empty else 'N/A'
    # #         combined_data.loc[combined_data['Roll No'] == roll_no, subject] = attendance_value
    combined_data = combined_data.drop_duplicates()

    return combined_data

# Function to create the output PDF with a table
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch

# Function to create the output PDF with proper text wrapping
def create_pdf(output_path, data):
    pdf = SimpleDocTemplate(output_path, pagesize=A4)
    elements = []

    # Get sample styles for Paragraph
    styles = getSampleStyleSheet()
    body_style = styles['BodyText']
    body_style.wordWrap = 'CJK'  # Allow wrapping for body text
    body_style.fontSize = 8
    header_style = styles['Heading6']  # Use a different style for headers
    header_style.wordWrap = 'CJK'  # Allow wrapping for header text
    header_style.fontSize = 10

    def wrap_text_with_whitespace(data, max_width=30):
        """
        Custom function to wrap text at white spaces first and then fall back to breaking words if necessary.
        :param data: String that needs to be wrapped.
        :param max_width: Maximum width/length of characters before wrapping.
        :return: Wrapped string.
        """
        words = data.split()  # Split by whitespace
        lines = []
        current_line = []

        for word in words:
        # If adding the word exceeds the max width, wrap to a new line
            if sum(len(w) for w in current_line) + len(word) + len(current_line) > max_width:
                lines.append(" ".join(current_line))
                current_line = [word]
            else:
                current_line.append(word)

        if current_line:
            lines.append(" ".join(current_line))  # Add the remaining text

        wrapped_text = "\n".join(lines)
        return wrapped_text

    # Convert DataFrame to list of lists (including headers)
    table_data = []

    max_column_width = 8.0/len(data.columns) * inch

    # Create wrapped header cells
    # wrapped_header = [Paragraph(str(col), header_style) for col in data.columns]
    wrapped_header = []
    for col in data.columns:
        wrapped_col = wrap_text_with_whitespace(col, max_width=max_column_width)
        wrapped_header.append(Paragraph(str(wrapped_col), header_style))
    table_data.append(wrapped_header)  # Add headers
    # Set maximum column width
      # Adjust this value as needed

    # Create table data with wrapped text
    for row in data.values.tolist():
        wrapped_row = []
        i = 0
        for cell in row:
            temp = cell
            if i==2:
                temp = wrap_text_with_whitespace(cell, max_width=max_column_width)
            wrapped_cell = Paragraph(str(temp), body_style)
            wrapped_row.append(wrapped_cell)
            i+=1
        table_data.append(wrapped_row)


    # Define column widths based on max width and number of columns

    col_widths = [max_column_width] * len(data.columns)  # All columns have the same width

    # Create the table
    table = Table(table_data, colWidths=col_widths)

    # Add table style
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),  # Header background color
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),  # Header text color
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Center alignment
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Header font
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Padding for the header row
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),  # Background color for rows
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),  # Grid lines for table
        ('FONTSIZE', (0, 0), (-1, -1), 8),  # Font size for all cells
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Align text vertically to middle
    ]))

    # Build the PDF
    elements.append(table)  # Append the table to the elements list
    pdf.build(elements)



# Function to handle file selection and PDF generation
def generate_report(pdf_files, progress_var, root, output_format):
    if not pdf_files:
        messagebox.showerror("Error", "No files selected!")
        return

    combined_data = combine_data_from_pdfs(pdf_files, root)

    file_extension = ".pdf" if output_format == "PDF" else ".xlsx"
    file_types = [("PDF Files", "*.pdf")] if output_format == "PDF" else [("Excel Files", "*.xlsx")]

    output_path = filedialog.asksaveasfilename(
        title="Save Combined Attendance Report As",
        defaultextension=file_extension,
        filetypes=file_types
    )

    if output_path:
        if output_format == "PDF":
            threading.Thread(target=create_pdf, args=(output_path, combined_data)).start()
        elif output_format == "Excel":
            threading.Thread(target=create_excel, args=(output_path, combined_data)).start()
        messagebox.showinfo("Success", f"The combined attendance report has been generated as a {output_format} file!")

def select_files(file_listbox, file_paths):
    pdf_files = filedialog.askopenfilenames(
        title="Select PDF Files",
        filetypes=[("PDF Files", "*.pdf")],
        multiple=True
    )
    if pdf_files:
        for file in pdf_files:
            if file not in file_paths:  # Prevent duplicate files
                file_listbox.insert(tk.END, file)
                file_paths.add(file)

def deselect_file(file_listbox, file_paths):
    selected_index = file_listbox.curselection()
    if selected_index:
        file = file_listbox.get(selected_index)
        file_listbox.delete(selected_index)
        file_paths.remove(file)

# Function to run the report generation in a separate thread
def start_report_generation(file_listbox, progress_var, root, format_var):
    pdf_files = file_listbox.get(0, tk.END)
    output_format = format_var.get()
    progress_var.set(0)
    generate_report(pdf_files, progress_var, root, output_format)

def create_excel(output_path, data):
    data.to_excel(output_path, index=False)

# Create the main GUI window
def create_gui():
    root = tk.Tk()
    root.title("Attendance Aggregator")

    # Style Configuration
    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", background="#ccc")
    style.configure("TLabel", padding=6)
    style.configure("TProgressbar", thickness=20)

    file_paths = set()

    # File selection section
    file_frame = ttk.Frame(root, padding="10")
    file_frame.grid(row=0, column=0, sticky=tk.W + tk.E + tk.N + tk.S)

    ttk.Label(file_frame, text="Selected PDF Files:").grid(row=0, column=0, sticky=tk.W)
    
    file_listbox = tk.Listbox(file_frame, height=10, width=60)
    file_listbox.grid(row=1, column=0, padx=10, pady=10)

    button_frame = ttk.Frame(file_frame)
    button_frame.grid(row=1, column=1, padx=10, pady=10, sticky=tk.N)

    ttk.Button(button_frame, text="Select PDF Files", command=lambda: select_files(file_listbox, file_paths)).pack(pady=5)
    ttk.Button(button_frame, text="Deselect File", command=lambda: deselect_file(file_listbox, file_paths)).pack(pady=5)

    # Progress bar
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
    # progress_bar.grid(row=1, column=0, padx=10, pady=20, sticky=tk.W + tk.E)

    ttk.Label(root, text="Select Output Format:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
    format_var = tk.StringVar(value="Excel")
    format_dropdown = ttk.Combobox(root, textvariable=format_var, values=["Excel", "PDF"], state="readonly")
    format_dropdown.grid(row=1, column=0, padx=150, pady=10)

    # Generate report button
    ttk.Button(root, text="Generate Attendance Report", command=lambda: start_report_generation(file_listbox, progress_var, root, format_var)).grid(row=2, column=0, padx=10, pady=20)

    footer_label = ttk.Label(root, text="Developed and Maintained by Jay Patel", anchor=tk.CENTER)
    footer_label.grid(row=3, column=0, padx=10, pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
