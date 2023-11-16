import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph

def convert_csv_to_pdf(csv_file):
    try:
        # Read CSV file into a pandas DataFrame
        df = pd.read_csv(csv_file)

        # Sort DataFrame based on 'Transaction_Date'
        df_sorted = df.sort_values(by=['Transaction_Date'])

        # Add a serial number column
        df_sorted['Sr_No'] = range(1, len(df_sorted) + 1)

        # Select specific columns (hardcoded for this example)
        selected_columns = ['Sr_No', 'Transaction_Date', 'Status', 'Amount', 'Customer_VPA']
        df_selected = df_sorted[selected_columns]

        # Reset the index after sorting
        df_selected.reset_index(drop=True, inplace=True)

        # Calculate the total amount
        total_amount = df_selected['Amount'].sum()

        # Create a row for the total amount
        total_row = ['Total', '', '', total_amount, '']

        # Concatenate the total row to the DataFrame
        df_with_total = pd.concat([df_selected, pd.DataFrame([total_row], columns=df_selected.columns)])

        # Determine the output PDF file name based on the input CSV file name
        pdf_file = os.path.splitext(csv_file)[0] + '_output.pdf'

        # Create a PDF file
        pdf = SimpleDocTemplate(pdf_file, pagesize=letter, topMargin=-15, bottomMargin=10)
        styles = getSampleStyleSheet()

        # Define a style for the header with background color
        header_style = ParagraphStyle(
            'Header1',
            parent=styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=10,  # Adjusted font size for header
            spaceAfter=6,
            textColor=colors.white,
            backgroundColor=colors.grey,
            spaceBefore=0,  # Adjusted space before for header
        )

        # Convert DataFrame to a list of lists
        data = [df_with_total.columns.tolist()] + df_with_total.values.tolist()

        # Calculate the number of rows per page
        rows_per_page = 40 if len(df_with_total) > 40 else len(df_with_total)

        # Create a list to hold the data for each page
        pages_data = [data[i:i + rows_per_page] for i in range(1, len(data), rows_per_page)]

        # Create a list to hold all paragraphs
        all_paragraphs = []

        # Iterate through pages and build the PDF document
        for page_data in pages_data:
            # Include the header row at the beginning of each page's data
            page_data_with_header = [data[0]] + page_data

            # Modify the Table object creation inside the loop
            table = Table(page_data_with_header, spaceBefore=0, spaceAfter=0)  # Adjusted spaceBefore for table

            # Apply styles to the table
            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # Header background color
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),  # Header text color
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),         # Center align all cells
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),        # Middle align all cells
                ('GRID', (0, 0), (-1, -1), 1, colors.black),   # Grid lines
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),  # Make the total row bold
                ('FONTSIZE', (0, 0), (-1, 0), 12),  # Adjusted font size for header
                ('FONTSIZE', (0, 1), (-1, -1), 10),  # Adjusted font size for table body
            ])

            table.setStyle(style)

            # Convert the table to a paragraph and add to the list
            all_paragraphs.append(Paragraph("Your Table Title", header_style))
            all_paragraphs.append(table)

        # Build the PDF document with all paragraphs
        pdf.build(all_paragraphs)

        messagebox.showinfo("Success", f"Conversion completed successfully. Output PDF: {pdf_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def browse_file(entry):
    filename = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, filename)

def run_conversion(csv_entry):
    csv_file = csv_entry.get()

    if not csv_file:
        messagebox.showerror("Error", "Please select a CSV file.")
        return

    print(f"CSV File: {csv_file}")

    convert_csv_to_pdf(csv_file)


# Create the main Tkinter window
window = tk.Tk()
window.title("CSV to PDF Converter")

# Create and place widgets in the window
tk.Label(window, text="CSV File:").grid(row=0, column=0, padx=10, pady=10)
csv_entry = tk.Entry(window, width=40)
csv_entry.grid(row=0, column=1, padx=10, pady=10)
btn_browse = tk.Button(window, text="Browse", command=lambda: browse_file(csv_entry))
btn_browse.grid(row=0, column=2, padx=10, pady=10)

# Replace the existing line for the "Convert" button with the following:
btn_convert = tk.Button(window, text="Convert", command=lambda: run_conversion(csv_entry))
btn_convert.grid(row=1, column=0, columnspan=3, pady=20)

# Start the Tkinter event loop
window.mainloop()
