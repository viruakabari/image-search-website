import os
import pandas as pd
import barcode
from barcode.writer import ImageWriter
from fpdf import FPDF
import tkinter as tk
from tkinter import filedialog, Label, Button, messagebox

# Function to open file explorer and select Excel file
def browseFiles():
    filePath = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Excel files", "*.xlsx*"), ("All files", "*.*")))
    if filePath:
        label_file_explorer.config(text="File open - " + filePath)

# Function to generate PDF with barcodes
def generatePDF(currency_type):
    label_text = label_file_explorer.cget("text")

    # Ensure a file is selected
    if "File open - " not in label_text:
        messagebox.showerror("Error", "No file selected. Please select a file first.")
        return

    path = label_text.replace("File open - ", "").strip()

    if not os.path.exists(path):
        messagebox.showerror("Error", f"Selected file does not exist: {path}")
        return

    output_path = os.path.join(os.path.dirname(path), f"barcode_{currency_type}.pdf")
    barcode_dir_path = os.path.join(os.path.dirname(path), "barcodes")

    # Create barcode directory if it does not exist
    os.makedirs(barcode_dir_path, exist_ok=True)

    # Load Excel data
    df = pd.read_excel(path)

    # Create a PDF object
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=10)

    # Print column headers (excluding 'barcode_number')
    columns = df.columns.tolist()
    if 'barcode_number' in columns:
        columns.remove('barcode_number')

    # Iterate through each row in the DataFrame
    for index, row in df.iterrows():
        for col in columns:
            if col == 'retail_price':
                barcode_data = str(row['barcode_number']).strip()

                # Ensure barcode data is 12 digits long (EAN-13 requirement)
                barcode_data = barcode_data.zfill(12)

                try:
                    # Generate barcode and save
                    barcode_img = barcode.get('ean13', barcode_data, writer=ImageWriter())
                    barcode_path = os.path.join(barcode_dir_path, f"barcode_{barcode_data}")
                    barcode_img.save(barcode_path)
                    barcode_path_with_extension = f"{barcode_path}.png"

                    # Add barcode image to PDF
                    if os.path.exists(barcode_path_with_extension):
                        pdf.image(barcode_path_with_extension, x=3, y=pdf.get_y(), w=60, h=10)
                        pdf.set_y(pdf.get_y() + 12)  # Move down after barcode
                    else:
                        print(f"Error: Barcode image not found at {barcode_path_with_extension}")

                except Exception as e:
                    print(f"Error generating barcode for {barcode_data}: {e}")
                    continue

            # Add pricing based on currency selection
            elif col == 'price':
                if currency_type == "MRP":
                    price_text = f"MRP: ₹{row[col]}"  # Properly display rupee symbol
                else:
                    price_text = f"Price: ${row[col]}"  # Dollar pricing
                
                pdf.cell(140, 5, price_text, border=0)
                pdf.ln()
            else:
                pdf.cell(140, 5, str(row[col]), border=0)
                pdf.ln()

        pdf.ln(10)  # Add spacing between rows

    # Save the PDF
    try:
        pdf.output(output_path, "F")
        messagebox.showinfo("Success", f"PDF created successfully: {output_path}")
        print(f"✅ PDF created successfully: {output_path}")
    except PermissionError:
        messagebox.showerror("Error", f"Permission denied: {output_path}. Close the file if it is open and try again.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create PDF: {e}")

# Create the root window
window = tk.Tk()
window.title('Barcode Generator')
window.geometry("600x400")
import os
import pandas as pd
import barcode
from barcode.writer import ImageWriter
from fpdf import FPDF
import tkinter as tk
from tkinter import *
from tkinter import filedialog


# Function for opening the 
# file explorer window

def browseFiles():
    filePath = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select a File",
                                          filetypes = (("excel files",
                                                        "*.xlsx*"),
                                                       ("all files",
                                                        "*.*")))
    # Change label contents
    label_file_explorer.configure(text="File open -"+filePath)
    
def generatePDF():
    input_file = "Type 1.xlsx"
    output_file = "\\barcode.pdf"

    # Directory to save barcode images
    barcode_dir = "\\barcodes"
    label_text = label_file_explorer.cget("text")
    path = label_text.split('-')[1]
    output_path = os.path.dirname(os.path.abspath(path)) + output_file
    barcode_dir_path = os.path.dirname(os.path.abspath(path)) + barcode_dir
    os.makedirs(barcode_dir_path, exist_ok=True)
    print("barcode_dir "+barcode_dir_path)
    print("FilePath "+path)

    # Load Excel data
    df = pd.read_excel(path)
    print(df)

    # Create a PDF object
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    # Add a page and set font
    pdf.add_page()
    pdf.set_font("Arial", size=8)

    # Print the remaining column headers
    columns = df.columns.tolist()
    columns.remove('barcode_number')  # Remove barcode_number column from the list

    # Iterate through each row in the DataFrame
    for index, row in df.iterrows():
        for col in columns:
            # If the current column is retail_price, insert the barcode image first
            if col == 'retail_price':
            # Generate a barcode for the barcode_number column
                barcode_data = str(row['barcode_number']).strip()

                # Ensure barcode data is 12 digits long (for EAN-13)
                if len(barcode_data) < 12:
                    barcode_data = barcode_data.zfill(12)  # Pad barcode number to 12 digits

                try:
                    # Use 'ean13' format for barcode (EAN-13 requires 12 digits)
                    barcode_img = barcode.get('ean13', barcode_data, writer=ImageWriter())

                    # Create the barcode image path inside 'barcodes' directory
                    barcode_path = os.path.join(barcode_dir_path, f"barcode_{barcode_data}")

                    # Save the barcode image (ImageWriter will add the .png extension)
                    barcode_img.save(barcode_path)

                    # Add .png explicitly for PDF inclusion
                    barcode_path_with_extension = f"{barcode_path}.png"

                    # Check if the image is successfully saved
                    if os.path.exists(barcode_path_with_extension):
                        print(f"Barcode image saved at: {barcode_path_with_extension}")

                        # Add the barcode image to the PDF with adjusted size and alignment
                        pdf.image(barcode_path_with_extension, x=3, y=pdf.get_y(), w=60, h=5)  # Fixed left margin alignment
                        pdf.set_y(pdf.get_y() + 5)  # Move down after the barcode
                    else:
                        print(f"Error: Barcode image not saved at {barcode_path_with_extension}")
                        label_error.configure(text=f"Error: Barcode image not saved at {barcode_path_with_extension}") # type: ignore
                except Exception as e:
                    print(f"Error generating barcode for {barcode_data}: {e}")
                    label_error.configure(text=f"Error generating barcode for {barcode_data}: {e}") # type: ignore
                    continue

            # Add other column headers and data
            else:
                header_data = f"{row[col]}"

                pdf.cell(140, 4, header_data, border=0)
                pdf.ln()

        # Add a blank line after completing one set (one row of data)
        pdf.ln(10)

    # Save the PDF file
    try:
        print("Output path "+output_path)
        pdf.output(output_path)
        print(f"PDF created successfully with barcodes: {output_file}")
    except Exception as e:
        print(f"Error creating PDF: {e}")


# Create the root window
window = Tk()
  
# Set window title
window.title('File Explorer')
  
# Set window size
window.geometry("800x500")
  
#Set window background color
window.config(background = "white")

# Create a File Explorer label
label_file_explorer = Label(window, 
                            text = "File Explorer using Tkinter",
                            width = 100, height = 4, 
                            fg = "blue")
'''label_error = Label(window, 
                            text = "Error",
                            width = 100, height = 4, 
                            fg = "red")'''
button_explore = Button(window, 
                        text = "Browse Files",
                       command = browseFiles) 
button_convert = Button(window, 
                        text = "Convert",
                        command = generatePDF) 
  
#button_exit = Button(window, 
#                     text = "Exit",
#                     command = exit) 
  
# Grid method is chosen for placing
# the widgets at respective positions 
# in a table like structure by
# specifying rows and columns
label_file_explorer.grid(column = 1, row = 1)
  
button_explore.grid(column = 1, row = 2)
button_convert.grid(column = 1, row = 3)
#label_error.grid(column = 1, row = 4)
  
#button_exit.grid(column = 1,row = 4)

# Let the window wait for any events
window.mainloop()
window.config(background="white")

# Create UI elements
label_file_explorer = Label(window, text="Select an Excel file", width=80, height=4, fg="blue")
button_explore = Button(window, text="Browse Files", command=browseFiles)
button_convert_mrp = Button(window, text="Generate PDF (MRP)", command=lambda: generatePDF("MRP"))
button_convert_dollar = Button(window, text="Generate PDF ($)", command=lambda: generatePDF("Dollar"))

# Place UI elements in grid
label_file_explorer.grid(column=1, row=1)
button_explore.grid(column=1, row=2)
button_convert_mrp.grid(column=1, row=3)
button_convert_dollar.grid(column=1, row=4)

# Run Tkinter event loop
window.mainloop()
