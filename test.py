'''import aspose.pdf as ap
import pandas as pd

    # Initialize the Document object by calling its empty constructor
input_file = "input_data.xlsx"

# Load Excel data
df = pd.read_excel(input_file)
# Print the remaining column headers
columns = df.columns.tolist()



pdf_document = ap.Document()
pdf_document.pages.add()
    # Initializes a new instance of the Table
table = ap.Table()
    # Set the table border color as LightGray
#table.border = ap.BorderInfo(ap.BorderSide.ALL, 0.5, ap.Color.black)
    # Set the border for table cells
#table.default_cell_border = ap.BorderInfo(ap.BorderSide.ALL, 0.5, ap.Color.black)
    # Add 1st row to table

# Iterate through each row in the DataFrame
data1 = ""
data2 = ""
for index, excel_row in df.iterrows():
    row = table.rows.add()
    print(index)
    #if(index < 3):
    for col in columns:
        data1 = data1 + str(df.loc[index])
            #data2 = data2 + str(df.loc[index+1])
        print(data1)
            #row.cells.add(data1)
            #row.cells.add(data2)
    data1 = ""
    data2 = ""
for cellCount in range(1, 5):
        # Add table cells
        row = table.rows.add()
        row.cells.add("Test" + str(cellCount))
        row.cells.add("Test" + str(cellCount+1))
 # Add table object to first page of input document
pdf_document.pages[1].paragraphs.add(table)
    # Save updated document containing table object
pdf_document.save("test.pdf")'''

import tkinter as tk
window = tk.Tk()
label = tk.Label(window, text = "Hello World!").pack()
window.mainloop()

