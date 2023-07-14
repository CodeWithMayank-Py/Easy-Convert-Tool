from fpdf import FPDF
import pandas as pd
import csv
import openpyxl

# Button 1 : Text to PDF Format

def convert_text_to_pdf(input_file, output_file):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    with open(input_file, "r") as file:
        text = file.read()
        pdf.multi_cell(0, 10, txt=text)
    pdf.output(output_file)




# Button 2 : CSV to Excel Format


def convert_csv_to_excel(csv_file, excel_file):
    try:
        df = pd.read_csv(csv_file)
        writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
        df.to_excel(writer, index=False)
        writer.save()
        return True
    except Exception as e:
        print(f"Error converting CSV to Excel: {str(e)}")
        return False



# Button 3 :  Excel to CSV format

def convert_excel_to_csv(excel_file, csv_file):
    try:
        wb = openpyxl.load_workbook(excel_file)
        sheet = wb.active

        with open(csv_file, 'w', newline='') as file:
            csv_writer = csv.writer(file)
            for row in sheet.iter_rows():
                csv_writer.writerow([cell.value for cell in row])

        return True
    except Exception as e:
        print(f"Error converting Excel to CSV: {str(e)}")
        return False