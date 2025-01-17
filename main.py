from openpyxl import load_workbook
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import math
import os
from datetime import datetime

def excel_to_pdf():
    try:
        # Create output directory
        os.makedirs('out-files', exist_ok=True)
        
        # Hardcoded input path
        excel_file_path = "input.xlsx"
        
        # Generate output PDF filename with current date
        current_date = datetime.now().strftime("%Y-%m-%d")
        output_pdf_path = os.path.join('out-files', f'{current_date}.pdf')
        
        wb = load_workbook(excel_file_path)
        ws = wb.active
        
        # Get dimensions of the current excel file  
        max_row = ws.max_row
        max_col = ws.max_column
        
        # PDF document setup with landscape letter size and elements list   
        doc = SimpleDocTemplate(output_pdf_path, pagesize=landscape(letter))
        elements = []
        
        # Calculate how many columns can fit on one page
        tuner_columns_number = 8
        available_width = doc.width
        col_width = available_width / tuner_columns_number 
        cols_per_page = min(tuner_columns_number, max_col)
        
        # Calculate number of batches needed for the current excel file
        num_batches = math.ceil(max_col / cols_per_page)
        
        # Calculate number of row batches needed
        rows_per_page = 30  # Adjust this value based on your needs
        num_row_batches = math.ceil(max_row / rows_per_page)
        
        page_counter = 1
        
        for row_batch in range(num_row_batches):
            start_row = row_batch * rows_per_page + 1
            end_row = min((row_batch + 1) * rows_per_page, max_row)
            
            for col_batch in range(num_batches):
                start_col = col_batch * cols_per_page + 1
                end_col = min((col_batch + 1) * cols_per_page, max_col)
                
                # Create data for current batch
                data = []
                
                # Add header row for current batch
                header_row = []
                for col in range(start_col, end_col + 1):
                    cell_value = ws.cell(row=1, column=col).value
                    header_row.append(cell_value)
                data.append(header_row)
                
                # Add data rows for current batch   
                for row in range(max(2, start_row), end_row + 1):
                    data_row = []
                    for col in range(start_col, end_col + 1):
                        cell_value = ws.cell(row=row, column=col).value
                        data_row.append(cell_value)
                    data.append(data_row)
                
                # Create table for current batch mapping to the data and style it   
                table = Table(data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 10),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                
                elements.append(table)
                elements.append(Spacer(1, 20))
                
                # Add page number
                if col_batch == 0:
                    page_num = str(page_counter)
                else:
                    page_num = f"{page_counter}.{col_batch}"
                    
                page_number = Paragraph(f"Page {page_num}", getSampleStyleSheet()['Normal'])
                elements.append(page_number)
                elements.append(PageBreak())
            
            page_counter += 1
        
        # Build PDF
        doc.build(elements)
        return True
        
    except Exception as e:
        print(f"Error converting Excel to PDF: {str(e)}")
        return False

if __name__ == "__main__":
    excel_to_pdf()
