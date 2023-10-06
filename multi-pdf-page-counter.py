import os
from PyPDF2 import PdfReader
import xlsxwriter

path_to_pdf_directory = input(r"Enter the path of the folder containing pdf files: ")
files = os.listdir(path_to_pdf_directory)
pdf_files = [f for f in files if os.path.isfile(path_to_pdf_directory+'/'+f)] #Filtering only the files.

workbook = xlsxwriter.Workbook('pdf_pages_count.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'File Name')
worksheet.write('B1', 'Number of Pages')

row_cursor = 2
file_check_count = 1

for file in pdf_files:
    file_name, extension = os.path.splitext(file)
    if extension=='.pdf':
        reader = PdfReader(path_to_pdf_directory+'/'+file)
        number_of_pages = len(reader.pages)
        print(file, number_of_pages, sep=' => ')
        print("Files Checked",file_check_count, sep=" : ")
        file_name_cell = 'A'+str(row_cursor)
        number_of_pages_cell = 'B'+str(row_cursor)
        worksheet.write(file_name_cell,file)
        worksheet.write(number_of_pages_cell,number_of_pages)
        row_cursor += 1
        file_check_count += 1

workbook.close()