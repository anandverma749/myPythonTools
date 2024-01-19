import os
from PIL import Image
import xlsxwriter

dir_list = os.listdir(os.getcwd())
dirs = [f for f in dir_list if os.path.isdir(os.getcwd()+'/'+f)] #Filtering only the folders.

workbook = xlsxwriter.Workbook('Compression_Details.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Batch Name')
worksheet.write('B1', 'Compression')

row_cursor = 2
file_check_count = 1

for batch in dirs:
    files = os.listdir(os.getcwd() + '/' + batch)
    tif_files = [f for f in files if os.path.isfile(os.getcwd() + "/" + batch +'/'+f) and f.endswith(".tif")] 
    if len(tif_files):
        img = Image.open(batch + '/' + tif_files[0])
        compression = img.info["compression"]
        batch_name_cell = 'A'+str(row_cursor)
        compression_type_cell = 'B'+str(row_cursor)
        worksheet.write(batch_name_cell, batch)
        worksheet.write(compression_type_cell, compression)
        print("Folders Checked",file_check_count, sep=" : ")
        
    else:
        batch_name_cell = 'A'+str(row_cursor)
        compression_type_cell = 'B'+str(row_cursor)
        worksheet.write(batch_name_cell, batch)
        worksheet.write(compression_type_cell, "No .tif file present")
        print("Folders Checked",file_check_count, sep=" : ")
        
    row_cursor += 1
    file_check_count += 1

workbook.close()