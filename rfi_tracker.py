from xlrd import open_workbook
import xlwt
from xlutils.copy import copy
import arcpy
import subprocess
def main():
    
    #SET STYLE FOR EACH ROW/COLUMN THAT IS ADDED TO SPREADSHEET
    style = xlwt.easyxf(
        'font: name Calibri, color-index black, bold off;'
        'borders:left thin, right thin, top thin, bottom thin;',
            num_format_str='#,##0')

    #OPEN WORKBOOK USING XLRD 
    book = open_workbook('D:\Temp\TaskTracker.xls')

    #IDENTIFY WHICH SHEET TO MANIPULATE (FIRST SHEET)
    sh = book.sheet_by_index(0)

    #IDENTIFY THE ROW TO WRITE THE NEW RECORD (THE FIRST EMPTY ROW AT THE BOTTOM, AFTER THE LAST RECORD)
    row_to_write = sh.nrows 

    #REOPEN WORKBOOK USING XLRD AGAIN (NEEDED TO DO TWICE IN ORDER TO GET AN ACCURATE ROW_TO_WRITE VALUE ABOVE)
    book = open_workbook('D:\Temp\TaskTracker.xls',formatting_info = True)

    #USE XLUTILS LIBRARY TO COPY THE BOOK
    wb = copy(book)
    sheet = wb.get_sheet(0)

    #OBTAINS THE CELL VALUE IMMEDIATELY PRECEDING THE CELL TO WRITE THE CORRECT NEXT RECORD (EX IF THE LAST RECORD IS 77, THE NEXT ONE WILL BE 78) 
    value = sh.cell(row_to_write - 1,0).value

    #GET VARIABLES FROM THE USER
    title = arcpy.GetParameterAsText(0)
    request_date = arcpy.GetParameterAsText(1)
    due = arcpy.GetParameterAsText(2)
    requester = arcpy.GetParameterAsText(3)
    region = arcpy.GetParameterAsText(4)
    ga = arcpy.GetParameterAsText(5)
    rfi = arcpy.GetParameterAsText(6)
    map_size = arcpy.GetParameterAsText(7)
    number_of_maps = arcpy.GetParameterAsText(8)
    hard_soft = arcpy.GetParameterAsText(9)

    #WRITE VARIABLES TO THE WORKSHEET
    sheet.write(row_to_write, 0, value+1, style)
    sheet.write(row_to_write, 1, title, style)
    sheet.write(row_to_write, 2, request_date, style)
    sheet.write(row_to_write, 3, due, style)
    sheet.write(row_to_write, 4, requester, style)
    sheet.write(row_to_write, 5, region, style)
    sheet.write(row_to_write, 6, ga, style)
    sheet.write(row_to_write, 7, rfi, style)
    sheet.write(row_to_write, 8, map_size, style)
    sheet.write(row_to_write, 9, number_of_maps, style)
    sheet.write(row_to_write, 10, hard_soft, style)

    #RESAVE THE WORKBOOK
    wb.save('D:\Temp\TaskTracker.xls')
    
    #SPLIT GA INTO FIRST AND LAST NAME
    ga_lastname = string.split(ga, ' ')

    #ADD VARIABLE VALUES TO FILE
    rfi_file = open("D:\\RFIfile.txt", 'w')
    rfi_file.write("Map Title: "+title+'\n')
    rfi_file.write("Request Date: " +request_date[:10]+'\n')
    rfi_file.write("Due Date: "+due[:10]+'\n')
    rfi_file.write("Requestor: "+requester+'\n')
    rfi_file.write("Region: "+region+'\n')
    rfi_file.write("Tasked GA: "+ga_lastname[1]+'\n')
    rfi_file.write("RFI #: "+rfi+'\n')
    rfi_file.write("Map Size: "+map_size+'\n')
    rfi_file.write("Number of Maps: "+number_of_maps+'\n')
    rfi_file.write("Hard or Soft Copy: "+hard_soft+'\n')
    rfi_file.close()

    #RESAVE THE WORKBOOK
    wb.save('D:\Temp\TaskTracker.xls')

    #EMAIL GENERATION
    outlookpath2doc = '"C:/Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE"'
    compose = '/c ipm.note'
    recipients = '/m "' + "ProjectManager@example.com; deputyPM@example.com&subject=New project added to the tracker &body=PM/Deputy PM,\n\nAttached is the RFI/relevant information associated with my latest project.\n\nRespectfully,\n\n" +ga+ '"'
    attachment = '/a "D:\\RFIfile.txt"'
    command = ' '.join([outlookpath2doc, compose, recipients, attachment])
    process = subprocess.Popen(command, shell=False, stdout=subprocess.PIPE)

main()  
