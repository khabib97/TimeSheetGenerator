import xlsxwriter
import calendar
import datetime
from colorama import Fore, Back, Style 

print('Time & Salary Sheet Generator\nStrengthening of BGD e-GOV CIRT\n')
print('NOTE : Based on your information, this application generate your time and salary sheet in excel format.\nSo, you should provide correct information. You can update file after generation.\nSubmit your issue: https://github.com/khabib97/TimeSheetGenerator/issues\n')

CanRun = True
while CanRun:
    try:
        name= input('Your Name: ')
        designation = input('Your Designation: ')

        start = input('Start Date(DD-MM-YYYY): ')
        end = input('End Date(DD-MM-YYYY): ')

        salary = float(input('Your Salary(BDT Amount): '))
        vat = float(input('Your VAT(%): '))
        vat = (salary*vat)/100
        ati = float(input('Your ATI(%): '))
        ati = (salary*ati)/100 
        account_number = input('Your Bank Account Number: ')
        CanRun = False
    except ValueError:
        print("Sorry, Please provide correct information!")
        CanRun = True
        continue


start = start.split('-')
start_date = int(start[0])
start_month = int(start[1])
start_year = int(start[2])

end = end.split('-')
end_date = int(end[0])
end_month = int(end[1])
end_year =  int(end[2])

start_month_name = datetime.date(start_year, start_month , 1).strftime('%B')
end_month_name = datetime.date(end_year, end_month , 1).strftime('%B')

excle_file_name = 'Daily activities of '+ name +' '+ end_month_name +' '+ str(end_year) +'.xlsx'

workbook = xlsxwriter.Workbook(excle_file_name) 
# The workbook object is then used to add new  
# worksheet via the add_worksheet() method. 
worksheet = workbook.add_worksheet() 

worksheet.set_column('B:G', 14)

#date_formater = workbook.add_formater({'num_formater': 'mmmm d yyyy'})

banner_formater = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'text_wrap':'true'
})

info_formater = workbook.add_format({
    'border' : 1
})

date_formater = workbook.add_format({
    'border': 1,
    'align': 'center'
})

date_formater_holiday = workbook.add_format({
    'border': 1,
    'align': 'center',
    'bold': 1
})

internal_header_formater = workbook.add_format({
    'border' : 1,
    'bold': 1,
    'align': 'center'
})

data_formater = workbook.add_format({
    'border' : 1,
    'text_wrap':'true'
})

data_formater_holiday = workbook.add_format({
    'border' : 1,
    'bold': 1,
})

bold_formater = workbook.add_format({
    'border' : 1,
    'bold' : 1
})

# Use the worksheet object to write  
worksheet.merge_range('B1:G5','Government of the People\'s Republic of Bangladesh\nOffice of the Project Director\nStrengthening of BGD e-GOV CIRT\nICT Division, Ministry of Posts, Informaterion and Communication Technology\nICT Tower, Agargagaon, Dhaka-1207\n',banner_formater)
worksheet.merge_range('B6:G6', 'For the month: '+ str(start_date)+' '+ start_month_name +' '+ str(start_year)+ ' to '+ str(end_date) +' '+end_month_name+' '+str(end_year),info_formater) 
worksheet.merge_range('B7:G7', 'Name of the Consultants : '+ name,info_formater) 
worksheet.merge_range('B8:G8', 'Designation : '+ designation,info_formater) 
worksheet.merge_range('B9:G9', '',info_formater) 

#Table header
worksheet.write('B10','Date', internal_header_formater)
worksheet.write('C10','Day',internal_header_formater)
worksheet.merge_range('D10:F10','Activity',internal_header_formater)
worksheet.write('G10','Remark',internal_header_formater)

def day_week(year,month,day):
    date = datetime.date(year,month,day)
    return date.strftime('%A')


def row_generator(year,month,day,row):
    day_week_name = day_week(year,month,day)
    if day_week_name == 'Friday' or day_week_name == 'Saturday' :
        worksheet.write('B'+str(row),day,date_formater_holiday)
        worksheet.write('C'+str(row), day_week_name ,date_formater_holiday)
        worksheet.merge_range('D'+str(row)+':F'+str(row) ,'Weekly Holiday',data_formater_holiday)
    else:
        worksheet.write('B'+str(row),day,date_formater)
        worksheet.write('C'+str(row), day_week_name ,date_formater)
        worksheet.merge_range('D'+str(row)+':F'+str(row) ,'',data_formater)
    worksheet.write('G'+str(row),'',data_formater)

row = int(11)
#Starting month 
if start_month == end_month :
    for day in range(start_date, end_date+1):
        row_generator(end_year,end_month,day,row)
        row +=1
else:
    last_day_of_starting_month = calendar.monthrange(start_year,start_month)[1]
    for day in range(start_date, last_day_of_starting_month+1):
        row_generator(start_year,start_month,day,row)
        row +=1

    for day in range(1, end_date+1):
        row_generator(end_year,end_month,day,row)
        row +=1

row += 2
worksheet.merge_range('B'+str(row)+':C'+str(row+1),'Signature:',bold_formater)  
worksheet.merge_range('D'+str(row)+':G'+str(row+1),'',bold_formater)

row +=2
worksheet.merge_range('B'+str(row)+':C'+str(row),'Date',bold_formater)
worksheet.merge_range('D'+str(row)+':G'+str(row),'__________________________'+str(end_year),data_formater) 

# Finally, close the Excel file
# via the close() method. 
workbook.close() 

excle_file_name = 'Salary invoice of '+ name +' '+ end_month_name +' '+ str(end_year) +'.xlsx'

workbook = xlsxwriter.Workbook(excle_file_name) 
# The workbook object is then used to add new  
# worksheet via the add_worksheet() method. 
worksheet = workbook.add_worksheet() 

bold_formater = workbook.add_format({
    'bold' : 1
})

header_formater = workbook.add_format({
    'bold' : 1,
    'border': 1,
    'align': 'center'
})

data_formater = workbook.add_format({
    'border': 1
})

line_formater = workbook.add_format({
    'text_wrap':'true'
})

worksheet.merge_range('A1:G1','_______________________'+str(end_year)) 
worksheet.merge_range('A3:G3','The Project Director',bold_formater) 
worksheet.merge_range('A4:G4','Strengthening of BGD e-GOV CIRT')
worksheet.merge_range('A5:G5','Bangladesh Computer Council (BCC)')

worksheet.merge_range('A7:H7','Subject: Invoice for the Month of '+end_month_name+' '+str(end_year),bold_formater)

worksheet.merge_range('A9:J9', 'Name & Designation of the Consultant: '+ name + ', '+ designation)

worksheet.merge_range('A10:D10', 'Particular',header_formater )
worksheet.merge_range('E10:H10', 'Amount(BDT)',header_formater )

worksheet.merge_range('A11:D11', 'Net Payable',data_formater )
worksheet.merge_range('E11:H11', salary ,data_formater )

worksheet.merge_range('A12:D12', 'Less: VAT  deduction',data_formater )
worksheet.merge_range('E12:H12', vat ,data_formater )

worksheet.merge_range('A13:D13', 'Less: AIT deduction',data_formater )
worksheet.merge_range('E13:H13', ati ,data_formater )

worksheet.merge_range('A14:D14', 'Gross Payable',data_formater )
worksheet.merge_range('E14:H14', salary-vat-ati ,data_formater )

worksheet.merge_range('A16:G16', 'Net Payable: '+ str(salary-vat-ati) +' BDT only.',bold_formater)


worksheet.merge_range('A17:J19', 'This is to certify that the above statement is correct and the said bill has not been drawn earlier.\nIn any case of excess drawn the amount will be returned or adjusted with the subsequent bill. \nMonthly Activity Report is attached herewith.',line_formater)
worksheet.merge_range('A21:J21', 'It will be highly appreciated if you please pay the remuneration at the earliest.')

worksheet.merge_range('A24:G24', name, bold_formater)
worksheet.merge_range('A25:G25', designation)
worksheet.merge_range('A26:G26', 'Strengthening of BGD e-GOV CIRT')

worksheet.merge_range('A27:J27', '(Bank Information: Account No.:'+ account_number +', Janata Bank Limited, UGC Branch)	')

workbook.close() 

print('Done!\nPlease check your application directory for output files.')

input()

