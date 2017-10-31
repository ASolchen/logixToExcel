'''Converts RSLogix .csv to an xlsx with a chart'''
import xlsxwriter
import datetime as dt
from xlsxwriter.utility import xl_rowcol_to_cell
from tkFileDialog import askopenfilename
import sys



DATE_FORMAT = 'hh:mm:ss.000'
def excel_date(date1):
    temp = dt.datetime(1899, 12, 31)
    delta = date1 - temp
    return float(delta.days+1) + ((float(delta.seconds)+(float(delta.microseconds)/1000000)) / 86400)

def getCsvData(File):
    if debug:print('Opening {0}'.format(File))
    f=open(File, 'r')
    lines = f.readlines()
    f.close()
    lines.pop()
    for i in range(13):
        lines.pop(0) #  delete the first 13 lines of useless crap
    pens = len(lines[0].split(','))-3
    data = []
    for header in range(2,pens+3):
        data.append([lines[0].strip().split(',')[header].strip('\"'),[]])
    lines.pop(0)
    for line in range(len(lines)):
        if debug: print('line {0}'.format(line))
        for coulmn in range(pens+1):
            if coulmn > 0: val = float(lines[line].split(',')[coulmn+2].strip().strip('\"'))
            else:
                date = lines[line].split(',')[1].strip().strip('\"')
                time = lines[line].split(',')[2].strip().strip('\"').replace(';','.')
                date_time = dt.datetime(int(date.split('/')[2]),int(date.split('/')[0]), int(date.split('/')[1]),
                                        int(time.split(':')[0]),int(time.split(':')[1]),int(time.split(':')[2].split('.')[0]),
                                        int(time.split(':')[2].split('.')[1])*1000)
                val = excel_date(date_time)
            data[coulmn][1].append(val)                                        
    return data

def writeXLSX(data, xlsx):
    workbook = xlsxwriter.Workbook(xlsx)
    date_format = workbook.add_format({'num_format': DATE_FORMAT})
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})
    headings =[]
    for i in range(len(data)):
        headings.append(data[i][0])
    worksheet.write_row('A1', headings, bold)
    chart1 = workbook.add_chart({'type': 'line'})
    chartsheet = workbook.add_chartsheet()
    for i in range(len(data)):
        
        if i > 0 and data[i][0] != 'Speed_pc':
            worksheet.write_column(1, i, data[i][1])
            chart1.add_series({'values': ['Sheet1', 1, i, len(data[i][1]), i],
                               'name': ['Sheet1', 0, i],
                               'categories': ['Sheet1',1,0,len(data[i][1]),0],})
        elif i > 0 and data[i][0] == 'Speed_pc':
            worksheet.write_column(1, i, data[i][1])
            chart1.add_series({'values': ['Sheet1',1,i,len(data[i][1]),i],
                               'name': ['Sheet1', 0,i],
                               'categories': ['Sheet1',1,0,len(data[i][1]),0],
                               'y2_axis': 1 ,})
        else:
            worksheet.write_column(1, i, data[i][1], date_format)
    ##Create the chart
    chart1.set_x_axis({'date_axis':  True, 'num_format': DATE_FORMAT,})
    chart1.set_x_axis({'date_axis': True,'min': dt.date(2016, 6, 29),'max': dt.date(2016, 6, 29),})

  
    chart1.set_title ({'name': xlsx.strip('.xlsx').split('/')[-1]})
    chart1.set_x_axis({'name': 'Time'})
    chart1.set_y_axis({'name': 'Value'})
    chart1.set_y2_axis({'name': 'Speed %'})

    chart1.set_style(11)
    chartsheet.set_tab_color('#FF9900')
    chartsheet.set_chart(chart1)
    chart1.set_style(10)



    workbook.close()

    
def open_file_handler():
    file_opt = {}
    file_opt['defaultextension'] = '.csv'
    file_opt['filetypes'] = [('Rockwell CSV', '.csv'),]
    filePath= askopenfilename(**file_opt)
    filePath = filePath.replace('\\','/').lower()
    return filePath
            


if __name__=='__main__':
    debug = False
    csv = open_file_handler()

    if not csv: sys.exit(-1)
    xlsx = csv.replace('.csv', '.xlsx')
    data= getCsvData(csv)
    writeXLSX(data, xlsx)
    sys.exit(0)

    
    
        
