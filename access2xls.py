# -*- coding: utf-8 -*- 
#python 2.7.12
#access to xls in windows
import win32com.client, string
from xlrd import *
from xlutils import *
from xlwt import *

rows = 0
sheetNum = 1
wb = Workbook()
ws = wb.add_sheet('1')

def addXls(item, fieldNum):
    global rows, sheetNum, ws

    if rows >= 65000:                   #xlwt max rows 65535
        rows = 0
        sheetNum += 1
        ws = wb.add_sheet(str(sheetNum))
        
    #print 'rows:', rows
    for i in range(0, fieldNum):
        ws.write(rows, i, item(i).value)
        
    rows += 1

def printXlsResult():
    data = open_workbook('./test.xls')  
    table = data.sheet_by_index(0) 
    print 'xls rows:%d, cols:%d' % (table.nrows, table.ncols)

def main():
    conn = win32com.client.Dispatch(r'ADODB.Connection') 
    DSN = 'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=./test.mdb;'
    conn.Open(DSN)

    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    rs_name = 'table'   
    rs.Open('[' + rs_name + ']', conn, 1, 3) 

    rs.MoveFirst()   
    count = 0   
    while 1:   
        if (count % 1000) == 0:
            print "count:", count

        if rs.EOF:
            break   
        elif count == 5000:
            break
        else:          
            line = ''
            for i in range(0, rs.Fields.Count):
                line = line + str(rs.Fields.Item(i).Value) + " "
            print line
                
            addXls(rs.Fields.Item, rs.Fields.Count)
            count += 1
            rs.MoveNext() 

    print "count:", count
    global wb
    wb.save('./test.xls')
    printXlsResult()


if __name__=="__main__":
    main()

