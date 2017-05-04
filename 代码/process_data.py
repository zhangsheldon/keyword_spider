# -*- coding: utf-8 -*-
import  xdrlib ,sys
import xlrd
import xlwt
from xlutils.copy import copy
def main():
    xlsfile=xlrd.open_workbook("weiboData.xls")
    try:
        mysheet=xlsfile.sheet_by_name("Sheet1")
    except:
        print("no sheet in")
        return
    print("%d rows,%d cols"%(mysheet.nrows,mysheet.ncols))
   #f=xlwt.Workbook()
   # sheet_1=f.add_sheet('sheet_1',cell_overwrite_ok=True)
   # f.save("./weibo_data.xls")
    oldwb = xlrd.open_workbook('./weibo_data.xls')
    newwb = copy(oldwb)
    newws = newwb.get_sheet(2)
    num_row=0
    for row in range(1,mysheet.nrows):
        #tmp=""
        tmp=mysheet.cell(row,7).value
        pos=tmp.find("招商")
        #print(pos)
        if(pos>0):
            newws.write(num_row,0,tmp)
            num_row+=1
            #print("num_row",num_row)
    newwb.save('./weibo_data.xls')

main()
