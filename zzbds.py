import sys

import os

import calendar
import time
import xlwings as xw

with open('练习.txt',encoding='utf-8') as f:
    a=f.read().split()
    mo=a[0][0]
    b=a[0].split("/")
    date=[]#日期
    dch=[]#订仓号
    gh=[]#柜号
    fth=[]#封条号
    work=[]#提柜还柜装货合体版
    dd=[]#打单费
    gbf=[]#过磅费
    bz=[]#备注
    for c in a:
        c=c.split("/")
        date.append(c[0]+"月"+c[1]+"日")
        dch.append(c[2])
        gh.append(c[3])
        fth.append(c[4])
        work.append(c[5]+"-"+c[6]+"-"+c[7])
        dd.append(c[8])
        gbf.append(c[9])
        bz.append(c[10])

    dn = time.localtime()
    yc = '%d/%02d/01' % (dn.tm_year, int(mo))
    ymd = calendar.monthrange(dn.tm_year, int(mo))
    ym = '%d/%02d/%02d' % (dn.tm_year, int(mo), ymd[1])
    app = xw.App(visible=False, add_book=False)
    app.display_alerts=False
    name=str(dn.tm_year)+"出车记录.xls"
    app1=xw.App(visible=False, add_book=False)
    wb1 = app1.books.open(r'test.xls')
    b1 = wb1.sheets['example']
    if os.path.exists(name):
       wb=xw.Book(name)
    else:
       wb = app.books.add()
    try:
        b = wb.sheets.add( str(dn.tm_year)+ "-" + mo)
    except:
        b = wb.sheets[str(dn.tm_year) + "-" + mo]
    b1.range('A1:I23').api.Copy()
    b.range("A1").api.Select()
    b.api.Paste()
    b.range('C1').options(transpose=True).value = yc + "--" + ym
    b.range('B3').options(transpose=True).value = date
    b.range('C3').options(transpose=True).value = dch
    b.range('D3').options(transpose=True).value = gh
    b.range('E3').options(transpose=True).value = fth
    b.range('F3').options(transpose=True).value = work
    b.range('G3').options(transpose=True).value = gbf
    b.range('H3').options(transpose=True).value = dd
    b.range('I3').options(transpose=True).value = bz
    b.range('A1:I23').rows.autofit()
    b.range('A1:I23').columns.autofit()
    wb.save(name)
    wb.close()
    wb1.close()
    app.quit()
    app1.quit()
