# -*- coding: cp936 -*-
import win32com, win32com.client, sys,os,os.path, time,datetime
t1 = datetime.datetime.now() #开始时间
print r'开始运行时间%s' % str(t1)
import traceback

backupdate1 = time.strftime("%Y-%m-%d")

time1 = time.strftime("%H:%M")

zhixingname = r'蓝水华'
jianchaname = r'龚诚刚'
#

#目录
path1 = r'D:\Aworkyun\作业计划\week\2G'
path2 = r'D:\Aworkyun\作业计划\week\2Gdata'
datapath = r'D:\Aworkyun\作业计划\week\2Gtoday'
filename = r'D:\Aworkyun\作业计划\week\2Gtoday\X1.xlsx'

#新建文件
xlsApp = win32com.client.Dispatch("Excel.Application")
try:
    xlsBook1 = xlsApp.Workbooks.Add()
    xlsSheet1 = xlsBook1.Sheets(1)
    xlsSheet2 = xlsBook1.Sheets(2)    
    xlsBook1.SaveAs(filename)

except Exception, e:
    print traceback.print_exc()
    print e



filename = ['01','02','03','04','05','06','07','08','09','51','52','53','54','55','56','57','58','59','60','61','63','64','65','66','67','68','69','70','71','72','73','74','75','76','77','78','79','90']

BscNum1 = ['BSC']
OmcNum1 = ['ZJ_ZJ_U2000(1)_G','ZJ_ZJ_U2000(2)_G','ZJ_ZJ_U2000(3)_G','ZJ_ZJ_U2000(4)_G','ZJ_HZ_OMC2000(6)','ZJ_HZ_OMC2000(7)','U2000-CE','华为U2000-1_L','华为U2000-2_L','诺西OMCR_L ','中兴OMCR_L']
#sumnum1 = 11,
#sumnum2 = 11
i = 2
z1 = 2
N = {}
HRC = {}
ARC = {}
PJMAX = {}
FZMAX = {}
PJMIN = {}
FZMIN = {}

i = 9
for CB in ['1','2','3']:                     
    CPUBook1 = xlsApp.Workbooks.Open(r'%s\%s.xlsx' % (datapath,CB))
    CPUSheet1 = CPUBook1.Sheets(1)

    print r'开始统计%s.xlsx' % CB
    while CPUSheet1.Cells(i,3).Value:
         
        BSCname = CPUSheet1.Cells(i,3).Value    
        xpupingjun = float(CPUSheet1.Cells(i,5).Value)
        xpuzuidazhi = float(CPUSheet1.Cells(i,6).Value)        

        if BSCname not in ARC:
            N[BSCname] = 0
            HRC[BSCname] = 0.0
            ARC[BSCname] = 0.0
            PJMAX[BSCname] = 0
            FZMAX[BSCname] = 0
            PJMIN[BSCname] = 100
            FZMIN[BSCname] = 100
            

        ARC[BSCname] = ARC[BSCname] + xpupingjun
        if xpupingjun > PJMAX[BSCname]:
            PJMAX[BSCname] = xpupingjun
        if xpupingjun < PJMIN[BSCname]:
            PJMIN[BSCname] = xpupingjun
        HRC[BSCname] = HRC[BSCname] + xpuzuidazhi
        if xpuzuidazhi > FZMAX[BSCname]:
            FZMAX[BSCname] = xpuzuidazhi
        if xpuzuidazhi < FZMIN[BSCname]:
            FZMIN[BSCname] = xpuzuidazhi
        N[BSCname] +=1
            
  
        
        i += 1
    i = 9
    
    
    
    xlsBook1.Save()
    CPUBook1.Close()



xlsSheet2.Cells(1,1).Value = u'BSC名称'
xlsSheet2.Cells(1,2).Value = u'平均值的最大值'
xlsSheet2.Cells(1,3).Value = u'平均值的最小值'
xlsSheet2.Cells(1,4).Value = u'平均值的平均值'
xlsSheet2.Cells(1,5).Value = u'最大值的最大值'
xlsSheet2.Cells(1,6).Value = u'最大值的最小值'
xlsSheet2.Cells(1,7).Value = u'最大值的平均值'

#

for fn in filename:
    BSCname = 'HZBSC%sHWE' % fn
    bsccpu2 = '%s%%' % PJMAX[BSCname]    
    xlsSheet2.Cells(z1,1).Value = BSCname
    xlsSheet2.Cells(z1,2).Value = PJMAX[BSCname]
    xlsSheet2.Cells(z1,3).Value = PJMIN[BSCname]
    xlsSheet2.Cells(z1,4).Value = '%.1f' % (ARC[BSCname]/N[BSCname])
    xlsSheet2.Cells(z1,5).Value = FZMAX[BSCname]
    xlsSheet2.Cells(z1,6).Value = FZMIN[BSCname]
    xlsSheet2.Cells(z1,7).Value = '%.1f' % (HRC[BSCname]/N[BSCname])
    z1 += 1

    bscxlsBook = xlsApp.Workbooks.Open(r'%s\BSC%s.xls' % (path1,fn))
    bscxlsSheet = bscxlsBook.Sheets(1)
    bscxlsSheet.Cells(6,4).Value = bsccpu2
    bscxlsSheet.Cells(3,2).Value = backupdate1
    bscxlsSheet.Cells(10,2).Value = zhixingname
    bscxlsSheet.Cells(11,2).Value = jianchaname
    bscxlsBook.Save()
    bscxlsBook.Close()


PJZZDZ = xlsSheet2.Cells(2,2).Value
ZDZZDZ = xlsSheet2.Cells(2,5).Value
PJBSCNAME = xlsSheet2.Cells(2,1).Value
ZDBSCNAME = xlsSheet2.Cells(2,1).Value
hangshu = 3
while xlsSheet2.Cells(hangshu,1).Value:
    if xlsSheet2.Cells(hangshu,2).Value >PJZZDZ:
        PJZZDZ = xlsSheet2.Cells(hangshu,2).Value
        PJBSCNAME = xlsSheet2.Cells(hangshu,1).Value
    if xlsSheet2.Cells(hangshu,5).Value >ZDZZDZ:
        ZDZZDZ = xlsSheet2.Cells(hangshu,5).Value
        ZDBSCNAME = xlsSheet2.Cells(hangshu,1).Value
    hangshu += 1
print u'平均值的最大值：%s %s' % (PJBSCNAME,PJZZDZ)
print u'最大值的最大值：%s %s' % (ZDBSCNAME,ZDZZDZ)

xlsBook1.Save()
xlsBook1.Close()

for i2 in BscNum1:
    xlsBook1 = xlsApp.Workbooks.Open(r'%s\BACKUP_%s.xls' % (path1,i2))
    xlsSheet1 = xlsBook1.Sheets(1)
    xlsSheet1.Cells(3,4).Value = backupdate1
    xlsSheet1.Cells(6,4).Value = time1
    xlsSheet1.Cells(8,2).Value = zhixingname
    xlsSheet1.Cells(9,2).Value = jianchaname
    xlsBook1.Save()
    xlsBook1.Close()
    xlsBook2 = xlsApp.Workbooks.Open(r'%s\ZJ_HZ_A_%s.xls' % (path1,i2))
    xlsSheet2 = xlsBook2.Sheets(1)
    xlsSheet2.Cells(3,4).Value = backupdate1
    xlsSheet2.Cells(8,2).Value = zhixingname
    xlsSheet2.Cells(9,2).Value = jianchaname
    xlsBook2.Save()
    xlsBook2.Close()
xlsBook3 = xlsApp.Workbooks.Open(r'%s\机房清洁整理.xls' % path1)
xlsSheet3 = xlsBook3.Sheets(1)
xlsSheet3.Cells(3,5).Value = backupdate1
xlsSheet3.Cells(5,3).Value = backupdate1
xlsSheet3.Cells(6,3).Value = backupdate1
xlsSheet3.Cells(5,4).Value = zhixingname
xlsSheet3.Cells(6,4).Value = zhixingname
xlsSheet3.Cells(10,2).Value = zhixingname
xlsSheet3.Cells(11,2).Value = jianchaname
xlsBook3.Save()
xlsBook3.Close()


#OMC
for i2 in OmcNum1:
    xlsBook1 = xlsApp.Workbooks.Open(r'%s\%s.xls' % (path1,i2))
    xlsSheet1 = xlsBook1.Sheets(1)
    xlsSheet1.Cells(3,4).Value = backupdate1
    xlsSheet1.Cells(17,2).Value = zhixingname
    xlsSheet1.Cells(18,2).Value = jianchaname
    xlsBook1.Save()
    xlsBook1.Close()

#重命名文件夹和新建文件夹


os.rename(('%s' % datapath),(r'%s\%s' % (path2,backupdate1)))
os.makedirs('%s' % datapath)


xlsApp.Quit()

t2 = datetime.datetime.now() #结束时间
print r'结束运行时间%s' % str(t2) 
print r'程序运行总共耗时%s秒' % (t2 - t1).seconds


