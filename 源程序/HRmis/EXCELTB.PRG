*EXCELBB.prg
close all
clear
use d:\hrms\gz200410
oleapp=createob("excel.application")
if type('oleapp')='U' &&未能产生EXCEL对象
　　? '启动 EXCEL 失败...'
   release windows WaitWin
   return
endif
oleapp.visible=.f.
oleapp.workbooks.add
clear
startTime=time()
set cursor off
****** 等待 Excel 处理完毕后再继续
@1,4 say '已启动后台 Excel 表格处理程序,处理中...[退出=按下 Ctrl+End键]'
@3,4 say '开始时间:'
??starttime
go top
do while !eof()  &&每一行
 @ 10,10 say '正在将数据转换到 Excel 表格中...'+str(recno())
ii=recn()
for iii=1 to fcount() &&每一列
fld=field(iii)
oleapp.cells(ii,iii).value=&fld
*oleapp.range("a1:a5").select &&工资图表
oleapp.cells(1,1).select
endfor
&&VFP会在此等待，除非你从 EXCEL中退出或按下 Ctrl+End键 
if inkey(0.5)=23 
   exit
endif
skip
enddo
@5,4 say '现在时间:'
??time()
&&VFP会在此等待，除非你从 EXCEL中退出或按下 Ctrl+End键 
set cursor on
*oleapp.charts.add
*oleapp.activechart.charttitle.characters.text='工资图表'
oleapp.ActiveWindow.windowstate=2
oleapp.ActiveWorkbook.save()
oleapp.visible=.t.
close all
retu

************
proc abc
************
*第一步：转换
Select 部门,iif(性别='男',1,0) as 男,iif(性别='女',1,0) as 女,iif(文化程度='大学本科',1,0) as 大学本科,iif(文化程度='大学专科',1,0) as 大学专科,iif(inlist(职务,'工程师','技术员','总工程师'),1,0) as 技术人员,;
iif(inlist(职务,'总经理','经理','销售经理'),1,0) as 管理人员,;
iif(inlist(职务,'总经理','经理','销售经理'),1,0) as 管理人员,;
iif(inlist(职务,'秘书'),1,0) as 秘书;
from 程序1 into cursor query1

*第二步：转换并计算
Select 部门 as 部门,;
sum(男) as 男,;
sum(女) as 女,;
sum(大学本科) as 大学本科,;
sum(大学专科) as 大学专科,;
sum(技术人员) as 技术人员;
from query1 into cursor query1;
group by 部门
*往下面我们应该说说怎样把现在统计出的数据输出到 EXCEL 中了（当然什么样子的数据都可以）。
*首先用 CreateObject 函数建立一个 Excel automation 对象，程序如下：

local excel,abc
*define windows WaitWin at 0,0 size 20,60 system ;
*font 'times new roman',12 ;
*title '提示' close float zoom &&用于处理提示信息的窗口
*move windows waitWin center
*acti windows waitwin
*set color to w+/n
*clear

excel=createobject('excel.application')
if type('excel')='U' &&未能产生EXCEL对象
　　? '启动 EXCEL 失败...'
   release windows WaitWin
   return
endif

ReportTitle='XXXX公司人员分布情况'
ExcelFile="d:\camisbuild\xls1.xls"
excel.caption=ReportTitle
select query1
if  !file(ExcelFile)
   export to (ExcelFile) type xls
endif

excel.workbooks.open(ExcelFile)
abc=excel.activeworkbook.activesheet &&用对象代替这么长的一串字
abc.cells(1,1).value=ReportTitle 

*
* 第一行是标题,故从第二行开始
*
scan for recno()>1 &&每一行
 @ 10,10 say '正在将数据转换到 Excel 表格中...'+str(recno())
for i=1 to fcount() &&每一列
    fld=field(i)
    if type(fld)='C'&&如果字段太宽，Excel可能不能正确接受数据
abc.cells(recno()+3,i).value=trim(&fld)
*       else
*abc.cells(recno()+3,i).value=field(i)　
       endif
endfor
endscan

abc.cells(1,1).select
excel.ActiveWindow.windowstate=2
excel.ActiveWorkbook.save()
excel.visible=.t.
clear
startTime=time()
set cursor off

*
* 等待 Excel 处理完毕后再继续
*
do while .t.
@1,4 say '已启动后台 Excel 表格处理程序,处理中...'
@2,4 say 'Ctrl+End 继续处理...'
@4,4 say '开始时间:'
??starttime
@5,4 say '现在时间:'
??time()
&&VFP会在此等待，除非你从 EXCEL中退出或按下 Ctrl+End键 
if inkey(0.5)=23 or type('abc.name')='U'
   exit
endif
enddo
set cursor on
release abc,excel
release FindWindow
@ 10,10 say '处理完毕...'
=inkey(0.5)
release windows WaitWin
end note

 
