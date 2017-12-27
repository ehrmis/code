**************************
* 缴费工资.prg(2008.1.7)
**************************
nd=str(year(date()),4)
clear
_SCREEN.PICTURE = ''
@ 1,36 say '报 盘 提 示' font '黑体',20
@ 4,10 say '（1）请先顺利完成上年工资总额计算工作！' font '宋体',14
@ 6,10 say '（2）自动导出上报社保部电子表“D:\申报基数.XLS”电子表！' font '宋体',14
@ 8,10 say '（3）请打开该电子表按要求编辑申报盘或ERP模板！！！' font '宋体',14
yn = MESSAGEBOX("准备好了吗？",4+32,"提示")
if yn=7
   clear
 _SCREEN.WINDOWSTATE = 2
close all
use data\ccrq
go 8
repl 图片 with str(val(图片)+1,2)
if allt(图片)='10'
   repl 图片 with '0'
endi   
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
_SCREEN.MAXBUTTON = .F.
_SCREEN.MOVABLE = .F.
_SCREEN.CONTROLBOX = .F.
_SCREEN.CLOSABLE = .F.
retu
endif
@ 14,10 say "请输入当前需要申报缴费基数的缴费年：" font '宋体',14 get nd font '黑体',18
read
sn=str(val(nd)-1,4)
if file('zr1'+sn+'.dbf')
use 申报基数 excl
zap
appe from zr1&sn
repl 车间 with cjmc,部门 with bmmc,姓名 with ryxm,erp编号 with erprybh,个人编号 with scbh,身份证号 with sfz,工资总额 with hj,月平均 with pj,缴费基数 with jfjs,养老保险 with zr3 all

copy to D:\申报基数 fiel rybh,车间,部门,姓名,erp编号,个人编号,出生日期,工作日期,岗位级别,身份证号,工资总额,月平均,缴费基数,养老保险 type XL5
yn = MESSAGEBOX("“D:\申报基数”电子表已导出！现打开使用吗？",4+32,"提示")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
*myexcel.sheetsinnewworkbook=1
*myexcel.visible=.t.
*myexcel.workbooks.add
*myexcel.worksheets("sheet1").activate
*next
myexcel.workbooks.open("D:\申报基数")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"  
ENDIF
else
=MESSAGEBOX("数据导入失败！请先处理"+sn+"年工资总额",48,"提示")
clear
 _SCREEN.WINDOWSTATE = 2
close all
use data\ccrq
go 8
repl 图片 with str(val(图片)+1,2)
if allt(图片)='10'
   repl 图片 with '0'
endi   
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
_SCREEN.MAXBUTTON = .F.
_SCREEN.MOVABLE = .F.
_SCREEN.CONTROLBOX = .F.
_SCREEN.CLOSABLE = .F. 
retu
endif
clear
 _SCREEN.WINDOWSTATE = 2
close all
use data\ccrq
go 8
repl 图片 with str(val(图片)+1,2)
if allt(图片)='10'
   repl 图片 with '0'
endi   
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
_SCREEN.MAXBUTTON = .F.
_SCREEN.MOVABLE = .F.
_SCREEN.CONTROLBOX = .F.
_SCREEN.CLOSABLE = .F.
retu