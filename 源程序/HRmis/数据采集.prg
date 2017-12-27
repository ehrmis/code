yn = MESSAGEBOX("是否按当前人员状况库批量采集员工个人工资数据？",4+32,"提示")
IF yn = 6
do srny.spx
sn=STR(VAL(nd)-1,4)
NY = ND+YF
NY1=ND+'01'
NY12=ND+'12'
CLOSE all
DO 数组
FOR ny1=VAL(NY1) TO VAL(NY12)
    ny=STR(ny1,6)
SELECT 2
USE ryzk
INDEX on rybh TO ryzk
SELECT 1
IF files('gz'+ny+'.dbf')
USE gz&ny EXCLUSIVE
INDEX on rybh TO gz&ny
UPDATE on rybh from B repl a with b.a
DELETE FOR a<9
PACK
BROWSE
ENDIF
ENDFOR
CLOSE all 
ENDIF

yn = MESSAGEBOX("是否按emp.dbf模板批量采集eHRmis系统人员信息xls？",4+32,"提示")
IF yn = 6
USE emp1 EXCLUSIVE
ZAP
APPEND FROM ryzk
SET DATE TO ANSI
********yyyy/mm/dd************************
REPLACE cs WITH DTOC(csrq),cj WITH DTOC(cjgzrq),ht WITH dtoc(xhtrq) all
set date to long
********年/月/日**************************
USE emp EXCLUSIVE
ZAP
APPEND FROM emp1
FileName = GETFILE('XLS', '导出emp: ', '确定')
IF  EMPTY(FileName)
	=MESSAGEBOX("文件名未指定！", Exclaim,"错误")
ENDIF  
wjm=allt(FileName)
WHNY=''
do while .T.
getexpr '各字段之间用“,”连接；不输则表示全部字段导出：' to WHNY
WHNY=allt(WHNY)
IF  "'"$WHNY or '"'$WHNY  
     MESSAGEBOX('输入内容中不能含'+"'"+'"'+'字符！','提示')
else
     EXIT
ENDIF
enddo
WHTJ=''
getexpr '表达式要符合逻辑；不输条件则表示全部记录导出：' to WHTJ
WHTJ=allt(WHTJ)
if len(WHTJ)>0
if len(WHNY)>0
   copy to &wjm field &WHNY for &WHTJ type xl5
else
   copy to &wjm for &WHTJ type xl5   
endif
else
if len(WHNY)>0
   copy to &wjm field &WHNY type xl5
else
   copy to &wjm type xl5
endif
endif
yn = MESSAGEBOX("数据导出成功！现打开使用吗？",4+32,"提示")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&wjm")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"   
ENDI   
ENDIF



******************
PROCEDURE 数组
******************

DBFF="ryzk.dbf"
use &DBFF
***********为主程序服务******************
REPLACE a WITH 9 all
copy to my
*********20151013***************    
close all
use my EXCLUSIVE
count to hav
**数组*******
dimension sz(hav)
for i=1 to fcount()
ss=field(i)
if upper(alltrim(type(ss)))=upper(alltrim("d"))
for aa=1 to hav
go aa
sz(aa)=dtoc(&ss,1)
endfor
ALTER   TABLE  my  ALTER   COLUMN  &ss  c(24)
for aa=1 to hav
go aa
replace &ss with sz(aa)
endfor
else
ALTER   TABLE  my  ALTER   COLUMN  &ss  c(24)
endif
ENDFOR
brow


