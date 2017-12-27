close all
dbff=getfile('DBF', '打开数据库: ', '确定')
IF EMPTY(alltrim(dbff)) 
    dbff='&xp'+'\data\ryzk.dbf'
ENDIF
dbff=ALLTRIM(dbff)
**************标准默认打开数据源文件源代码*****20150921********  
use "&dbff" excl
go top
BROWSE TITLE '浏览字段名'+dbff   
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
whny=allt(WHNY)
*****************************************************文件“另存为”标准源程序代码***xls***txt*****20150907************************
FileName = GETFILE('dbf', '另存为: ', '确定')
IF  EMPTY(ALLTRIM(FileName))
	use
    thisform.release
    retu
ENDIF
FPath=justpath(FileName)
*************路径********** 
FName=juststem(FileName)
*************主名**********
wjm='&FPath'+'\'+'&FName'
**********路径＋主名**************
if len(WHTJ)>0
if len(WHNY)>0
   copy to &wjm fiel &whny for &WHTJ  
else
   copy to &wjm for &whtj
endif
else
if len(WHNY)>0
   copy to &wjm fiel &whny
else
   copy to &wjm 
endif   
endif
USE &wjm 
go top
brow titl '请认真维护所筛选的数据'
copy to &wjm type xls  
yn = MESSAGEBOX("需要使用电子表编辑报表吗？",4+32,"提示")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&wjm"+".xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"          
ENDIF 
