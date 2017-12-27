 =messagebox("利用第三方软件“掌上工资条”快速推送本月手机短信工资条",48,"提示")  
  yn = MESSAGEBOX("需要使用电子表编辑工资条吗？",4+32,"提示")
IF yn = 6
FileName = GETFILE('XLS', '工资条: ', '确定')
IF EMPTY(alltrim(FileName)) 
    FileName='&xp'+'\工资条模板.xls'
ENDIF
wjm=ALLTRIM(FileName)
**************标准Excel读取文件*************
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&wjm")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\工资条.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑报表" 
release oleapp
WAIT CLEAR 
run web\工资管理.exe
******直接调用第三方应用***20171220***               
ENDIF   