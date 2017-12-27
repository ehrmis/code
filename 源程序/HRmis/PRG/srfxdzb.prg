if file("d:\收入分析.xls")   
  yn = MESSAGEBOX("需要使用电子表编辑打印收入分析报表吗？",4+32,"提示")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           myexcel.workbooks.open("d:\收入分析.xls")
           myexcel.visible=.t.
           myexcel.caption="使用电子表编辑打印收入分析报表" 
           release oleapp
           WAIT CLEAR
ENDIF            
endif
if file("d:\工资累计排序.xls")   
  yn = MESSAGEBOX("需要使用电子表编辑打印工资累计排序报表吗？",4+32,"提示")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           myexcel.workbooks.open("d:\工资累计排序.xls")
           myexcel.visible=.t.
           myexcel.caption="使用电子表编辑打印工资累计排序报表" 
           release oleapp
           WAIT CLEAR
ENDIF            
endif
if file("d:\人员分类收入分析.xls")   
  yn = MESSAGEBOX("需要使用电子表编辑打印人员分类收入分析报表吗？",4+32,"提示")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           myexcel.workbooks.open("d:\人员分类收入分析.xls")
           myexcel.visible=.t.
           myexcel.caption="使用电子表编辑打印人员分类收入分析报表" 
           release oleapp
           WAIT CLEAR
ENDIF            
endif