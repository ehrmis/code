if file("d:\�������.xls")   
  yn = MESSAGEBOX("��Ҫʹ�õ��ӱ�༭��ӡ�������������",4+32,"��ʾ")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           myexcel.workbooks.open("d:\�������.xls")
           myexcel.visible=.t.
           myexcel.caption="ʹ�õ��ӱ�༭��ӡ�����������" 
           release oleapp
           WAIT CLEAR
ENDIF            
endif
if file("d:\�����ۼ�����.xls")   
  yn = MESSAGEBOX("��Ҫʹ�õ��ӱ�༭��ӡ�����ۼ����򱨱���",4+32,"��ʾ")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           myexcel.workbooks.open("d:\�����ۼ�����.xls")
           myexcel.visible=.t.
           myexcel.caption="ʹ�õ��ӱ�༭��ӡ�����ۼ����򱨱�" 
           release oleapp
           WAIT CLEAR
ENDIF            
endif
if file("d:\��Ա�����������.xls")   
  yn = MESSAGEBOX("��Ҫʹ�õ��ӱ�༭��ӡ��Ա�����������������",4+32,"��ʾ")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           myexcel.workbooks.open("d:\��Ա�����������.xls")
           myexcel.visible=.t.
           myexcel.caption="ʹ�õ��ӱ�༭��ӡ��Ա���������������" 
           release oleapp
           WAIT CLEAR
ENDIF            
endif