 =messagebox("���õ�������������Ϲ��������������ͱ����ֻ����Ź�����",48,"��ʾ")  
  yn = MESSAGEBOX("��Ҫʹ�õ��ӱ�༭��������",4+32,"��ʾ")
IF yn = 6
FileName = GETFILE('XLS', '������: ', 'ȷ��')
IF EMPTY(alltrim(FileName)) 
    FileName='&xp'+'\������ģ��.xls'
ENDIF
wjm=ALLTRIM(FileName)
**************��׼Excel��ȡ�ļ�*************
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&wjm")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\������.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭����" 
release oleapp
WAIT CLEAR 
run web\���ʹ���.exe
******ֱ�ӵ��õ�����Ӧ��***20171220***               
ENDIF   