close all
dbff=getfile('DBF', '�����ݿ�: ', 'ȷ��')
IF EMPTY(alltrim(dbff)) 
    dbff='&xp'+'\data\ryzk.dbf'
ENDIF
dbff=ALLTRIM(dbff)
**************��׼Ĭ�ϴ�����Դ�ļ�Դ����*****20150921********  
use "&dbff" excl
go top
BROWSE TITLE '����ֶ���'+dbff   
WHNY=''
do while .T.
getexpr '���ֶ�֮���á�,�����ӣ��������ʾȫ���ֶε�����' to WHNY
WHNY=allt(WHNY)
IF  "'"$WHNY or '"'$WHNY  
     MESSAGEBOX('���������в��ܺ�'+"'"+'"'+'�ַ���','��ʾ')
else
     EXIT
ENDIF
enddo
WHTJ=''
getexpr '���ʽҪ�����߼��������������ʾȫ����¼������' to WHTJ
WHTJ=allt(WHTJ)
whny=allt(WHNY)
*****************************************************�ļ������Ϊ����׼Դ�������***xls***txt*****20150907************************
FileName = GETFILE('dbf', '���Ϊ: ', 'ȷ��')
IF  EMPTY(ALLTRIM(FileName))
	use
    thisform.release
    retu
ENDIF
FPath=justpath(FileName)
*************·��********** 
FName=juststem(FileName)
*************����**********
wjm='&FPath'+'\'+'&FName'
**********·��������**************
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
brow titl '������ά����ɸѡ������'
copy to &wjm type xls  
yn = MESSAGEBOX("��Ҫʹ�õ��ӱ�༭������",4+32,"��ʾ")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&wjm"+".xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"          
ENDIF 
