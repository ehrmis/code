yn = MESSAGEBOX("�Ƿ񰴵�ǰ��Ա״���������ɼ�Ա�����˹������ݣ�",4+32,"��ʾ")
IF yn = 6
do srny.spx
sn=STR(VAL(nd)-1,4)
NY = ND+YF
NY1=ND+'01'
NY12=ND+'12'
CLOSE all
DO ����
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

yn = MESSAGEBOX("�Ƿ�emp.dbfģ�������ɼ�eHRmisϵͳ��Ա��Ϣxls��",4+32,"��ʾ")
IF yn = 6
USE emp1 EXCLUSIVE
ZAP
APPEND FROM ryzk
SET DATE TO ANSI
********yyyy/mm/dd************************
REPLACE cs WITH DTOC(csrq),cj WITH DTOC(cjgzrq),ht WITH dtoc(xhtrq) all
set date to long
********��/��/��**************************
USE emp EXCLUSIVE
ZAP
APPEND FROM emp1
FileName = GETFILE('XLS', '����emp: ', 'ȷ��')
IF  EMPTY(FileName)
	=MESSAGEBOX("�ļ���δָ����", Exclaim,"����")
ENDIF  
wjm=allt(FileName)
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
yn = MESSAGEBOX("���ݵ����ɹ����ִ�ʹ����",4+32,"��ʾ")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&wjm")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"   
ENDI   
ENDIF



******************
PROCEDURE ����
******************

DBFF="ryzk.dbf"
use &DBFF
***********Ϊ���������******************
REPLACE a WITH 9 all
copy to my
*********20151013***************    
close all
use my EXCLUSIVE
count to hav
**����*******
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


