***************************
*ERPkq.PRG(Build2005.08.01)
***************************
close all
on key label F1 do grcx
YF1=month(date())
ND1=year(date())
do srny.spx
clear
NY = ND+YF
close all
***********************��.��ʼ��*****************
IF !file("GZ"+NY+".txt")
    =MESSAGEBOX("ERP�����ļ�GZ"+ny+".TXTδ����...",0+48,"����")
    retu
ELSE   
use erpgz excl
zap
APPEND FROM gz&ny..txt DELIMITED WITH CHAR TAB
REPLACE erprybh WITH ALLTRIM(STR(d)) all
REPLACE erprybh WITH '00'+ALLTRIM(erprybh) FOR LEN(ALLTRIM(erprybh))=6
REPLACE erprybh WITH '0'+ALLTRIM(erprybh) FOR LEN(ALLTRIM(erprybh))=7
ENDIF
close all
sele 2
use ryzk
inde on erprybh to erpgz
sele 1
if !file('kq'+ny+'.dbf')
   use kqk excl
   zap
   appe from erpgz
   set relation to erprybh into 2
repl rybh with b.rybh,cjmc with b.cjmc,cjdm with b.cjdm,bmbh WITH b.bmbh,bmmc WITH b.bmmc all 
   sort on bmbh,rybh to kq&ny
use kq&ny excl  
repl �а� with zbgr,ҹ�� with ybgr,���ռӰ� with jrjb,���� with bj,�¼� with sj,���� with cj,���� with gs,̽�׼� with tqj,;
��ɥ�� with hsj,���� with jl,���� with kg,���ݼ� with gj all
go top
brow noedit part 40 titl '���������ά����ERP�������ݵ��������'
nn=MESSAGEBOX("���ݵ���ɹ�...",0+64,"��Ϣ")
else
nn=MESSAGEBOX(nd+"��"+allt(str(val(yf)))+"�·ݿ����Ѵ��ڣ����ݵ���ʧ�ܣ�",0+64,"��Ϣ")
*****************��ȫ��******************
endif
clear
close all
delete file erp*.*
