USE  ��ѵ�Ǽ�
jls=RECCOUNT()
if jls=0
=MESSAGEBOX('�����ڡ���Ա״��������ά���н�����ѵ�Ǽǡ���', 16,"��ʾ")
 RETURN
ELSE
********��ݱ����е��á���ѵ�Ǽ����ݴ������̣����ֲ��������ӱ�����***20170906************
CLOSE all
SELECT 2 
USE ryzk
REPLACE a WITH 1 all  
INDEX on rybh TO ryzk
SELECT 1
use ��ѵ�Ǽ� excl
repl a with 0 all
SET RELATION TO rybh inTO 2
REPLACE dwdm with b.dwdm,dwmc with b.dwmc,cjdm with b.cjdm,cjmc with b.cjmc,bmbh with b.bmbh,bmmc with b.bmmc,xb WITH b.xb,xb1 WITH b.xb1,zw WITH b.zw,zw1 WITH b.zw1,csrq WITH b.csrq,;
nl WITH b.nl,cjgzrq WITH b.cjgzrq,gl WITH b.gl,gz1 WITH b.gz1,gw WITH b.gw,gwfl1 with b.gwfl1,a with b.a FOR rybh=b.rybh
*******����ʹ��for����****��Ϊͬһ������ж�����¼��������UPDATE���*****UPDATE on rybh from B repl���������A���ж���ͬ�ż�¼ʱֻ���µ�һ����¼****20171130***************
GO top
loca for a=0 
IF !eof()
   delete for a=0 
   brow for a=0 titl '���������������Ա�Ƿ�ɾ��'
    yn = MESSAGEBOX("��Ҫɾ�����������Ա��",4+32,"��ʾ")
   IF yn = 6
      PACK
****���ܣ�����ʾɾ���ٹ���������Ա״����ȷ���Ƿ��м�����Ա��******20171130*************
   ELSE
       RECALL all
   ENDIF         
ENDIF
repl a with 1 all
CLOSE ALL
**********�౾֤����ֻ���±�����߼���ѵ(��һ�м�¼)***********  
use ��ѵ�Ǽ� excl
inde on ���� to ��ѵ�Ǽ�
DELETE FOR ryxm<>���� 
*********ryxm�������Զ����������˵ļ�¼ʱ���ɵļ�¼����������������һ���ռ�¼(ֻ����ˡ���ѵ�Ǽǡ���û�������κ�����)*****20151115�޸�����ѵ�Ǽǡ�Bug***����ϵͳ�ȶ���**************
COUNT FOR ryxm<>���� TO rs
IF rs>0
  GO top
  BROWSE FIELDS ���,��λ,����,����,����,��ѵ����,������ҵ,��Ŀ,ͽ��,ʦ��ͽ����,��ʼ����,��ֹ����,��ѵ�ȼ�,֤������,֤����,��֤��λ,��֤����,fszq:h='��������',���ڸ��� FOR ryxm<>���� titl '�������ѵ���[��F1���Ҹ���  ��ʹ�á����˵��������ӻ�ɾ�������¼]'
ENDIF
inde on ��� to ��ѵ�Ǽ� 
GO top 
BROWSE FIELDS ���,��λ,����,����,����,��ѵ����,������ҵ,��Ŀ,ͽ��,ʦ��ͽ����,��ʼ����,��ֹ����,��ѵ�ȼ�,֤������,֤����,��֤��λ,��֤����,fszq:h='��������',���ڸ��� titl '�������ѵ���[��F1���Ҹ���  ��ʹ�á����˵��������ӻ�ɾ�������¼]'
repl djdm with '08' for '�߼�'$��ѵ�ȼ� and gwfl1='������λ'
repl djdm with '09' for '�м�'$��ѵ�ȼ� and gwfl1='������λ'
repl djdm with '10' for '����'$��ѵ�ȼ� and gwfl1='������λ'
*****************������λ��Աְҵ�ʸ�֤��������ٶԲ�ͬ��λ��Ա���������Ա����******20170906*****************************************
repl djdm with '99' for gwfl1<>'������λ' OR VAL(djdm)=0
repl djdm with '06' for ALLTRIM(��ѵ�ȼ�)='�߼���ʦ'
repl djdm with '07' for ALLTRIM(��ѵ�ȼ�)='��ʦ'
**********************���洦������ȼ�����***********20170906*********************************
REPLACE ��ѵ�ȼ� WITH '������ҵ',djdm with '11',֤������ with '��ȫ����֤' FOR !EMPTY(������ҵ)
REPLACE ��ѵ�ȼ� WITH 'ʦ��ͽ',djdm with '12',��ѵ���� with 'ʦ��ͽ��ѵ' FOR !EMPTY(ʦ��ͽ����)
repl a with 1,֤���� with allt(֤����) all
**************��ѵ����*************20170906**********************************
copy to temp for EMPTY(֤����) AND VAL(djdm)<>12
************�ȹ����ޱ����ѵ֤��***********************
sort on ֤���� to temp1 for !EMPTY(֤����)
copy to temp2 for EMPTY(֤����) AND VAL(djdm)=12
************һ��Ϊ���ȹ����ޱ����ѵ֤�飬���������¼(֤���غ�)*****��׼�ǼǷ�����֤���б��****20170906**************
use temp1 excl
go top
do while !eof()
   bh=allt(��ѵ����)+allt(֤����)+allt(��ѵ�ȼ�)
   skip
   if bh=allt(��ѵ����)+allt(֤����)+allt(��ѵ�ȼ�)
      repl a with 2
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a<>1 to aaa
IF aaa>0
   inde on ֤���� to temp1
   go top
   DELETE FOR a=2
   brow for a<>1 titl '����'+allt(str(aaa,4))+'֤���غţ����޸���ȷ��ɾ�������¼' 
   yn = MESSAGEBOX("�Ƿ�ɾ���غ���Ա��",4+32,"��ʾ")
IF yn = 6
PACK
ELSE
RECALL
ENDIF
ENDIF
CLOSE all
use temp
********��֤��*************
appe from temp1
********��֤��*******************
appe from temp2
********************��ʦ��ͽ����¼***********20170906**********************
sort on rybh,djdm to temp3
use temp3 excl
delete for empty(��ѵ����)
***********û�Ǽ���ѵ����*******20170906****************
pack 
use ��ѵ�Ǽ� excl 
zap
appe from temp3
********�ȼ����������Ա״����******20170906***************
REPLACE ��λ with ALLTRIM(dwmc),���� with ALLTRIM(cjmc),���� with ALLTRIM(bmmc),pxmc with ALLTRIM(��ѵ����),pxdj with ALLTRIM(��ѵ�ȼ�),a with 1 all
*********Ҫ��һ�˶�֤����߼�����RYZK.DBF*********************
inde on ���� to ��ѵ�Ǽ�
GO top
BROWSE FIELDS ����,��λ,����,����,��ѵ����,������ҵ,��Ŀ,ͽ��,ʦ��ͽ����,��ʼ����,��ֹ����,��ѵ�ȼ�,֤������,֤����,��֤��λ,��֤����,��������,���ڸ��� TITLE '���������Ա����ѵ�Ǽ����'
CLOSE ALL
SELECT 2 
USE ryzk  
go top
do while !eof()
   SELECT 1
   use ��ѵ�Ǽ�    
   go top
   loca for rybh=b.rybh and gwfl1='������λ' and VAL(djdm)<11
   ***************��ѵ��Ϣ��������Ա״������**************��������ҵ��11�����á�ʦ��ͽ��12������ѵ����********20170906**********************************
   if !eof()
      repl b.pxdj with ALLTRIM(pxdj),b.pxmc with ALLTRIM(pxmc),b.֤���� with ALLTRIM(֤����),b.��֤��λ with ALLTRIM(��֤��λ),b.��֤���� with ��֤����,b.zcrq with ��֤����
**********�౾֤�������λ��Աֻ���±�����߼�ְҵ�ʸ���ѵ(��һ�м�¼)***********20151115*******������λ��ְҵ�ʸ�ȼ�ֻ֤�з�֤���ڣ�Ҳ���ǡ�ְ��/�ʸ��϶����ڡ�*****20160816***   
   endif
  sele 2
  wait window NOWAIT '���ڱ�����ѵ�Ǽ�......' 
  skip 
enddo 
CLOSE all
USE ryzk 
repl zcdj1 with allt(pxdj) for gwfl1='������λ'
REPLACE zc with LEFT(zcdj1,4)+allt(pxmc) for gwfl1='������λ' AND !'��ʦ'$pxdj
REPLACE zc with allt(pxmc)+LEFT(zcdj1,8) for gwfl1='������λ' AND '��ʦ'$pxdj
REPLACE zc WITH ALLTRIM(zc) all
SELECT 2
USE jsdj
scan
  select 1
  USE ryzk
  wait window NOWAIT '���ڱ���ְҵ�ʸ�ȼ�......' 
  repl zcdj with b.code for ALLTRIM(zcdj1)=ALLTRIM(b.name)
ENDSCAN
CLOSE ALL
USE ��ѵ�Ǽ�
SORT ON bmbh,rybh,djdm TO temp
********���ȼ���������****20170906********************
USE deset
LOCATE FOR oop='backup'
I=seted
USE temp
REPLACE ��� with str(recn(),5),��λ with ALLTRIM(dwmc),���� with ALLTRIM(cjmc),���� with ALLTRIM(bmmc),a WITH 1,���� WITH 1 all
COPY TO data\��ѵ�Ǽ�
********ϸ�ڣ���ԭĸ��*************
DBFF1='��ѵ�Ǽ�'
WHNY=ND+'��ѵ�Ǽ�'
do Qdir
********ϸ�ڣ���ʾ����*************
CLOSE ALL
ENDIF