****************************************
* SJCSH.prg( Build 20170505 )
****************************************
close all
set safety off
set talk off
ND=str(year(date()),4)
YF=iif(month(date())>9,str(month(date()),2),'0'+str(month(date()),1))
ny=nd+yf
****ȫ�ֱ�����ֵ*****20161019*************
set date long
***********************************************************
USE ryzk1 EXCLUSIVE
PACK
USE ryzk EXCLUSIVE
PACK
repl a with 1 all
sort on rybh to temp
close all
use temp excl
go top
do while !eof()
   bh=rybh
   skip
   if bh=rybh
      repl a with 0
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a=0 to aaa
if aaa>0
   inde on rybh to temp
   go top
   brow noedit for a=0 titl '����'+allt(str(aaa,4))+'���غţ����ڡ���Ա״��ά�������޸���ȷ��ɾ�������¼��' 
endif
repl a with 1 all
close all
delete file temp?.dbf
delete file temp?.idx
******************
    use ryzk 
    COPY TO temp FIELDS dwdm,dwmc,CJDM,CJMC,xb,a
    COPY TO temp1 FIELDS dwdm,dwmc,a
    COPY TO cj FIELDS dwdm,dwmc,CJDM , CJMC , a
    COPY TO bm FIELDS  dwdm,dwmc,CJDM , CJMC ,BMBH , BMMC , a
***********����ά�������Զ����ɸ���ɸѡ�ּ����ܼ���Ȼ״��ͳ������******20170104**********      
   USE temp1
    INDEX ON dwdm TO temp1
    TOTAL ON dwdm TO DATA\dwdm
    USE data\dwdm
    ��λ����=ALLTRIM(dwmc)
    USE data\dmk
    REPLACE mc WITH ��λ����
    USE temp
    INDEX ON dwdm+CJDM+xb TO temp
    TOTAL ON dwdm+CJDM+xb TO temp1
   close all  
    SELECT 1
    USE EXCLUSIVE xb
    ZAP
    SELECT 2
    USE temp1
    GO TOP
    DO WHILE  .NOT. EOF()
       SELECT 1
       APPEND BLANK
       REPLACE dwdm WITH B.dwdm , dwmc with b.dwmc , cjdm WITH B.CJDM , cjmc WITH B.CJMC , code WITH B.xb , num WITH  ;
            B.a
       SELECT 2
       SKIP
    ENDDO
    SELECT 1
    REPLACE name WITH 'Ů' FOR code='0'
    REPLACE name WITH '��' FOR code='1'
close all    
    USE cj
    INDEX ON dwdm+CJDM TO cj
    TOTAL ON dwdm+CJDM TO DATA\cjk
    *********���Ӽ�ɸѡ���ּ��������******20170104****************
    USE bm
    INDEX ON dwdm+cjdm+BMBH TO bm
    TOTAL ON dwdm+cjdm+BMBH TO DATA\bmdm
    *********���Ӽ�ɸѡ���ּ��������******20170104****************
    SELECT 1
    USE EXCLUSIVE ����
    ZAP
    SELECT 2
    USE cjk
    GO TOP
    DO WHILE  .NOT. EOF()
       SELECT 1
       APPEND BLANK
       REPLACE ��λ���� WITH B.dwdm , ��λ���� with b.dwmc , ������� WITH B.CJDM , �������� WITH B.CJMC , ���� WITH  ;
            B.a
       SELECT 2
       SKIP
    ENDDO
close all
    SELECT 1
    USE EXCLUSIVE ����
    ZAP
    SELECT 2
    USE bmdm
    GO TOP
    DO WHILE  .NOT. EOF()
       SELECT 1
       APPEND BLANK
       REPLACE ��λ���� WITH B.dwdm , ��λ���� with b.dwmc ,������� WITH B.CJDM , �������� WITH B.CJMC , ���ű�� WITH B.BMBH , �������� WITH B.BMMC , ���� WITH B.a
       SELECT 2
       SKIP
    ENDDO
close all
**********************
use ryzk
count for year(csrq)=0 OR year(cjgzrq)=0 OR  empty(erprybh) to rs
if rs>0
   go top
   brow for year(csrq)=0 OR year(cjgzrq)=0 OR  empty(erprybh) fiel cjmc:r:h='����',bmmc:r:h='����',ryxm:r:h='����',erprybh:h='ERP���',csrq:h='��������',cjgzrq:h='��������' titl '����������'+allt(str(rs))+'�˵���Ҫ��Ϣ [ ��Esc���˳� ]'
ENDIF
count FOR empty(ygxz1) or empty(zgqk1) to rs
if rs>0
   loca FOR empty(ygxz1) or empty(zgqk1)
   brow FOR empty(ygxz1) or empty(zgqk1) fiel dwmc:r:h='��λ',cjmc:r:h='����',bmmc:r:h='����',ryxm:r:h='����',ygxz1:r:h='�ù�����',zgqk1:r:h='�ڸ����' titl '������Ա״��ά������ά������'+allt(str(rs))+'�ˡ��ù����ʻ��ڸ������ [ ��Esc���˳� ]'
endif
count FOR gwgz=0 AND ygxz1='������ǲ��' to rs
if rs>0
   loca FOR gwgz=0 AND ygxz1='������ǲ��'
   brow FOR gwgz=0 AND ygxz1='������ǲ��' fiel dwmc:r:h='��λ',cjmc:r:h='����',bmmc:r:h='����',ryxm:r:h='����',gz:r:h='���ִ���',gz1:r:h='����',gw:r:h='��λ',gwdc1:r:h='��λ����',gwgz:r:h='��λ����' titl '������Ա״��ά������ά������'+allt(str(rs))+'�ˡ����ָ�λ���λ���ʡ� [ ��Esc���˳� ]'
ENDIF
count FOR cjgz>0 AND zgqk1='�ڸ���Ա' to rs
if rs>0
   loca FOR cjgz>0 AND zgqk1='�ڸ���Ա'
   brow FOR cjgz>0 AND zgqk1='�ڸ���Ա' fiel dwmc:r:h='��λ',cjmc:r:h='����',bmmc:r:h='����',ryxm:r:h='����',cjgz:h='���⹤��',gz1:r:h='����',gw:r:h='��λ',zgqk1:r:h='�ڸ����' titl '���������'+allt(str(rs))+'�ˡ����⹤�ʡ� [ ��Esc���˳� ]'
endif
*sum jfjs to js
*if js>0
*   count for jfjs=0 to rs
*   if rs>0
*      go top
*      brow for jfjs=0 fiel cjmc:r:h='����',bmmc:r:h='����',ryxm:r:h='����',csrq:r:h='��������',cjgzrq:r:h='��������'whcd1:r:h='ѧ��',jfjs:h='�ɷѻ���' titl '����������'+allt(str(rs))+'�ˡ��ɷѻ����� [ ��Esc���˳� ]'
*   endif
*endif
*******��ְ�����乤�ʴ���*****************
replace YLF with 2002-year(CJGZRQ)+1+20 for CJGZRQ<ctod('2002-7-1')
replace gwjt with (2003-year(CJGZRQ)+1)*15 for CJGZRQ<ctod('2003-7-1') and allt(zgqk1)='�ڸ���Ա'
**************�������䣨���꣩���䣨���꣩����******������λС��****************
replace nL with ROUND(year(date())-year(Csrq)+(month(date())-month(Csrq))/12,2) for month(Csrq)-month(date())<=0 and year(csrq)>0
replace nL with ROUND(year(date())-year(Csrq)-(month(Csrq)-month(date()))/12,2) for month(Csrq)-month(date())>0 and year(csrq)>0
*************************************�������䣨���꣩����***********************************************************************************
*replace GL with ND1-year(Cjgzrq) for month(Cjgzrq)-month(date())<=0 and year(cjgzrq)>0
*replace GL with ND1-year(Cjgzrq) for month(Cjgzrq)-month(date())>0 and year(cjgzrq)>0 
***********************************��������Թ��乤�ʼ���ʱʹ�����䣬ռ��һ����һ��ļ��㷽��******20170125************************************************* 
replace gL with ROUND(year(date())-year(Cjgzrq)+(month(date())-month(Cjgzrq))/12,2) for month(Cjgzrq)-month(date())<=0 and year(cjgzrq)>0
replace gL with ROUND(year(date())-year(Cjgzrq)-(month(Cjgzrq)-month(date()))/12,2) for month(Cjgzrq)-month(date())>0 and year(cjgzrq)>0 
*************************************���ݼ�ʹ��ʵ�ʹ����������***********************************************************************************
replace �������� with ALLTRIM(LEFT(STR(nl,5,2),2))+'��'+ALLTRIM(STR(VAL(right(STR(nl,5,2),2))/100*12))+'����' all
replace �������� with ALLTRIM(STR(VAL(right(STR(nl,5,2),2))/100*12))+'����' for ALLTRIM(LEFT(STR(nl,5,2),2))='0'
replace �������� with ALLTRIM(LEFT(STR(nl,5,2),2))+'��' for ALLTRIM(STR(VAL(right(STR(nl,5,2),2))/100*12))='0'
*************************************�������ޣ����꣩����***********************************************************************************
replace �������� with ALLTRIM(LEFT(STR(gl,5,2),2))+'��'+ALLTRIM(STR(VAL(right(STR(gl,5,2),2))/100*12))+'����' all
replace �������� with ALLTRIM(STR(VAL(right(STR(gl,5,2),2))/100*12))+'����' for ALLTRIM(LEFT(STR(gl,5,2),2))='0'
replace �������� with ALLTRIM(LEFT(STR(gl,5,2),2))+'��' for ALLTRIM(STR(VAL(right(STR(gl,5,2),2))/100*12))='0'
replace NLD with '20������' for NL<21
replace NLD with '21��25��' for NL>=21 and NL<26
replace NLD with '26��30��' for NL>=26 and NL<31
replace NLD with '31��35��' for NL>=31 and NL<36
replace NLD with '36��40��' for NL>=36 and NL<41
replace NLD with '41��45��' for NL>=41 and NL<46
replace NLD with '46��50��' for NL>=46 and NL<51
replace NLD with '51��55��' for NL>=51 and NL<56
replace NLD with '56��60��' for NL>=56 
replace GLD with '25������' for GL<25
replace GLD with '25��29��' for GL>=25 and GL<30
replace GLD with '30��34��' for GL>=30 and GL<35
replace GLD with '35������' for GL>=35
COPY TO temp FIELDS dwdm,dwmc,CJDM,CJMC,nld,a
COPY TO temp1 FIELDS dwdm,dwmc,CJDM,CJMC,gld,a
USE temp
INDEX ON dwdm+CJDM+nld TO temp
TOTAL ON dwdm+CJDM+nld TO nld
USE temp1
INDEX ON dwdm+CJDM+gld TO temp1
TOTAL ON dwdm+CJDM+gld TO gld 
USE ryzk
sort on bmbh,zw,rybh to temp 
use temp
GO top
bh=rybh
REPLACE a WITH 1,���� WITH 1 all
COPY TO ry&ny
*************����͸�����Ա��Ϣ�����±���ryzk.dbf��****bug*******
COPY TO ��Ա״�� 
COPY TO data\ryzk
COPY TO data\tmpdbf
close all
USE gzze1 EXCLUSIVE
ZAP
USE lwtemp EXCLUSIVE
ZAP
use ccrq
go 8
��ֽ��=fdmax
I=val(ͼƬ)
pic=ALLTRIM(ͼƬ)
repl ͼƬ with ALLTRIM(str(I+1))
IF I>=��ֽ��
   repl ͼƬ with '0'
ENDIF
**********�û�������־�����׽���*********
pict_fd='fd'+pic+'.jpg'
_SCREEN.PICTURE = '&pict_fd'
close all
