*****************************
* .\SRFX.PRG(Build2017.1.10)
*****************************
close all
USE sbdmk
ͳ���·�=�����·�
**********��ϵͳ�����Զ����·�*******20170110************
USE deset
LOCATE FOR ALLTRIM(oop)='px'
fbl=seted
USE dmk
��ְ��=VAL(��ְ�����)
Сְ��=VAL(Сְ�����)
set safety off
on key label F1 do grcx
**************˫λ�����Զ�����*******************
IF month(DATE())<=ͳ���·�
   jn=STR(YEAR(DATE()),4)
   nd=STR(YEAR(DATE())-1,4)
   YF1 = 12
   MESSAGEBOX("ϵͳ���ýɷѻ��������·�Ϊ"+ALLTRIM(STR(ͳ���·�))+"�£��ڴ���֮ǰͳ�Ʒ���"+nd+"�깤�����ݣ���Ҫͳ�Ʒ���"+jn+"�깤������,����ϵͳ���ýɷѻ��������·��ڱ���֮ǰ��",64,"��ʾ")
*******������ȫ������ͳ�Ʒ���ʱ����·�������ڵ����ɷѸ����·ݡ���֮��Ĭ����Ȼ��*******20170110***********
ELSE
   nd=STR(YEAR(DATE()),4)
   YF1 = month(DATE())
ENDIF 
LY1=str(val(ND)-1,4)
*�������=ƽ������
*************************************************
if  not file('zr'+LY1+'.dbf')
  =MESSAGEBOX("���ȴ���"+LY1+"�깤���ܶ�!",48,"��ʾ")
    close all
    return
endif    
if  not file('zr'+ND+'.dbf')
  =MESSAGEBOX("���ȴ���"+ND+"�깤���ܶ�!",48,"��ʾ")
    close all
    return
endif
close all
delete file zr11&LY1.dbf
delete file zr11temp.dbf
delete file temp?.dbf
sele 2
use ryzk
inde on rybh to ryzk
sele 1
use zr&ND
inde on rybh to zr&ND
upda on rybh from B repl zzdm with b.zzdm,zzjg with b.zzjg,cjdm with b.cjdm , cjmc with b.cjmc , bmbh with b.bmbh , bmmc with b.bmmc, zw with b.zw,;
zw1 with b.zw1 , gwfl with b.gwfl , gwfl1 with b.gwfl1,zcdj with b.zcdj,zcdj1 with b.zcdj1,ygxz with b.ygxz,ygxz1 with b.ygxz1,zgqk with b.zgqk,zgqk1 with b.zgqk1
copy to zr11&ND for pj>0 AND  'ְ��'$ygxz1 and  '�ڸ�'$zgqk1 
*****��������ְ�����������ʱ�⣬���ϸ��޹�������ְ������*********
copy to zr11&LY1 stru
use zr11&LY1 excl
appe from ryzk for 'ְ��'$ygxz1 and  '�ڸ�'$zgqk1 
**** &&ʹ�����ڵ�ryzk.dbf�ڸ�ְ����Ϣ������������������ݿ�********��Ҫ�����֯�����䶯������������ԣ���Ŀǰ��ryzk.dbf��Ա״����Ϣ������������겻���ڵ�������֯������������ʹ��������ʵ�������ݣ���������������Ա************20160121*************************
repl ylbx with 0,qynj with 0,sybx with 0,ybx with 0,gjj with 0,sds with 0 all
***********�﷨�Ͻ��ԣ�ѭ���ݼӳ�ʼ��*****����ryzk.dbf��ԭ������************
yesno = MESSAGEBOX('��Ҫ��'+LY1+'��'+ALLTRIM(STR(YF1))+'��ͬ�ڱȽ��𣿷�����'+LY1+'����ƽ������Ƚ�',291,'��ʾ')
*****ѡ��������ͬ�ڱȽϻ���ƽ���Ƚ�
 DO CASE 
CASE yesno=6
   M1=1
  do while M1<=yf1
  MGZ = iif(M1>9,str(M1,2),'0'+str(M1,1))
  wait wind  '������д'+LY1+'��'+allt(str(val(mgz)))+'��ְ�������ܶ�...'  nowa
  if  not file('GZ'+LY1+MGZ+'.dbf')
 ***************ERP****************************
    MESSAGEBOX("�뵼��ERPϵͳ��"+LY1+"��"+allt(str(val(MGZ)))+"�¹�������",48,"��ʾ")
  endif  
***********���˲�����Ա�����޹��ʵ���ylbx**********
    close all
    select 2
    use gzk excl
    zap
    appe from gz&LY1.&Mgz
    ***********************��1��ͬ�ڱȽ��������¹��ʿ������ۼ�*********20160121**************************
    index on RYBH to gzk
    select 1
    use zr11&LY1
    index on RYBH to zr11&LY1
    upda on rybh from B replace ylbx with ylbx+b.ylbx,qynj with qynj+b.qynj,ybx with ybx+b.ybx,sybx with sybx+b.sybx,gjj with gjj+b.gjj,sds with sds+b.sds,sfje with sfje+b.sfje,m&Mgz with b.hj-b.jang ,j&Mgz with b.jang 
    M1 = M1+1   
 enddo
close all
wait wind  'ͬ�ڱȽ����������'+LY1+'��'+ALLTRIM(STR(yf1))+'�¹��ʿ������ۼƣ�ʵ��������...'  nowa
use zr11&LY1 excl
go top
do while !eof()
   M111=1
   for MJB1 = 1 to 12
    J111 = iif(MJB1>9,str(MJB1,2),'0'+str(MJB1,1))
    if (M&J111>0 or J&J111>0) and MOU<>12
       replace MOU with M111
       M111=M111+1
    endif      
  endfor
  skip
enddo
delete for mou=0
pack
repl gzhj with m01+m02+m03+m04+m05+m06+m07+m08+m09+m10+m11+m12,jjhj with j01+j02+j03+j04+j05+j06+j07+j08+j09+j10+j11+j12,hj with gzhj+jjhj,pj with round(hj/mou,0) all
************************************************************���¼�������ͬ������ˮƽ�������µ���ƽ����*************20160121***********************************************************************************************
CASE yesno=7
****�Ƿ�ͬ�ȣ���7��Ĭ�ϰ�������ƽ���Ƚ�*******20160121********************
close all
sele 2
use zr&LY1
inde on rybh to zr&LY1
***********��2����������ԭʼ���������ͬ������Ƚ�******20160121********************
sele 1
use zr11&LY1 excl
inde on rybh to zr11&LY1
upda on rybh from B repl mou with b.mou,sfje with b.sfje,jjhj with b.jjhj,gzhj with b.gzhj,hj with b.HJ,pj with b.pj;
ylbx with b.ylbx,qynj with b.qynj,ybx with b.ybx,sybx with b.sybx,gjj with b.gjj,sds with b.sds
CASE yesno=2
     RETURN
ENDCASE
delete for mou=0
pack
**************************��ʼ�����***��������ͬ�ڹ����ܶ����������ʱ���ݿ⣨�����������**********20160121*********************
close all
use zr11&LY1 excl 
index on rybh to zr11&LY1
repl cjdm with '77',cjmc with '���Ÿ�����' for val(zw)>��ְ�� and val(zw)<=Сְ��
repl cjdm with '88',cjmc with '��λ�쵼' for val(zw)<=��ְ��
index on CJDM+rybh to zr11&LY1
DELETE FOR pj=0
****&&(�ؼ�ϸ�ڣ�����λ�����ܶ������λ��������ܶ������һ�£���������������Ա������õ�����λ������ʵ���˾�����ˮƽ�����ܼ���������ʵ��������%******20160121*****************************************)
GO top
BROWSE TITLE '����������������'+LY1+'��'+ALLTRIM(STR(yf1))+'��������������������Դ˱�������'+ND+'��ͬ������Ƚϣ������������λʵ����������' 
pack
index on CJDM to zr11&LY1
total on CJDM to zr11
USE zr11
sum to LJ1,RS1 HJ,A
appe blank
go bott
repl cjmc with '�ϡ�����',hj with LJ1,a WITH rs1
if yesno=6 
   replace PJ with round(HJ/A,0) all      
  else 
   replace PJ with round(HJ/A/yf1,0) all   
endif
GO top
browse field CJMC:H = ' ��λ���� ' , A :P = '@z 9,999' : 15 :H = '����', HJ :P = '@z 999,999,999.99' : 14 :H = '������ۼ�(Ԫ)' , PJ :P = '@z 999,999' : 15 :H = 'ƽ��' TITLE '���������'+LY1+'��ͬ���ۼ����'
******** &&(����ͬ�ڻ�������*�������*���и���λ�������������������������Ա������ή������ͬ�ڸ���λ����ʵ����ˮƽ**20160121**************)
use zr11&ND excl 
inde on rybh to zr11&ND
repl cjdm with '77',cjmc with '���Ÿ�����' FOR val(zw)>��ְ�� and val(zw)<=Сְ��
repl cjdm with '88',cjmc with '��λ�쵼' for val(zw)<= ��ְ��
index on CJDM+rybh to zr11&ND
DELETE FOR pj=0
****&&(�ؼ�ϸ�ڣ�ֱ��ɾ����������޹����ܶ����Ա��ȷ������λ�����ܶ������λ��������ܶ������һ�£���������������Ա������õ�����λ������ʵ���˾�����ˮƽ�����ܼ���������ʵ��������%******20160121*****************************************)
GO top
BROWSE TITLE '����������������'+ND+'��'+ALLTRIM(STR(yf1))+'��������������������Դ˱�������'+LY1+'��ͬ������Ƚϣ������������λʵ����������' 
pack
index on CJDM to zr11&ND
total on CJDM to zr1  &&(���������ͳ�ƻ���)
USE zr1 EXCLUSIVE
sum to LJ,RS HJ,A
appe blank
go bott
repl cjmc with '�ϡ�����',hj with LJ,a WITH rs
if yesno=6 
   replace PJ with round(HJ/A,0) all      
  else 
   replace PJ with round(HJ/A/yf1,0) all   
endif
GO top
browse field CJMC:H = ' ��λ���� ' , A :P = '@z 9,999' : 15 :H = '����', HJ :P = '@z 999,999,999.99' : 14 :H = '������ۼ�(Ԫ)' , PJ :P = '@z 999,999' : 15 :H = 'ƽ��' TITLE '���������'+ND+'���ۼ����'
DELETE FOR  '�ϡ�����'$cjmc
pack
close all
select 2
use zr11
repl m12 with ylbx+qynj+sybx+ybx+gjj all
sum HJ,m12,sds,sfje,A to LJ1,sb1,gs1,sf1,RS1
index on CJDM to zr11
if yesno=6 
   replace PJ with round(HJ/A,0),jang with round(sfje/A,0) all      
 else 
   replace PJ with round(HJ/A/yf1,0),jang with round(sfje/A/yf1,0) all   
endif
select 1
use zr1 
repl m12 with ylbx+qynj+sybx+ybx+gjj all
sum HJ,m12,sds,sfje,A to LJ,sb,gs,sf,RS
index on CJDM to zr1
if yesno=6 
   replace PJ with round(HJ/A,0),jang with round(sfje/A,0) all      
 else 
   replace PJ with round(HJ/A/yf1,0),jang with round(sfje/A/yf1,0) all   
endif
upda on cjdm from B repl m01 with round(PJ-b.PJ,0),m03 with round(jang-b.jang,0),m02 with round(m01/b.PJ*100,2),m04 with round(m03/b.jang*100,2) 
appe blank
repl cjdm with '99' for empty(cjdm)
go bott
if yesno=6 
   repl cjmc with '�ϡ�����',hj with LJ,pj with round(lj/rs,2),m01 with round(LJ/rs,1)-round(LJ1/rs1,1),;
m02 with m01/round(LJ1/rs1,1)*100,jang with round(sf/rs,2),m04 with (round(sf/rs,1)-round(sf1/rs1,1))/round(sf1/rs1*100,2),m12 with sb,sds with gs,sfje with sf,a with rs      
 else 
   repl cjmc with 'ȫ�깤���ܶ�',hj with LJ,pj with round(lj/rs/yf1,2),m01 with round(LJ/rs/yf1,1)-round(LJ1/rs1/12,1),;
m02 with m01/round(LJ1/rs1/12,1)*100,jang with round(sf/rs/yf1,2),m04 with (round(sf/rs/yf1,1)-round(sf1/rs1/12,1))/round(sf1/rs1/12*100,2),m12 with sb,sds with gs,sfje with sf,a with rs  
ENDIF
   go top
    browse field CJMC:H = ' ��λ���� ' , HJ :P = '@z 999,999;
,999.99' : 14 :H = '������ۼ�(Ԫ)' , PJ :P =;
 '@z 999,999' : 15 :H = 'ƽ��' , m01;
 :P = '@z 999,999.99':20 :H = '����' , m02;
 : 14 :H = '������%��', m12:P = '@z 999,999,999.99':h='���۴�����������',sds:P = '@z 999,999,999.99':h='��������˰',sfje:P = '@z 99,999,999.99':h='ʵ�����',jang:P = '@z 9999,999.99':h='˰���˾�����',m04;
 : 14 :H = 'ʵ����%��' noedit titl 'һ��������ۼƹ����ܶ�ͳ�Ʊ���λ��Ԫ��'
 copy to &bf.\������� fiel cjmc,jjhj,gzhj,hj,pj,m01,m02,m12,sds,sfje,jang,m04,a type xl5
 close all
 sele 2
 use zr11&LY1 &&�����ۼ���ϸ��
 inde on rybh to zr11&LY1
 sele 1
  use zr11&ND &&�����ۼ���ϸ��
  inde on RYBH to zr11&ND
  upda on rybh from B repl m01 with b.hj,m02 with b.pj
 **********************���������ݸ��µ����깤���ܶ��бȽ�*******��ʵ��*****************
  repl jang with PJ- m02,m04 with round(jang/m02*100,2) for m02>0
  inde on bmbh+rybh to zr11&ND
  go top
  browse field  CJMC : 12 :H = '�� ��',bmmc:12:h='�� ��',RYXM : 8 :H = '��  ��  ';
 , HJ :P = '@z 99,999,999.99' : 10 :H = ' �����ܶ� ' , m01 :P =;
 '@z 99,999,999.99' : 15 :H = '�����ܶ� ' , PJ :P = '@z 999,999' :;
 10 :H = ' ƽ������ ' , m02 :P = '@z 999,999' : 17 :H = ' ������ƽ��;
 ' , jang :P = '@z 999,999' : 17 :H  = '������������Ԫ��',m04 : 17 :H = '������%��';
 part 20 noedit titl '����ְ���������һ���� [�� F1 ������ĳһְ��]'    
  sort on hj to temp DESC
  use temp 
  go top
  browse field  CJMC : 12 :H = '�� ��',bmmc:12:h='�� ��',RYXM : 8 :H = '��  ��  ';
 , HJ :P = '@z 99,999,999.99' : 10 :H = ' �����ܶ� ' , m01 :P =;
 '@z 99,999,999.99' : 15 :H = '�����ܶ� ' , PJ :P = '@z 999,999' :;
 10 :H = ' ƽ������ ' , m02 :P = '@z 999,999' : 17 :H = ' ������ƽ��;
 ' , jang :P = '@z 999,999' : 17 :H  = '������������Ԫ��',m04 : 17 :H = '������%��',ylbx:h='���ϱ���',ybx:h='ҽ�Ʊ���',sybx:h='ʧҵ����',gjj:h='������',qynj:h='��ҵ���',sds:h='����˰',sfje:P = '@z 999,999.99':h='ʵ�����' ;
 part 20 noedit titl '���������ۼ�����һ���� [�� F1 ������ĳһְ��]'
 copy to &bf.\�����ۼ����� fiel rybh,cjmc,bmmc,ryxm,hj,m01,pj,m02,jang,m04,ylbx,ybx,sybx,gjj,qynj,sds,sfje
 copy to &bf.\�����ۼ����� fiel rybh,cjmc,bmmc,ryxm,hj,m01,pj,m02,jang,m04,ylbx,ybx,sybx,gjj,qynj,sds,sfje type xl5
 close all
***********************
use zr11&ND excl
yn3 = MESSAGEBOX("����Ǹɲ��μӹ����ܶ������Ƚ���",4+32,"��ʾ")
IF yn3 = 6
delete for val(zw)=Сְ��
pack
ENDI
inde on gwfl to zr11&ND
tota on gwfl to hztemp
use hztemp
repl pj with round(hj/a,0) all
inde on pj to hztemp
go top
brow fiel gwfl1:h='��Ա����',HJ :P = '@z 99,999;
,999.99' : 14 :H = '������ۼ�(Ԫ)' , PJ :P =;
 '@z 999,999' : 15 :H = 'ƽ��(Ԫ/��)' ,a:h='����'  noedit titl '����λ��Ա������ۼƹ����ܶ�ͳ�Ʊ���λ��Ԫ��'
     copy to &bf.\��Ա����������� fiel gwfl1,jjhj,gzhj,hj,pj,a type xl5
use zr11&ND
inde on zcdj to zr11&ND
tota on zcdj to hztemp
use hztemp
repl pj with round(hj/a,0) all
inde on pj to hztemp
go top
brow fiel zcdj1:h='ְ�ƣ�ְҵ�ʸ񣩵ȼ�',HJ :P = '@z 99,999;
,999.99' : 14 :H = '������ۼ�(Ԫ)' , PJ :P =;
 '@z 999,999' : 15 :H = 'ƽ��(Ԫ/��)' ,a:h='����'  noedit titl '��ְ�ƣ�ְҵ�ʸ񣩵ȼ�ͳ��������ۼƹ����ܶ��λ��Ԫ��'
 copy to &bf.\ְ�Ƶȼ�������� fiel zcdj1,jjhj,gzhj,hj,pj,a type xl5
close all
use zr11&LY1 excl
IF yn3 = 6
delete for val(zw)=Сְ��
pack
ENDI
repl PJ with round(jjHJ/mou,0) all
******�������**������**��Ч���ʷ���*****
repl cjdm with '77',cjmc with '���Ÿ�����' for val(zw)>��ְ�� and val(zw)<=Сְ��
repl cjdm with '88',cjmc with '��λ�쵼' for val(zw)<=��ְ��
index on CJDM to zr11&LY1 
total on CJDM to zr11 &&(����ͬ�ڻ�������*��Ч���ʷ���*��)
use zr11&ND
repl PJ with round(jjHJ/mou,0) all
******�������**������**��Ч���ʷ���*****
repl cjdm with '77',cjmc with '���Ÿ�����' for val(zw)>��ְ�� and val(zw)<=Сְ��
repl cjdm with '88',cjmc with '��λ�쵼' for val(zw)<=��ְ��
index on CJDM to zr11&ND
total on CJDM to zr1  &&(�������)
close all
select 2
use zr11
if yesno=6
   replace PJ with round(jjHJ/A,0) all      
 else 
   replace PJ with round(jjHJ/A/yf1,0) all   
endif
sum to LJ1,RS1 jjHJ,A
index on CJDM to zr11
select 1
use zr1
if yesno=6
   replace PJ with round(jjHJ/A,0) all      
 else 
   replace PJ with round(jjHJ/A/yf1,0) all   
endif 
sum jjHJ,A to LJ,RS
index on CJDM to zr1
upda on cjdm from B repl m02 with b.PJ
repl m01 with round(PJ-m02,0) , m03 with round(m01/m02*100,2) FOR m02>0
*********************************************�½�Ա���������ƽ���������***************
appe blank
repl cjdm with '99' for empty(cjdm)
go bott
repl cjmc with '�ϡ�������',jjHJ with LJ,pj with round(lj/rs,2),;
m01 with round(LJ/RS,1)-round(LJ1/RS1,1),m03 with (round(LJ/RS,1)-round(LJ1/RS1,1))/round(LJ1/RS1,1)*100 
go top
    browse field CJMC:H = ' ��λ���� ' , jjHJ :P = '@z 99,999;
,999.99' : 14 :H = '������ۼ�(Ԫ)' , PJ :P =;
 '@z 999,999' : 15 :H = 'ƽ��(Ԫ/��)' , m01:P = '@z 999,999';
 : 20 :H = '����������(Ԫ/��)' , m03;
 : 14 :H = '������%��' noedit titl '�ġ�������ۼƼ�Ч����ͳ�Ʊ���λ��Ԫ��'
 copy to &bf.\��Ч����ͳ�Ʒ��� fiel cjmc,jjhj,pj,m02,m01,m03,a type xl5
copy to jjtemp fiel cjmc,jjHJ,pj,m02,m01,m03
use ��Ч����ͳ�� excl
zap
appe from jjtemp
repl ���� with cjmc,��Ч���� with jjHJ,�˾� with pj,���ʶ� with m01,���� with m03 all
close all
 sele 2
 use zr11&LY1 &&�����ۼ���ϸ��
 inde on rybh to zr11&LY1
 sele 1
 use zr11&ND &&�����ۼ���ϸ��
 inde on RYBH to zr11&ND
 upda on rybh from B repl m01 with b.jjHJ,m02 with b.pj
 **********************���������ݸ��µ����깤���ܶ��бȽ�*******��ʵ��*****************
  repl jang with round((PJ-m02)/m02*100,2) for m02>0
  inde on bmbh+rybh to zr11&ND
  for zdz=1 to fcount()
  zd=fiel(zdz)
  zdz1=(fsize(zd)+fsize(zd))*2.4
  if uppe(allt(zd))='RYXM' 
     exit
  endif
  endfor 
  go top
  browse field  CJMC : 12 :H = '����λ',bmmc:12:h='�� ��',RYXM : 8 :H = '��  ��  ';
 , jjHJ :P = '@z 99,999,999.99' : 10 :H = ' ��Ч���� ' , m01 :P =;
 '@z 99,999,999.99' : 15 :H = '���꼨Ч���� ' , PJ :P = '@z 999,999' :;
 10 :H = ' ƽ����Ч���� ' , m02 :P = '@z 99,999' : 17 :H = ' ������ƽ��;
 ' ,  jang :H   =  '������������%��';
 part zdz1 noedit titl '�ġ�Ա����Ч���ʷ��估������� [�� F1 ������ĳһְ��]'
 sort on jjHJ to temp DESC
  use temp
  go top
  browse field  CJMC : 12 :H = '����λ',bmmc:12:h='�� ��',RYXM : 8 :H = '��  ��  ';
 , jjHJ :P = '@z 99,999,999.99' : 10 :H = ' ��Ч���� ' ,m01 :P =;
 '@z 99,999,999.99' : 15 :H = '���꼨Ч���� ' , PJ :P = '@z 999,999' :;
 10 :H = ' ƽ����Ч���� ' , m02 :P = '@z 99,999' : 17 :H = ' ������ƽ��;
 ' , jang :H   = '������������%��';
 part zdz1 noedit titl '�ġ�Ա����Ч�����ۼ����� [�� F1 ������ĳһְ��]'
 copy to  jjtemp fiel cjmc,bmmc,ryxm,jjHJ,m01,pj,m02,jang
use ��Ч�������� excl
zap
appe from jjtemp
repl ���� with cjmc,���� with bmmc,��Ч���� with jjHJ,ƽ�� with pj,ƽ�����ʶ� with pj-m02,���� with jang all
close all
use ryzk
copy to temp for 'ְ��'$ygxz1 and  '�ڸ�'$zgqk1 
close all
select 3
use zr&ND
copy to tempfx
use tempfx excl
yxyn = MESSAGEBOX("���Ÿ����������쵼�μ����״��������������",4+32,"��ʾ")
IF yxyn = 7
delete FOR val(zw)>��ְ�� and val(zw)<=Сְ��
pack
ENDI
index on RYBH to zr&ND
sele 2
use temp
index on RYBH to temp
select 1
use ry0 excl
zap
appe from temp
repl hj with 1 all
index on RYBH to ry0
upda on rybh from C repl hj with c.hj,sfje with c.sfje
set relation to RYBH into 2
replace ��� with '00001' , NLD with '35������' for   b.NLD='20������' or b.NLD='21��25��' or b.NLD='26��30��' or;
 b.NLD='31��35��'
replace ��� with '00002' , NLD with '36��50��' for b.NLD='36��40��' or b.NLD='41��45��' or b.NLD='46��50��'
replace ��� with '00003' , NLD with '51��60��' for b.NLD='51��55��' or b.NLD='56��60��'
close all
use ry0
replace ��� with '  1  ' for NLD='35������'
replace ��� with '  2  ' for NLD='36��50��'
replace ��� with '  3  ' for NLD='51��60��'
count for ('С'$WHCD2 and NLD='35������') or ('��'$WHCD2 and NLD='35������');
 to XC1
count for ('��'$WHCD2 and NLD='35������') or ('˶ʿ'$WHCD2 and NLD='35������');
 to DZ1
count for ('С'$WHCD2 and NLD='36��50��') or ('��'$WHCD2 and NLD='36��50��');
 to XC2
count for ('��'$WHCD2 and NLD='36��50��') or ('˶ʿ'$WHCD2 and NLD='36��50��');
 to DZ2
count for ('С'$WHCD2 and NLD='51��60��') or ('��'$WHCD2 and NLD='51��60��');
 to XC3
count for ('��'$WHCD2 and NLD='51��60��') or ('˶ʿ'$WHCD2 and NLD='51��60��');
 to DZ3
count for 'С'$WHCD2 or '��'$WHCD2 to XC
count for '��'$WHCD2 or '˶ʿ'$WHCD2 to DZ
sum to ������ , ������ , ZNL , ZGL  hj , A , NL , GL 
��ƽ�� = round(������/������,2)
count for hj<��ƽ��*0.75 and NLD='35������' to Z101
count for hj>=��ƽ��*0.75 and hj<��ƽ��*1.25 and NLD='35������';
 to Z102
count for hj>=��ƽ��*1.25 and hj<��ƽ��*1.75 and NLD='35������';
 to Z103
count for hj>=��ƽ��*1.75 and NLD='35������' to Z104
***********************************
count for hj<��ƽ��*0.75 and NLD='36��50��' to Z201
count for  hj>=��ƽ��*0.75 and hj<��ƽ��*1.25 and NLD='36��50��';
 to Z202
count for hj>=��ƽ��*1.25 and hj<��ƽ��*1.75 and NLD='36��50��';
 to Z203
count for hj>=��ƽ��*1.75 and NLD='36��50��' to Z204
***********************************
count for hj<��ƽ��*0.75 and NLD='51��60��' to Z301
count for hj>=��ƽ��*0.75 and hj<��ƽ��*1.25 and NLD='51��60��';
 to Z302
count for hj>=��ƽ��*1.25 and hj<��ƽ��*1.75 and NLD='51��60��';
 to Z303
count for hj>=��ƽ��*1.75 and NLD='51��60��' to Z304
***********************************
count for hj<��ƽ��*0.75 to Z01
count for hj>=��ƽ��*0.75 and hj<��ƽ��*1.25 to Z02
count for hj>=��ƽ��*1.25 and hj<��ƽ��*1.75 to Z03
count for hj>=��ƽ��*1.75 to Z04
************************************
index on NLD to ry0
total on NLD to ry1
use ry1
sort on ��� to ry2
use ry3 excl
zap
append from ry2
close all
select 2
use ry3
index on ��� to ry3
go top
select 1
use srfx excl
zap
append from ry3
index on ��� to srfx
upda on ��� from B  replace ����ṹ with b.NLD , ���������� with str(round(b.A/������*100;
,1),2,1)+'%' , ƽ������ with round(b.NL/b.A,1) , ƽ������ with round(b.GL/b.A;
,1) , ƽ��н�� with b.hj/b.A , ���������;
 with str(b.hj/������*100,2)+'%'
delete for empty(allt(����ṹ))
pack
append blank
replace ��� with '�� ��'
replace ռƽ��75�� with str(Z101/������*100,3,1)+'%' for ���='  1  '
replace ռƽ��75�� with str(Z201/������*100,3,1)+'%' for ���='  2  '
replace ռƽ��75�� with str(Z301/������*100,3,1)+'%' for ���='  3  '
replace ռƽ��75�� with str(Z01/������*100,3,1)+'%' for ���='�� ��'
replace ռ75_125�� with str(Z102/������*100,3,1)+'%' for ���='  1  '
replace ռ75_125�� with str(Z202/������*100,3,1)+'%' for ���='  2  '
replace ռ75_125�� with str(Z302/������*100,3,1)+'%' for ���='  3  '
replace ռ75_125�� with str(Z02/������*100,3,1)+'%' for ���='�� ��'
replace ռ125_175 with str(Z103/������*100,3,1)+'%' for ���='  1  '
replace ռ125_175 with str(Z203/������*100,3,1)+'%' for ���='  2  '
replace ռ125_175 with str(Z303/������*100,3,1)+'%' for ���='  3  '
replace ռ125_175 with str(Z03/������*100,3,1)+'%' for ���='�� ��'
replace ռ175���� with str(Z104/������*100,3,1)+'%' for ���='  1  '
replace ռ175���� with str(Z204/������*100,3,1)+'%' for ���='  2  '
replace ռ175���� with str(Z304/������*100,3,1)+'%' for ���='  3  '
replace ռ175���� with str(Z04/������*100,3,1)+'%' for ���='�� ��'
replace �������� with str(round(XC1/������*100,1),3,1)+'%' , �����Ļ� with;
 str(round((val(left(����������,2))/100*������-XC1-DZ1)/������*100,1),3,1)+'%';
 , ��ר�Ļ� with str(round(DZ1/������*100,1),3,1)+'%' for ���='  1  '
replace �������� with str(round(XC2/������*100,1),3,1)+'%' , �����Ļ� with;
 str(round((val(left(����������,2))/100*������-XC2-DZ2)/������*100,1),3,1)+'%';
 , ��ר�Ļ� with str(round(DZ2/������*100,1),3,1)+'%' for ���='  2  '
replace �������� with str(round(XC3/������*100,1),3,1)+'%' , �����Ļ� with;
 str(round((val(left(����������,2))/100*������-XC3-DZ3)/������*100,1),3,1)+'%';
 , ��ר�Ļ� with str(round(DZ3/������*100,1),3,1)+'%' for ���='  3  '
replace ����ṹ with 'ȫ��Ա��' , ��������� with '100%' , ���������� with;
 str(������,3)+'��' , ƽ������ with round(ZNL/������,2) , ƽ������ with round(ZGL/������;
,2) , ƽ��н�� with round(������/������;
,2) , �������� with str(XC/������*100,3,1)+'%' , �����Ļ� with str((������-XC-DZ)/������*100;
,3,1)+'%' , ��ר�Ļ� with str(DZ/������*100,3,1)+'%' for ���='�� ��'
browse noedit titl '�塢������乫ƽ�����������״����ߡ����ݱ�'
wjm='&bf'+'\���빫ƽ����'
copy to &wjm type xl5 
close all
sele 2
use ry0
inde on rybh to ry0
sele 1
use temp
repl gjj with 1 all
inde on rybh to temp
upda on rybh from B repl gjj with b.hj
close all
on key label F1
IF fbl=1
DEFINE WINDOW temp FROM INT((SROWS()-65)/2)  , INT((SCOLS()-310)/2)  TO INT((SROWS()+65)/2) , INT((SCOLS()+310)/2)  FLOAT CLOSE SHADOW DOUBLE title '����Ա�����밴˰ǰӦ��н��Ƚ� [�ӡ��ļ����˵��е㡰���ҡ�ѡ����Ҹ���]'
ELSE
DEFINE WINDOW temp FROM INT((SROWS()-40)/2)  , INT((SCOLS()-160)/2)  TO INT((SROWS()+40)/2) , INT((SCOLS()+160)/2)  FLOAT CLOSE SHADOW DOUBLE title '����Ա�����밴˰ǰӦ��н��Ƚ� [�ӡ��ļ����˵��е㡰���ҡ�ѡ����Ҹ���]'
ENDIF
activate window temp
sele gjj as н�� , cjmc as ����,bmmc as ����,rybh as ���,ryxm as ����,whcd2 as ѧ��,zyfl as רҵ,zc as ְ��,�ڼ�,;
gz1 as ����,gw as ��λ,�ڵ�,zw1 as ְ��,nl as ����,nld as �����,gl as ����,gld as ����� from temp ORDER BY gjj DESC 
clear window temp
close all
sele 2
use ry0
inde on rybh to ry0
sele 1
use temp
repl gjj with 1 all
inde on rybh to temp
upda on rybh from B repl gjj with b.sfje
close all
on key label F1
IF fbl=1
DEFINE WINDOW temp FROM INT((SROWS()-65)/2),INT((SCOLS()-310)/2)  TO INT((SROWS()+65)/2) , INT((SCOLS()+310)/2)  FLOAT CLOSE SHADOW DOUBLE title '����Ա�����밴˰��Ӧ��н��Ƚ� [�ӡ��ļ����˵��е㡰���ҡ�ѡ����Ҹ���]'
ELSE
DEFINE WINDOW temp FROM INT((SROWS()-40)/2),INT((SCOLS()-160)/2)  TO INT((SROWS()+40)/2) , INT((SCOLS()+160)/2)  FLOAT CLOSE SHADOW DOUBLE title '����Ա�����밴˰��ʵ��н��Ƚ� [�ӡ��ļ����˵��е㡰���ҡ�ѡ����Ҹ���]'
ENDIF
activate window temp
sele gjj as н�� , cjmc as ����,bmmc as ����,rybh as ���,ryxm as ����,whcd2 as ѧ��,zyfl as רҵ,zc as ְ��,�ڼ�,;
gz1 as ����,gw as ��λ,�ڵ�,zw1 as ְ��,nl as ����,nld as �����,gl as ����,gld as ����� from temp ORDER BY gjj DESC 
clear window temp
close all
use ryzk
copy to temp
if file('zr'+LY1+'.dbf')
   sele 2
   use zr&LY1
   copy to zrtemp
   inde on rybh to zr&LY1
   BSND=LY1+'��Ӧ������'
   do bs
   sele 2
   use zrtemp  
   inde on rybh to zrtemp
   sum sfje to sf
   if sf>0
      repl hj with sfje all
      BSND=LY1+'��ʵ���������ʻ�����'
      do bs
   endif    
endif
close all
use ryzk
copy to temp
if file('zr'+ND+'.dbf')
   sele 2
   use zr&ND
   copy to zrtemp
   inde on rybh to zr&ND
   BSND=ND+'��Ӧ������'
   do bs
   sele 2
   use zrtemp
   inde on rybh to zrtemp
   sum sfje to sf
   if sf>0
      repl hj with sfje all
      BSND=ND+'��ʵ���������ʻ�����'
      do bs
   endif    
endif
clear
yn = MESSAGEBOX("���ݵ����ɹ����ִ򿪵��ӱ�༭ʹ����",4+32,"��ʾ")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\�������.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\�����ۼ�����.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\��Ա�����������.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\ְ�Ƶȼ��������.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\��Ч����ͳ�Ʒ���.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\���빫ƽ����.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
ENDI  
close all
delete files zr11temp.idx
delete files temp*.*
delete files srzf.dbf




proc bs
sele 1
use temp
repl gjj with 1 all
inde on rybh to temp
upda on rybh from B repl gjj with b.hj
close all
IF fbl=1
DEFINE WINDOW temp FROM INT((SROWS()-65)/2),INT((SCOLS()-310)/2)  TO INT((SROWS()+65)/2) , INT((SCOLS()+310)/2)  FLOAT CLOSE SHADOW DOUBLE title 'Ա�����밴��н�Ƚ�('+BSND+')[�ӡ��ļ����˵��е㡰���ҡ�ѡ����Ҹ���]'
ELSE
DEFINE WINDOW temp FROM INT((SROWS()-40)/2),INT((SCOLS()-160)/2)  TO INT((SROWS()+40)/2) , INT((SCOLS()+160)/2)  FLOAT CLOSE SHADOW DOUBLE title 'Ա�����밴��н�Ƚ�('+BSND+')[�ӡ��ļ����˵��е㡰���ҡ�ѡ����Ҹ���]'
ENDIF
activate window temp
sele gjj as ��н , cjmc as ����,bmmc as ����,rybh as ���,ryxm as ����,whcd2 as ѧ��,zyfl as רҵ,zc as ְ��,�ڼ�,;
gz1 as ����,gw as ��λ,�ڵ�,zw1 as ְ��,nl as ����,nld as �����,gl as ����,gld as ����� from temp ORDER BY gjj DESC 
clear window temp
on key label F1 do grcx.fxp