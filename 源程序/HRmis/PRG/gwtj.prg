***********
* gwtj.prg
***********
dhgw=''
dh14=''
dh1a=''
dhgwdy=''
YF1=month(date())
ND1=year(date())
ND = str(ND1,4)
if YF1>9
  YF = str(YF1,2)
else
  YF = '0'+str(YF1,1)
endif
NY = ND+YF
********��ǰϵͳ����*************
NYR = Right(NY,4)
YFs = iif(val(yf)>9,str(val(yf)-1,2),'0'+str(val(yf)-1,1))
IF val(yf)=1
   NY1 = str(year(date())-1,4)+'12'
   YFs = '12'
else   
   if val(yf)=10
      NY1 = str(year(date()),4)+'09'
   ELSE
      NY1 = str(year(date()),4)+YFs
    endif   
ENDIF
********��ǰϵͳ����[���±��ʽ]*************
if !file('gw'+ny+'.dbf')
=MESSAGEBOX('�������ά�����¹��ָ�λ��', 16,"��ʾ")
 return
endif
clear
do dhk14
do case
case dh14=1
******************************ͳ�Ʊ���*********************
use gw&NY
coun for sfgj>0 to �ϸ�
use ��λͳ��
COPY TO temp FOR gwdm<>'lw'
***************����������ǲ����λ����*********************
USE temp
sum ���� to rs
sum val(gwdm)*���� to gs
sum ���� for '����'$gwfl1 to cz
sum ���� for '����'$gwfl1 to js
sum ���� for '����'$gwfl1 to glry
go top
brow field ���,gwfl1:h='��λ����',gwjb1:h='�ϸ�ǰ����',gwdc1:h='����',�ڵ�:h='��λ�ȼ�',bzgz:h='�ϸ����λ����',����:4 noedit titl 'һ;
����λ���(������λ:'+str(cz,3)+'�� �����λ:'+str(glry,3)+'�� ������λ:'+str(js,3)+'��  �ϼ�:'+str(rs,4)+'�� ƽ��:'+str(round(gs/rs,1),4,1)+'�� �ϸ�:'+str(�ϸ�,3)+'�� )'
use ��������
COPY TO temp FOR gwdm<>'lw'
***************����������ǲ����λ����*********************
USE temp
repl ��� with str(recn(),4) all
sum ���� to zcc for gwfl1='������λ'
sum ���� to jss for gwfl1='������λ'
go top
brow field ���,gwfl1:h='��λ����',gwjb1:h='�ϸ�ǰ����',gwdc1:h='����',�ڵ�:h='��λ�ȼ�',bzf:h='ע��೤',gwgz:h='�ϸ����λ����',����:4 noedit titl '������֤������Ա����('+str(zcc+jss,4)+'��):������λ('+str(zcc,4)+'��) ������λ('+str(jss,3)+'��)'
use ��֤����
repl ��� with str(recn(),4) all
sum ���� to wzry
go top
brow field ���,gwfl1:h='��λ����',gwjb1:h='�ϸ�ǰ����',gwdc1:h='����',bzf:h='ע��೤',bzgz:h='�ϸ����λ����',����:4 noedit titl '������֤��Ա����('+str(wzry,3)+'��)'

use �������
repl ��� with str(recn(),4) all
sum ���� to gll for gwfl1='�����λ'
sum ���� to zyy for gwfl1='������λ'
go top
brow field ���,gwfl1:h='��λ����',gwjb1:h='�ϸ�ǰ����',gwdc1:h='����',�ڵ�:h='��������ְ��',bzgz:h='�ϸ����λ����',����:4 noedit titl '�ġ�������Ա����('+str(gll+zyy,3)+'��):�����λ('+str(gll,3)+'��) ������λ('+str(zyy,3)+'��)'

use ����ͳ��
COPY TO temp FOR gwdm<>'lw'
***************����������ǲ����λ����*********************
USE temp 
sum a to rs
go top
browse field gz:r:h='����' , gz1:r:h='����' ,gw:h='��λ':r,gwjb1:h='�ڼ�':r,�ڵ�:r,a:h='����':r:4 titl '�塢���ָ�λ�ֲ�(�ϼ�:'+str(rs,4)+'��)'
if file('��ѵ����.dbf')
use ��ѵ����
sum a to rs
go top
browse field gwfl1:h='��λ����':r,gz1:r:h='����' ,gw:h='��λ':r,pxmc:h='��ѵ����':r,pxdj:h='��ѵ�ȼ�':r,a:h='����':r:4 titl '���������ָ�λ��ѵ���(�ϼ�:'+str(rs,4)+'��)'
endif
case dh14=2
do dhkdy2
if dh1a=1
   do dhkgwdy
   if dhgwdy=1
      do grgw
   else
      do gbgw
   endi
else
   do ��λͳ��
endif
case dh14=3
retu
endcase
close all
clear
delete file gzhz.dbf
delete file gzhz.idx
return
