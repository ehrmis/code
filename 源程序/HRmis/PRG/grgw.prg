***************************
* .\GRGW.PRG
***************************
close all
CJBH = ''
DH = ''
DH1 = ''
DYCJ = ''
DYBM = ''
dwh=3
yjl=30
zt="font '����',11"
zt1="font '����',12"
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
store 0 to gj1 , gj2 , gj3 , gj4 , gj5 , gj6 , gj7 , gj8 , gj9 , sf
select * from ��λ���� where gwfl1='������λ' into dbf ccrygwzk
close all
do DHKDY
select 3
use ccrygwzk
select 2
use cjk
go top
do dhkdy1
wait wind "���ڴ�ӡ�����Ժ򡤡���������" nowa
if dh1=2
  do srdwh.spr
  set device to printer
  set printer on
else
  set device to printer
  set printer to ������λ.TXT &&Ԥ��
  set printer on 
endi
do while  not eof()
  CJBH = alltrim(cjDM)
  do case
  case DH=1
    select * from ccrygwzk into dbf rygw where CJDM='&CJBH'
  case DH=2
    select * from ccrygwzk into dbf rygw where CJDM='&DYCJ'
  case DH=3
    select * from ccrygwzk into dbf rygw where BMBH='&DYBM'
  endcase
  use
  select 1
  use rygw
  count for allt(gwjb1)='�Ÿ�' to gj9
  count for allt(gwjb1)='�˸�' to gj8
  count for allt(gwjb1)='�߸�' to gj7
  count for allt(gwjb1)='����' to gj6
  count for allt(gwjb1)='���' to gj5
  count for allt(gwjb1)='�ĸ�' to gj4
  count for allt(gwjb1)='����' to gj3
  count for allt(gwjb1)='����' to gj2
  count for allt(gwjb1)='һ��' to gj1
  ZRS = reccount()
  count for sfgj>0 to sf
  PG = 1
  go top
  do while  not eof()
    @ dwh , 0 say space(10)+'&MC111'+'ְ����λ������ϸ��' font '����',20
    @ prow()+2 , 0 say space(45)+'( ������λ )' font '����',12
    @ prow()+2 , 8 say '�� λ:'+alltrim(cjmc)+space(30)+'��'+str(PG,2)+'ҳ'+space(25)+'����:'+str(year(date()),4)+'��'+str(month(date()),2)+'��'+str(day(date()),2)+'��' &zt1
    T1 = '�������������ש��������ө��������ө������ө������������ө����ө��������ө������������ө����������������ө�����'
    T2 = ' ��������   �� ��  �� ����ǰ�ڼ�����  ����ְҵ�ʸ���ѵ���೤�������ȼ��� �ִ��¹��� �� �� �� �� �� λ �� ���� '
    T3 = '�������������橤�������੤�������੤�����੤�����������੤���੤�������੤�����������੤���������������੤����'
    T4 = '�������������ߩ��������۩��������۩������۩������������۩����۩��������۩������������۩����������������۩�����'
    @ prow()+1 , 0 say T1 &zt
    @ prow()+1 , 0 say T2 &zt
    @ prow()+1 , 0 say T3 &zt
    N = 1
    do while N<yjl
     if  not eof()
      @ prow()+1 , 0 say bmmc+'��' &zt
      @ prow() , pcol() say ryxm &zt
      @ prow() , pcol() say '��'+space(1) &zt
      @ prow() , pcol() say gwjb1 &zt
      @ prow() , pcol() say space(1)+'��' &zt
      @ prow() , pcol() say gwdc1 &zt
      @ prow() , pcol() say '��' &zt
      @ prow() , pcol() say left(pxmc,12)+'��' &zt
      if !empty(bzf)
         @ prow() , pcol() say ' �� ��' &zt
      else
         @ prow() , pcol() say '    ��' &zt
      endif
      @ prow() , pcol() say �ڵ� &zt 
      @ prow() , pcol() say '��' &zt     
      @ prow() , pcol() say left(gz1,12)+'��'+left(gw,16) &zt
      @ prow() , pcol() say '��'+str(GL,2,0)+'��' &zt
      @ prow()+1 , 0 say T3 &zt
     endif
     if mod(N,yjl)=0
        @ prow()+1 , 0 say T4 &zt
      endif
      if  not eof()
        N = N+1
        skip
      else       
        @ prow()+1 , 0 say ' С  ��(��) ' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say ZRS picture '@ 99999999' &zt
        @ prow() , pcol() say '��' &zt
@ prow() , pcol() say 'ƽ���ڣ�'+str(round((gj9*9+gj8*8+gj7*7+gj6*6+gj5*5+gj4*4+gj3*3+gj2*2+gj1)/ZRS;
,1),4,1)+'��   �ϸ���'+str(sf,3)+'��' &zt
        @ prow()+1 , 0 say T4 &zt
        exit
      endif
    enddo
    PG = PG+1
  enddo
  if DH<>1
    exit
  endif
  select 2
  skip
enddo
select 3
  count for allt(gwjb1)='�Ÿ�' to gj9
  count for allt(gwjb1)='�˸�' to gj8
  count for allt(gwjb1)='�߸�' to gj7
  count for allt(gwjb1)='����' to gj6
  count for allt(gwjb1)='���' to gj5
  count for allt(gwjb1)='�ĸ�' to gj4
  count for allt(gwjb1)='����' to gj3
  count for allt(gwjb1)='����' to gj2
  count for allt(gwjb1)='һ��' to gj1
  count for sfgj>0 to sf
ZRS = reccount()
@ prow()+1 , 0 say T1 &zt
@ prow()+1 , 0 say ' ��      �� ��' &zt
@ prow() , pcol() say ZRS picture '@ 99999999' &zt
@ prow() , pcol() say '��' &zt
@ prow() , pcol() say 'ƽ����:'+str(round((gj9*9+gj8*8+gj7*7+gj6*6+gj5*5+gj4*4+gj3*3+gj2*2+gj1)/ZRS;
,1),4,1)+'��   �ϸ���'+str(sf,3)+'��' &zt
@ prow()+1 , 0 say T4 &zt
@ prow()+1 , 0 say space(20)+'���ʸ�����:'+'&RS111'+space(45)+'�Ʊ���:'+'&ZB111' &zt1
set device to screen
if dh1=2
   set printer to lpt1
endif
set printer off
set printer to
close all
if dh1=1
   clear
   modi file ������λ.txt
   *run notepad ������λ.txt
endi
clear windows
delete file rygw.dbf
return
