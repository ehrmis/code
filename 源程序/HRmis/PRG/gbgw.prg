***************************
* .\GBGW.PRG
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
select * from ��λ���� into dbf rygwzk where gwfl1<>'������λ'
do dhkDY
select 3
use rygwzk
repl gwfl1 with '������Ա' for ryfl='04'
repl gwfl1 with 'רҵ��Ա' for ryfl<>'04'
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
  set printer to �����λ.TXT &&Ԥ��
  set printer on 
endi
do while not eof()
  CJBH = alltrim(cjDM)
  do case
  case DH=1
    select * from rygwzk into table rygw where CJDM='&CJBH'
  case DH=2
    select * from rygwzk into table rygw where CJDM='&DYCJ'
  case DH=3
    select * from rygwzk into table rygw where BMBH='&DYBM'
  endcase
  use
  select 1
  use rygw
  count for RYFL='04' to GLRY
  count for RYFL<>'04' to zyry
  count for allt(gwjb1)='ʮ�߸�' to gj17
  count for allt(gwjb1)='ʮ����' to gj16
  count for allt(gwjb1)='ʮ�ĸ�' to gj14
  count for allt(gwjb1)='ʮ����' to gj13
  count for allt(gwjb1)='ʮһ��' to gj11
  count for allt(gwjb1)='ʮ��' to gj10
  count for allt(gwjb1)='�Ÿ�' to gj9
  count for allt(gwjb1)='�˸�' to gj8
  count for allt(gwjb1)='�߸�' to gj7
  count for allt(gwjb1)='����' to gj6
  count for allt(gwjb1)='���' to gj5
  count for allt(�ڵ�)='��ְ(Ա)' to YJ
  count for allt(�ڵ�)='��ְ(��)' to ZJ
  count for allt(�ڵ�)='��ְ' to ZZ
  count for allt(�ڵ�)='����ְ' to FGZ
  count for allt(�ڵ�)='����ְ' to ZGZ
  count for allt(�ڵ�)='������' to zCJ
  count for allt(�ڵ�)='������' to fCJ
  count for allt(�ڵ�)='���Ƽ�' to ZK
  count for allt(�ڵ�)='���Ƽ�' to FK
  count for allt(�ڵ�)='һ����Ա' to YJKY
  count for allt(�ڵ�)='������Ա' to LJKY
  count for allt(�ڵ�)='������Ա' to SJKY
  ZRS = reccount()
  PG = 1
  go top
do while  not eof()
    @ dwh , 0 say space(10)+'&MC111'+'ְ����λ������ϸ��'  font '����',20
    @ prow()+2 , 0 say space(37)+'( �����λ��רҵ������λ )' font '����',12
    @ prow()+2 , 8 say '�� λ:'+alltrim(cjmc)+space(28)+'��'+str(PG,2)+'ҳ'+space(22)+'����:'+str(year(date()),4)+'��'+str(month(date()),2)+'��'+str(day(date()),2)+'��' &zt1
    T1 = '�������������ש��������ө��������������������������������������������ө��������������ө��������ө�����'
    T2 = '            ��        ��           ��  ��  ��  ��  ְ  ��           ��              ��        ��      '
    T3 = ' ����(����) �� ��  �� �������Щ����Щ����Щ��Щ��Щ��Щ��Щ��Щ��Щ����� �� �� �� λ����Ա���੦�ڼ�  '
    T4 = '            ��        �� Ա �� �� �� �� ���ߩ����������������������Ʃ�              ��        ��      '
    T5 = '  ��   ��   ��        �� ְ �� ְ �� ְ ��ְ�������Ʃ��Ʃ��ܩ��쩦Ա��  (  ����  )  ��        ��      '
    T6 = '�������������橤�������੤���੤���੤���੤�੤�੤�੤�੤�੤�੤�੤�������������੤�������੤����'
    T7 = '�������������橤�������੤���ة����ة����ة��ة��ة��ة��ة��ة��ة��੤�������������੤�������੤����'
    T8 = '�������������ߩ��������۩��������������������������������������������۩��������������۩��������۩�����'
    T9 = '�������������ש��������ө����ө����ө����ө��ө��ө��ө��ө��ө��ө��ө��������������ө��������ө�����'
    @ prow()+1 , 0 say T1 &zt
    @ prow()+1 , 0 say T2 &zt
    @ prow()+1 , 0 say T3 &zt
    @ prow()+1 , 0 say T4 &zt
    @ prow()+1 , 0 say T5 &zt
    @ prow()+1 , 0 say T6 &zt
    N = 1
    do while N<yjl
    IF  not eof()
        @ prow()+1 , 0 say bmmc+'��' &zt
        @ prow() , pcol() say ryxm &zt
      if allt(�ڵ�)='��ְ(Ա)' 
         @ prow() , pcol() say '�� �� ��    ��    ��  ��  ��  ��  ��  ��  ��  ��' &zt
      endif
      if allt(�ڵ�)='��ְ(��)'
         @ prow() , pcol() say '��    �� �� ��    ��  ��  ��  ��  ��  ��  ��  ��' &zt
      endif
      if allt(�ڵ�)='��ְ'
         @ prow() , pcol() say '��    ��    �� �� ��  ��  ��  ��  ��  ��  ��  ��' &zt
      endif
      if '��ְ'$�ڵ�
         @ prow() , pcol() say '��    ��    ��    ���̩�  ��  ��  ��  ��  ��  ��' &zt
      endif
      if '����'$�ڵ�
         @ prow() , pcol() say '��    ��    ��    ��  ���̩�  ��  ��  ��  ��  ��' &zt
      endif
      if allt(�ڵ�)='���Ƽ�'
         @ prow() , pcol() say '��    ��    ��    ��  ��  ���̩�  ��  ��  ��  ��' &zt
      endif
      if allt(�ڵ�)='���Ƽ�'
         @ prow() , pcol() say '��    ��    ��    ��  ��  ��  ���̩�  ��  ��  ��' &zt
      endif
      if allt(�ڵ�)='ҵ������'
         @ prow() , pcol() say '��    ��    ��    ��  ��  ��  ��  ���̩�  ��  ��' &zt
      endif
      if allt(�ڵ�)='ҵ������'
         @ prow() , pcol() say '��    ��    ��    ��  ��  ��  ��  ��  ���̩�  ��' &zt
      endif
      if allt(�ڵ�)='ҵ��Ա'
         @ prow() , pcol() say '��    ��    ��    ��  ��  ��  ��  ��  ��  ���̩�' &zt
      endif
      if !empty(zw)
         @ prow() , pcol() say left(gw,14)+'��' &zt
      else
         @ prow() , pcol() say gz1+'  ��' &zt
      endif
      @ prow() , pcol() say left(gwfl1,8)+'��' &zt
      @ prow() , pcol() say gwjb1 &zt
      @ prow()+1 , 0 say T6 &zt
     ENDIF
      if mod(N,yjl)=0
        @ prow()+1 , 0 say T7 &zt
      endif
      if  not eof()
        N = N+1
        skip
      else
        @ prow()+1 , 0 say ' С  ��(��) ' &zt
        @ prow() , pcol() say '��'+str(glry+zyry,3) &zt
        @ prow() , pcol() say '     ��' &zt
        @ prow() , pcol() say yJ picture '@Z 9999' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say zJ picture '@Z 9999' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say zz picture '@Z 9999' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say zgz+fgz picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say Zcj+fcj picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say Zk picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say fk picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say yjKY picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say ljky picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say sjky picture '@Z 99' &zt
        @ prow() , pcol() say '��������Ա:' &zt
        @ prow() , pcol() say GLRY picture '@Z 999' &zt
        @ prow() , pcol() say ' רҵ������Ա:' &zt
        @ prow() , pcol() say zyry picture '@Z 999' &zt
        @ prow()+1 , 0 say T7 &zt
@ prow()+1 , 0 say ' ��  ��(��) ��'+str(yj+zj+zz+fgz+zgz+zcj+fcj+zk+fk+yjky+ljky+sjky,3)+'     ��'+space(9)+'������λ:'+str(yj+zj+zz+fgz+zgz,3)+'  �����λ:'+str(zcj+fcj+zk+fk+yjky+ljky+sjky,3)+'         ��' &zt
@ prow() , pcol() say 'ƽ����:'+str(round((gj5*5+gj6*6+gj7*7+gj8*8+gj9*9+gj10*10+gj11*11+gj13*13+gj14*14+gj16*16+gj17*17)/ZRS,1),4,1)+'��' &zt
@ prow()+1 , 0 say T8 &zt
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
  count for RYFL='04' to GLRY
  count for RYFL<>'04' to zyry
  count for allt(gwjb1)='ʮ�߸�' to gj17
  count for allt(gwjb1)='ʮ����' to gj16
  count for allt(gwjb1)='ʮ�ĸ�' to gj14
  count for allt(gwjb1)='ʮ����' to gj13
  count for allt(gwjb1)='ʮһ��' to gj11
  count for allt(gwjb1)='ʮ��' to gj10
  count for allt(gwjb1)='�Ÿ�' to gj9
  count for allt(gwjb1)='�˸�' to gj8
  count for allt(gwjb1)='�߸�' to gj7
  count for allt(gwjb1)='����' to gj6
  count for allt(gwjb1)='���' to gj5
  count for allt(�ڵ�)='��ְ(Ա)' to YJ
  count for allt(�ڵ�)='��ְ(��)' to ZJ
  count for allt(�ڵ�)='��ְ' to ZZ
  count for allt(�ڵ�)='����ְ' to FGZ
  count for allt(�ڵ�)='����ְ' to ZGZ
  count for allt(�ڵ�)='������' to zCJ
  count for allt(�ڵ�)='������' to fCJ
  count for allt(�ڵ�)='���Ƽ�' to ZK
  count for allt(�ڵ�)='���Ƽ�' to FK
  count for allt(�ڵ�)='ҵ������' to YJKY
  count for allt(�ڵ�)='ҵ������' to LJKY
  count for allt(�ڵ�)='ҵ��Ա' to SJKY
ZRS = reccount()
@ prow()+1 , 0 say T9 &zt
@ prow()+1 , 0 say '��   ��(��) ��' &zt
        @ prow() , pcol() say str(glry+zyry,3) &zt
        @ prow() , pcol() say '     ��' &zt
        @ prow() , pcol() say yJ picture '@Z 9999' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say zJ picture '@Z 9999' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say zz picture '@Z 9999' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say zgz+fgz picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say Zcj+fcj picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say Zk picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say fk picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say yjKY picture '@Z 99' &zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say ljky picture '@Z 99'&zt
        @ prow() , pcol() say '��' &zt
        @ prow() , pcol() say sjky picture '@Z 99' &zt
        @ prow() , pcol() say '��������Ա:' &zt
        @ prow() , pcol() say GLRY picture '@Z 999' &zt
        @ prow() , pcol() say ' רҵ������Ա:' &zt
        @ prow() , pcol() say zyry picture '@Z 999' &zt
        @ prow()+1 , 0 say T7 &zt
@ prow()+1 , 0 say '��   ��(��) ��'+str(yj+zj+zz+fgz+zgz+zcj+fcj+zk+fk+yjky+ljky+sjky,3)+'     ��'+space(9)+'������λ:'+str(yj+zj+zz+fgz+zgz,3)+'  �����λ:'+str(zcj+fcj+zk+fk+yjky+ljky+sjky,3)+'         ��' &zt
@ prow() , pcol() say 'ƽ����:'+str(round((gj5*5+gj6*6+gj7*7+gj8*8+gj9*9+gj10*10+gj11*11+gj13*13+gj14*14+gj16*16+gj17*17)/ZRS,1),4,1)+'��' &zt
@ prow()+1 , 0 say T8 &zt
@ prow()+1 , 0 say space(10)+'���ʸ�����:'+'&RS111'+space(45)+'�Ʊ���:'+'&ZB111' &zt1
set device to screen
if dh1=2
   set printer to lpt1
endif
set printer off
set printer to
close all
if dh1=1
   clear
   modi file �����λ.txt
endi
clear windows
delete file rygw.dbf
clear
return
