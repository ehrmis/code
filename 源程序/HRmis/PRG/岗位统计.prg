***********
* ��λͳ��
***********
YF1=month(date())
ND1=year(date())
do srnd.spx
NY = ND+YF
NYR=right(NY,4)
NYR = Right(NY,4)
YFs = iif(month(date())>9,str(month(date())-1,2),'0'+str(month(date())-1,1))
NY1 = str(year(date()),4)+YFs
if month(date())=1
   NY1 = str(year(date())-1,4)+'12'
   YFs = '12'
endi
if month(date())=10
   NY1 = str(year(date()),4)+'09'
endi
if !file('gw'+ny+'.dbf')
=MESSAGEBOX('�������ά�����¹��ָ�λ��', 16,"��ʾ")
 return
endif
clear
zt="font '����',11"
zt1="font '����',10"
set device to printer
set printer on
YN = MESSAGEBOX('��Ҫ��ӡ��λ�ȼ���˱��𣿣�',4+32,'��ʾ')
if yn=6
t1='�������ө��������ө������ө������ө��������ө��������ө��������ө�����'
t2='����ũ���λ���੦ ���� �� ���� ����λ�ȼ���������ʩ����ʼ���������'
t3='�ĩ����੤�������੤�����੤�����੤�������੤�������੤�������੤����'
t4='�������۩��������۩������۩������۩��������۩��������۩��������۩�����'
use ��λͳ��
sum ���� for '����'$gwfl1 to cz
sum ���� for '����'$gwfl1 to js
sum ���� for '����'$gwfl1 to glry
sum ���� to rs
sum val(gwdm)*���� to gs
go top
xh=1
do while  not eof()
      @ 0 , 0 say space(5)+'&MC111'+'��λ�ȼ���˱�(��һ)' font  '����',11
      @ 2 , 0 say T1 &zt
      @ 3 , 0 say T2 &zt
      @ 4 , 0 say T3 &zt
  do while  not eof()
  @ prow()+1 , 0 say '��' &zt
  @ prow() , pcol() say str(xh,4) &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwfl1 &zt
  @ prow() , pcol() say '��'&zt
  @ prow() , pcol() say gwjb1 &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwdc1 &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say �ڵ� &zt
  @ prow() , pcol() say '��  ' &zt
  @ prow() , pcol() say gwgz &zt
  @ prow() , pcol() say '  �� ' &zt
  @ prow() , pcol() say �ڼ� &zt
  @ prow() , pcol() say ' ��' &zt
  @ prow() , pcol() say ���� picture '@Z 9999' &zt
  @ prow() , pcol() say '��' &zt
  skip
  xh=xh+1
  if mod(xh,26)=0
     @ prow()+1 , 0 say t4 &zt
     exit
  else
     @ prow()+1 , 0 say T3 &zt
  endif
  enddo
enddo
  @ prow()+1 , 0 say '��' &zt
  @ prow() , pcol() say '�ϼƩ�' &zt
  @ prow() , pcol() say '������λ:'+str(cz,3)+'�� ������λ:'+str(js,3)+'�� �����λ:'+str(glry,2) &zt
  @ prow() , pcol() say '�� ƽ��' &zt
  @ prow() , pcol() say str(round(gs/rs,1),4,1) &zt
  @ prow() , pcol() say '�ک�' &zt
  @ prow() , pcol() say rs picture '@Z 9999' &zt
  @ prow() , pcol() say '��' &zt
  @ prow()+1 , 0 say t4 &zt
  @ prow()+1 , 10 say '�����:'+alltrim(str(year(date()),4))+'��'+alltrim(str(month(date());
,2))+'��'+alltrim(str(day(date()),2))+'��' &zt1
@ prow() , pcol() say '                                ���Ա:&ZB111' &zt1
set printer to lpt1
endif
?
YN = MESSAGEBOX('��Ҫ��ӡ��֤������Ա��λ�ֲ����𣿣�',4+32,'��ʾ')
if yn=6
t1='�������ө��������ө������ө������ө��������ө��������ө��������ө����ө�����'
t2='����ũ���λ���੦ ���� �� ���� ����λ�ȼ���������ʩ����ʼ��𩦰೤��������'
t3='�ĩ����੤�������੤�����੤�����੤�������੤�������੤�������੤���੤����'
t4='�������۩��������۩������۩������۩��������۩��������۩��������۩����۩�����'
use ��������
sum ���� to rs
sum val(gwdm)*���� to gs
sum ���� for '����'$gwfl1 to cz
sum ���� for '����'$gwfl1 to js
sum ���� for !empty(bzf) to bz
go top
xh=1
do while  not eof()
      @ 0 , 0 say space(5)+'&MC111'+'��λ�ȼ�ͳ�Ʊ������' font '����',20
      @ 2 , 0 say '                              ( ��֤������Ա��λ�ֲ� )' &zt1
      @ 3 , 0 say T1 &zt
      @ 4 , 0 say T2 &zt
      @ 5 , 0 say T3 &zt
do while  not eof()
  @ prow()+1 , 0 say '��' &zt
  @ prow() , pcol() say str(xh,4) &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwfl1 &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwjb1 &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwdc1 &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say �ڵ� &zt
  @ prow() , pcol() say '��  ' &zt
  @ prow() , pcol() say gwgz &zt
  @ prow() , pcol() say '  �� ' &zt
  @ prow() , pcol() say �ڼ� &zt
  @ prow() , pcol() say ' ��' &zt
if !empty(bzf)
   @ prow() , pcol() say ' �� ��' &zt
else
   @ prow() , pcol() say '    ��' &zt
endif
  @ prow() , pcol() say ���� picture '@Z 9999' &zt
  @ prow() , pcol() say '��' &zt
  skip
  xh=xh+1
  if mod(xh,26)=0
     @ prow()+1 , 0 say t4 &zt
     exit
  else
     @ prow()+1 , 0 say T3 &zt
  endif
  enddo
enddo
  @ prow()+1 , 0 say '��' &zt
  @ prow() , pcol() say '�ϼ�' &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say '������λ:'+str(cz,3)+'�� ������λ:'+str(js,3) &zt
  @ prow() , pcol() say '�� ƽ��:'+str(round(gs/rs,1),4,1)+'  ��' &zt
  @ prow() , pcol() say ' ��ע��೤: '+str(bz,2)+'�˩�' &zt
  @ prow() , pcol() say rs picture '@Z 9999' &zt
  @ prow() , pcol() say '��' &zt
@ prow()+1 , 0 say t4 &zt
@ prow()+1 , 10 say '�����:'+alltrim(str(year(date()),4))+'��'+alltrim(str(month(date());
,2))+'��'+alltrim(str(day(date()),2))+'��' &zt1
@ prow() , pcol() say '                                ���Ա:&ZB111' &zt1
set printer to lpt1
endif
?
YN = MESSAGEBOX('��Ҫ��ӡ��֤��Ա��λ�ֲ����𣿣�',4+32,'��ʾ')
if yn=6
t1='�������ө��������ө������ө������ө��������ө������ө����ө�����'
t2='����ũ���λ���੦ ���� �� ���� ��������ʩ� �ڼ� ���೤��������'
t3='�ĩ����੤�������੤�����੤�����੤�������੤�����੤���੤����'
t4='�������۩��������۩������۩������۩��������۩������۩����۩�����'
use ��֤����
sum ���� to rs
sum val(gwdm)*���� to gs
go top
xh=1
      @ 0 , 0 say space(1)+'&MC111'+'��λ�ȼ�ͳ�Ʊ�������' font '����',20
      @ 2 , 0 say '                         ( ��֤��Ա��λ�ֲ� )' &zt1
      @ 3 , 0 say T1 &zt
      @ 4 , 0 say T2 &zt
      @ 5 , 0 say T3 &zt
do while  not eof()
  @ prow()+1 , 0 say '��' &zt
  @ prow() , pcol() say str(xh,4) &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwfl1 &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwjb1 &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwdc1 &zt
  @ prow() , pcol() say '��  ' &zt
  @ prow() , pcol() say gwgz &zt
  @ prow() , pcol() say '  ��' &zt
  @ prow() , pcol() say �ڼ�+'��' &zt
if !empty(bzf)
   @ prow() , pcol() say ' �� ��' &zt
else
   @ prow() , pcol() say '    ��' &zt
endif
  @ prow() , pcol() say ���� picture '@Z 9999' &zt
  @ prow() , pcol() say '��' &zt
  @ prow()+1 , 0 say t3 &zt
  skip
  xh=xh+1
enddo
  @ prow()+1 , 0 say '��' &zt
  @ prow() , pcol() say '�ϼ�' &zt
  @ prow() , pcol() say '��        ��      ��      ��        ��      ��    ��' &zt
  @ prow() , pcol() say rs picture '@Z 9999' &zt
  @ prow() , pcol() say '��' &zt
@ prow()+1 , 0 say t4 &zt
@ prow()+1 , 10 say '�����:'+alltrim(str(year(date()),4))+'��'+alltrim(str(month(date());
,2))+'��'+alltrim(str(day(date()),2))+'��' &zt1
@ prow() , pcol() say '                                ���Ա:&ZB111' &zt1
set printer to lpt1
endif
?
YN = MESSAGEBOX('��Ҫ��ӡ������Ա��λ�ֲ����𣿣�',4+32,'��ʾ')
if yn=6
t1='�������ө��������ө������ө������ө��������ө������ө��������ө��������ө�����'
t2='����ũ���λ���੦ ���� �� ���� ��������ʩ� �ڼ� ������ְ�񩦼���ְ�Ʃ�������'
t3='�ĩ����੤�������੤�����੤�����੤�������੤�����੤�������੤�������੤����'
t4='�������۩��������۩������۩������۩��������۩������۩��������۩��������۩�����'
use �������
sum ���� for '����'$gwfl1 to js
sum ���� for '����'$gwfl1 to glry
sum ���� to rs
sum val(gwdm)*���� to gs
go top
xh=1
      @ 0 , 0 say space(7)+'&MC111'+'��λ�ȼ�ͳ�Ʊ����ģ�' font '����',20
      @ 2 , 0 say '                                      ( ������Ա��λ�ֲ� )' &zt1
      @ 3 , 0 say T1 &zt
      @ 4 , 0 say T2 &zt
      @ 5 , 0 say T3 &zt
do while  not eof()
  @ prow()+1 , 0 say '��' &zt
  @ prow() , pcol() say str(xh,4) &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwfl1 &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwjb1 &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gwdc1 &zt
  @ prow() , pcol() say '��  ' &zt
  @ prow() , pcol() say gwgz &zt
  @ prow() , pcol() say '  ��' &zt
  @ prow() , pcol() say �ڼ� &zt
  @ prow() , pcol() say '��' &zt
if '��'$gwfl1
  @ prow() , pcol() say '        ��' &zt
  @ prow() , pcol() say �ڵ� &zt
  @ prow() , pcol() say '��' &zt
else
  @ prow() , pcol() say �ڵ� &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say '        ��' &zt
endif
  @ prow() , pcol() say ���� picture '@Z 9999' &zt
  @ prow() , pcol() say '��' &zt
  @ prow()+1 , 0 say t3 &zt
  skip
  xh=xh+1
enddo
  @ prow()+1 , 0 say '��' &zt
  @ prow() , pcol() say '�ϼ�' &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say '�����λ:'+str(glry,3)+'�� ������λ:'+str(js,3) &zt
  @ prow() , pcol() say '�� ƽ��:'+str(round(gs/rs,1),4,1)+'��' &zt
  @ prow() , pcol() say '  ��:' &zt
  @ prow() , pcol() say rs picture '@Z 9999' &zt
  @ prow() , pcol() say '��'+space(16)+'��' &zt
@ prow()+1 , 0 say t4 &zt
@ prow()+1 , 10 say '�����:'+alltrim(str(year(date()),4))+'��'+alltrim(str(month(date());
,2))+'��'+alltrim(str(day(date()),2))+'��' &zt1
@ prow() , pcol() say '                                ���Ա:&ZB111' &zt1
set printer to lpt1
endif
?
YN = MESSAGEBOX('��Ҫ��ӡ����ͳ�Ʊ��𣿣�',4+32,'��ʾ')
if yn=6
t1='�������ө��������ө������������ө����������������ө������ө���������'
t2='����ũ� ��  �� ��  ��    ��  �� ��          λ �� ���� ����λ�ȼ���'
t3='�ĩ����੤�������੤�����������੤���������������੤�����੤��������'
t4='�������۩��������۩������������۩����������������۩������۩���������'
use ����ͳ��
sum a to rs
inde on �ڵ� to ����ͳ��
go top
xh=1
M = 1
  do while  not eof()
      @ 0 , 0 say space(1)+'&MC111'+str(year(date()),4)+'��'+allt(str(month(date()),2))+'�·ݹ���ͳ�Ʊ�' font '����',20
      @ 2 , 45 say '��'+str(M,2)+'ҳ' &zt1
      @ 3 , 0 say T1 &zt
      @ 4 , 0 say T2 &zt
      @ 5 , 0 say T3 &zt
  do while !eof()
  @ prow()+1 , 0 say '��' &zt
  @ prow() , pcol() say str(xh,4) &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gz &zt
  @ prow() , pcol() say ' ��' &zt
  @ prow() , pcol() say gz1 &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say gw &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say a picture '@Z 999999' &zt
  @ prow() , pcol() say '��' &zt
  if gwfl1='������λ'
     @ prow() , pcol() say space(1)+gwjb1+' ��' &zt
  else
     @ prow() , pcol() say �ڵ�+'��' &zt
  endi
  skip
  xh=xh+1
  if mod(xh,26)=0
     @ prow()+1 , 0 say t4 &zt
     m=m+1
     exit
  else
     @ prow()+1 , 0 say t3 &zt
  endif
  endd
enddo
  @ prow()+1 , 0 say '��' &zt
  @ prow() , pcol() say '�ϼ�' &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say space(8) &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say space(12) &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say space(16) &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say str(rs,6) &zt
  @ prow() , pcol() say '��' &zt
  @ prow() , pcol() say space(8)+'��' &zt
@ prow()+1 , 0 say t4 &zt
?
@ prow()+1 , 0 say '�����:'+alltrim(str(year(date()),4))+'��'+alltrim(str(month(date());
,2))+'��'+alltrim(str(day(date()),2))+'��' &zt1
@ prow() , pcol() say '          ���Ա:&ZB111' &zt1
?''
set printer to lpt1
endif
use
??chr(13)+chr(10)
set device to screen
set printer off
set console off
set confirm off
set talk off
retu
