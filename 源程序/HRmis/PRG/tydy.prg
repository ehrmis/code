***************************
* .\TYDY.PRG
***************************
close all
_SCREEN.PICTURE = ''
XM = '        '
�������� = space(5)
�Ʒ����� = space(4)
store 0 to JN���� , GW���� , GL���� , JT���� , SB���� , XL���� , С�� , ���� , ʵ��
store 1 to �������� , JLH , YX1
YF1=month(date())
ND1=year(date())
do srrq.spx
NY = ND+YF
use tydy excl
zap
append from ryzk
repl shf with 35 for shf=0
close all
select 3
use zr1&ly1
select 2
use ryzk
index on RYBH to ry
select 1
use tydy
set relation to RYBH into 2
repl shf with 0 for b.tybl='100'
select 2
go top
do while .t.      
      do xmsr.spr
      locate for upper(XM)=RYXM or upper(XM)=RYBH
      if eof()
        do dhkjg.spx
    M = readkey()
    if M=12 or M=268
      return
    endif
      else
      xm=allt(ryxm)
      clear
      @ 15 , 38 say '��ȷ��'+xm+'��������:' font '����',20 get CSRQ  font '����',20 picture '99999999'
      read
      @ 18 , 38 say '��ȷ��'+xm+'�μӹ�������:' font '����',20 get CJGZRQ  font '����',20 picture '9999999999' 
      read
      replace NL with val(ND)-val(left(CSRQ,4))+round((val(YF)-val(substr(CSRQ,5,2)))/12,2)
      JLH = recno()
        exit
      endif
enddo    
  select 3
  go top
  locate for RYXM=left(XM,8) or RYBH=left(upper(XM),5)
  ����ƽ��=pj
  select 1 
  go JLH
  T1 = '  ��λ����:&MC111'+'                                ��λ:Ԫ(������)' 
  T2 = '���������ө��������ө��ө��ө����ө����������ө������ө����������ө����ө�����������'
  T3 = '��      ��        ����  �����©�          �������ک�          �����©�;
          ��'
  T4 = '�ĩ������ة��������ة��ةЩة����ة����������੤�����ةЩ��������ة����ة�����������'
  T5 = '�ĩ��Щ������������������ة����������������ԩةЩ������ة���������������������������'
  T6 = '����   Ӧ      ��      ��      ��       ����   ��      ��   ;
   ��      ��     ��'
  T7 = '��  ���������������Щ������Щ����Щ���������  �������������Щ������Щ����Щ���������'
  T8 = '���ũ�  ��      Ŀ��ԭ�������� ��  �� ���ũ� ��    Ŀ ��ԭ��������;
 ��  �� ��'
  T9 = '�ĩ��੤�����������੤�����੤���੤�������橤�੤���������੤�����੤���੤��������'
  T10 = '�ĩ��੤�����������੤�����੤���੤�������橤�੤���������੤�����੤���੤��������'
  T11 = '�ĩ��੤�����������ة������ة����੤�������橤�੤���������ة������ة����੤��������'
  �������� = val(ND)-val(left(b.CJGZRQ,4))+round((val(YF)-val(substr(b.CJGZRQ,5,2)))/12,2)
  ���� = val(ND)-val(left(b.CJGZRQ,4))+1
  ҽ�Ʒ� = 2002-val(left(b.CJGZRQ,4))+1+20
  if ��������>=25
     ���乤�� = round(����*2.4,1)
  else
     ���乤�� = round(����*2,1)
  endif
  clear
  @ 15 , 38 say '��ȷ��'+xm+'��ǰ��������:' font '����',20 get �������� font '����',20 picture;
 '99.99' 
  @ row() , col() say '��' font '����',20
  read
  do case
  case ��������<25 and b.XB='0'
    �Ʒ����� = ' 65��'
  case ��������<30 and b.XB='1'
    �Ʒ����� = ' 65��'
  case ��������>=25 and ��������<30 and b.XB='0'
    �Ʒ����� = ' 70��'
  case ��������>=30 and ��������<35 and b.XB='1'
    �Ʒ����� = ' 70��'
  case ��������>=30 and b.XB='0'
    �Ʒ����� = ' 75��'
  case ��������>=35 and b.XB='1'
    �Ʒ����� = ' 75��'
  endcase
  if (b.NL>=55 and b.XB='1') or (b.NL>=45 and b.XB='0')
    if �Ʒ�����<' 70��'
      �Ʒ����� = ' 70��'
    endif
  endif
  if b.tybl='100'
     �Ʒ����� = '100��'
  endif
  @ 18 , 38 say '��ȷ��'+xm+'�������ʼƷ�����(%):' font '����',20 get �Ʒ����� font '����',20
  read
  JN���� = round(val(left(�Ʒ�����,3))/100*JNGZ,1)
  GW���� = round(val(left(�Ʒ�����,3))/100*GWGZ,1)
  GL���� = round(val(left(�Ʒ�����,3))/100*���乤��,1)
  JT���� = round(val(left(�Ʒ�����,3))/100*JTBT,1)
  XL���� = round(val(left(�Ʒ�����,3))/100*XLF,1)
  SB���� = round(val(left(�Ʒ�����,3))/100*SBF,1)
  С�� = JN����+GW����+GL����+JT����+XL����+SB����+ҽ�Ʒ�+HT+MT+SDBT+SHF
  ���� = b.FZSD+b.YLBX+b.YBX+b.SYBX+b.gjj
  ʵ�� = round(С��-����,1)
  clear
  @ 20 , 20 say '��������:Ӧ�� 'font '����',20 
  @ 20 , col() say alltrim(str(С��,8,1)) font '����',20 
  @ 20 , col() say 'Ԫ �� ���� ' font '����',20 
  @ 20 , col() say alltrim(str(����,8,1)) font '����',20 
  @ 20 , col() say 'Ԫ��= ʵ�� ' font '����',20 
  @ 20 , col() say alltrim(str(ʵ��,8,1)) font '����',20 
  @ 20 , col() say 'Ԫ ' font '����',20 
  = inkey(30,'M')  
  set device to printer
  set printer on
  go JLH  
    @ 3 , 5 say '�������������������ι�˾ְ����������Ԥ���' font '����',20
    set print font "����",12
    @ prow()+2 , 0 say T1
    @ prow()+0.8 , 0 say T2
    @ prow()+0.8 , 0 say '����  ����'+b.RYXM+'���ԩ�'+b.XB1+'��������'+left(b.CSRQ;
,4)+'��'+str(val(substr(b.CSRQ,5,2)),2)+'��'
    @ prow() , pcol() say '���μӹ���'+left(b.CJGZRQ,4)+'��'+str(val(substr(b.CJGZRQ;
,5,2)),2)+'��'
    @ prow() , pcol() say '��������'+ND+'��'+str(val(YF),2)+'�©�'
    @ prow()+0.8 , 0 say T3
    @ prow()+0.8 , 0 say T4
    @ prow()+0.8 , 0 say '��ԭ���¹���(ְ��)����  ��'+Gw+'  ���������䩦'+left(str(��������;
,5,2),2)+'����'+str(val(right(str(��������,5,2),2))/100*12,2)+'���� ;
             ��'
    @ prow()+0.8 , 0 say T5
    @ prow()+0.8 , 0 say T6
    @ prow()+0.8 , 0 say T7
    @ prow()+0.8 , 0 say T8
    @ prow()+0.8 , 0 say T9
    @ prow()+0.8 , 0 say '��1 �����ܹ���    ��'
    @ prow() , pcol() say JNGZ picture '9999.9'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say left(�Ʒ�����,3)+'%'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say JN���� picture '99999.99'
    @ prow() , pcol() say '��12������ˮ��ѩ�      ��    ��'
    @ prow() , pcol() say b.FZSD picture '99999.99'
    @ prow() , pcol() say '��'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '��2 ����λ����    ��'
    @ prow() , pcol() say GWGZ picture '9999.9'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say left(�Ʒ�����,3)+'%'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say GW���� picture '99999.99'
    @ prow() , pcol() say '��13��ס��������      ��    ��'
    @ prow() , pcol() say b.gjj picture '99999.99'
    @ prow() , pcol() say '��'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '��3 �����乤��    ��'
    @ prow() , pcol() say ���乤�� picture '999.99'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say left(�Ʒ�����,3)+'%'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say GL���� picture '99999.99'
    @ prow() , pcol() say '��14�����ϱ��ս�      ��    ��'
    @ prow() , pcol() say b.YLBX picture '99999.99'
    @ prow() , pcol() say '��'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '��4 ��ˮ�粹��    ��'
    @ prow() , pcol() say SDBT picture '999.99'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say '100%'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say SDBT picture '99999.99'
    @ prow() , pcol() say '��15��ҽ�Ʊ��ս�      ��    ��'
    @ prow() , pcol() say b.YBX picture '99999.99'
    @ prow() , pcol() say '��'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '��5 ����ͨ����    ��'
    @ prow() , pcol() say JTBT picture '999.99'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say left(�Ʒ�����,3)+'%'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say JT���� picture '99999.99'
    @ prow() , pcol() say '��16��ʧҵ���ս�      ��    ��'
    @ prow() , pcol() say SYBX picture '99999.99'
    @ prow() , pcol() say '��'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '��6 ���鱨��      ��'
    @ prow() , pcol() say SBF picture '999.99'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say left(�Ʒ�����,3)+'%'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say SB���� picture '99999.99'
    @ prow() , pcol() say '��17��          ��      ��    ��        ��'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '��7 ��ϴ���      ��'
    @ prow() , pcol() say XLF picture '999.99'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say left(�Ʒ�����,3)+'%'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say XL���� picture '99999.99'
    @ prow() , pcol() say '��18��          ��      ��    ��        ��'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '��8 ��ҽҩ��      ��'
    @ prow() , pcol() say ҽ�Ʒ� picture '999.99'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say '100%'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say ҽ�Ʒ� picture '99999.99'
    @ prow() , pcol() say '��19��          ��      ��    ��        ��'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '��9 ��ú��        ��'
    @ prow() , pcol() say MT picture '999.99'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say '100%'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say MT picture '99999.99'
    @ prow() , pcol() say '��20��          ��      ��    ��        ��'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '��10������������ѩ�'
    @ prow() , pcol() say HT+SHF picture '999.99'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say '100%'
    @ prow() , pcol() say '��'
    @ prow() , pcol() say HT+SHF picture '99999.99'
    @ prow() , pcol() say '��21��          ��      ��    ��        ��'
    @ prow()+0.8 , 0 say T11
    @ prow()+0.8 , 0 say '��11��С��(1+2+3+4+5+6+7+8+9+10)��'
    @ prow() , pcol() say С�� picture '99999.99'
    @ prow() , pcol() say '��22��     �� ʵ �� �� ��     ��'
    @ prow() , pcol() say ʵ�� picture '99999.99'
    @ prow() , pcol() say '��'
    @ prow()+0.8 , 0 say '�ĩ��੤�������������������������ة��������ܩ��ة������������������������ة���������'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '������                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��ע��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '��  ��                                          ;
                                  ��'
    @ prow()+0.8 , 0 say '�����۩�����������������������������������������������������������������������������'   
set device to screen
?''
set print off
set print to
clear
 _SCREEN.WINDOWSTATE = 2
close all
use data\ccrq
go 8
repl ͼƬ with str(val(ͼƬ)+1,2)
if allt(ͼƬ)='10'
   repl ͼƬ with '0'
endi   
pict_fd='fd'+allt(ͼƬ)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
_SCREEN.MAXBUTTON = .F.
_SCREEN.MOVABLE = .F.
_SCREEN.CONTROLBOX = .F.
_SCREEN.CLOSABLE = .F.
close all
return
