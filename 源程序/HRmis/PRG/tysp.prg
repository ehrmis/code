***************************
* .\TYSP.PRG
***************************
set talk off
close all
_SCREEN.PICTURE = ''
XM = '        '
�������� = space(5)
�绰 = space(12)
store 1 to �������� , JLH , YX1
YF1=month(date())
ND1=year(date())
do srrq.spx
NY = ND+YF
clear
use ryzk
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
  go JLH
T1 = '   ��λ��&MC111'+space(25)+'��˾������ţ�' 
T2='�����������ө������ө����ө������������ө��������ө����������ө��������ө�����������'   
T3='�ĩ��������੤�����੤���੤�����������੤�������੤���������੤�������੤����������'    
T4='�ĩ��������੤�����ة����ة������������੤�������੤���������੤�������ة�����������'    
T5='�ĩ��������੤�������������������������ة��������ة����������ة���������������������'
T6='�ĩ��������੤����������������������������������������������������������������������'
T7='�����������۩�����������������������������������������������������������������������'
  SR1 = ND
  SR2 = str(val(YF),2)
  SR4 = str(val(substr(CSRQ,5,2)),2)
  do case
  case XB='0' and empty(ZCJT)
    SR3 = val(left(CSRQ,4))+50
  case XB='0' and  not empty(ZCJT)
    SR3 = val(left(CSRQ,4))+55
  case XB='1'
    SR3 = val(left(CSRQ,4))+60
  endcase
  SR3 = str(SR3,4)
  clear
  @ 15 , 38 say '����ʱ��:' font '����',20 get SR1 font '����',20 picture '9999'
  @ row() , col()+2 say '��' font '����',20 get SR2 font '����',20 picture '99'
  @ row() , col()+2 say '������������' font '����',20 get SR3 font '����',20 picture '9999'
  @ row() , col()+2 say '��' font '����',20 get SR4 font '����',20 picture '99'
  @ row() , col()+2 say '��' font '����',20
  read  
  @ 18 , 38 say '��ȷ��'+xm+'��ͥ��ַ:' font '����',20 get ZZ font '����',20
  read
   �������� = val(ND)-val(left(CJGZRQ,4))+round((val(YF)-val(substr(CJGZRQ;
,5,2)))/12,2)
***************
  ���� = val(ND)-val(left(CJGZRQ,4))+1
  ҽ�Ʒ� = 2002-val(left(CJGZRQ,4))+1+20
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
  case ��������<25 and XB='0'
    �Ʒ����� = ' 65��'
  case ��������<30 and XB='1'
    �Ʒ����� = ' 65��'
  case ��������>=25 and ��������<30 and XB='0'
    �Ʒ����� = ' 70��'
  case ��������>=30 and ��������<35 and XB='1'
    �Ʒ����� = ' 70��'
  case ��������>=30 and XB='0'
    �Ʒ����� = ' 75��'
  case ��������>=35 and XB='1'
    �Ʒ����� = ' 75��'
  endcase
  if (NL>=55 and XB='1') or (NL>=45 and XB='0')
    if �Ʒ�����<' 70��'
      �Ʒ����� = ' 70��'
    endif
  endif
  if tybl='100'
     �Ʒ����� = '100  '
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
  ���� = FZSD+YLBX+YBX+SYBX+gjj
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
***************
set device to printer
set printer on
  go JLH  
    @ 3 , 12 say '�������������������ι�˾ְ������������' font '����',20
    set print font "����",12
    @ prow()+3 , 0 say T1 
    @ prow()+2 , 0 say T2
    @ prow()+0.8 , 0 say '����    ����'+left(RYXM,6)+'���Ա�     '+XB1+'     ���������©�'+left(CSRQ;
,4)+'��'+str(val(substr(CSRQ,5,2)),2)+'��'
    @ prow() , pcol() say '������ʱ�䩦'+left(CJGZRQ,4)+'��'+str(val(substr(CJGZRQ;
,5,2)),2)+'�©�'
    @ prow()+0.8 , 0 say T3
    @ prow()+0.8 , 0 say '���Ļ��̶ȩ� '+WHCD2+' �����֩�'+Gz1
    @ prow() , pcol() say '���������䩦'+left(str(��������,5,2),2)+'��'+str(val(right(str(��������;
,5,2),2))/100*12,2)+'���©����ܱ�����  '+left(�Ʒ�����,3)+'%    ��'
    @ prow()+0.8 , 0 say T4
    @ prow()+0.8 , 0 say '��        ��                          ��        �� Ӧ����� ��'+str(С��,10,2)+'Ԫ        ��'
    @ prow()+0.8 , 0 say '����ͥסַ��'+ZZ+space(6)+'���������ʩ������������੤��������������������'
    @ prow()+0.8 , 0 say '��        ��                          ��        �� ʵ����� ��'+str(ʵ��,10,2)+'Ԫ        ��'
    @ prow()+0.8 , 0 say T5
    @ prow()+0.8 , 0 say '��        ��                                                                      ��'
    @ prow()+0.8 , 0 say '��        ��                                                                      ��'
    @ prow()+0.8 , 0 say '������ԭ��                                                                      ��'
    @ prow()+0.8 , 0 say '��        ��                                                                      ��'
    @ prow()+0.8 , 0 say '��        ��                                          ��     ��     ��            ��'
    @ prow()+0.8 , 0 say T6
    @ prow()+0.8 , 0 say '������ʱ�䩦'+space(12)+SR1+' ��'+SR2+'����������������'+'                                 ��'
    @ prow()+0.8 , 0 say T6
@ prow()+0.8 , 0 say '��   ��   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ��   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ��   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   λ   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ��   ��                                          ��     ��     ��            ��'
@ prow()+0.8 , 0 say '��   ��   ��                                                                      ��'
@ prow()+0.8 , 0 say T6
@ prow()+0.8 , 0 say '��   ��   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ��   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ��   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ˾   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ��   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ��   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ��   ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ��   ��                                          ��     ��     ��            ��'
@ prow()+0.8 , 0 say T6
@ prow()+0.8 , 0 say '��        ��                                                                      ��'
@ prow()+0.8 , 0 say '��        ��                                                                      ��'
@ prow()+0.8 , 0 say '��   ��   ��1������˾������š��ɼ��Ź�˾�Ͷ����´���д��                         ��'
@ prow()+0.8 , 0 say '��        ��2�������ɼ��Ź�˾�Ͷ����´�������                                     ��'
@ prow()+0.8 , 0 say '��   ע   ��3������һʽ���ݣ���λ�浵һ�ݣ��������ش浵һ�ݡ�                     ��'
@ prow()+0.8 , 0 say '��        ��                                                                      ��'
@ prow()+0.8 , 0 say '��        ��                                                                      ��'
@ prow()+0.8 , 0 say T7
@ prow()+1 ,0 say '����  ְ������ǩ����            ��λ�����ˣ�'+'&ZB111'+'    ��˾���������ˣ�'      
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
