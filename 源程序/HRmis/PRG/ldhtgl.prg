***************************
*LDHTGL.PRG(Build20080415)
***************************
close all
set date to YMD
set date to long
CJBH = ''
DH = ''
DH1 = ''
DH1A = ''
DYCJ = ''
DYBM = ''
do FJDY
do DHK1
do DHK1A
if DH1A=2 and DH1=2
   do htdc
endif
close all
if DH1=2 and DH1A=1
  nn=MESSAGEBOX("����ӡ���ɹ�������Ԥ�����ӡ��...",0+64,"��Ϣ") 
  set device to printer
  set printer on  
else
  set print font "����",8
*********************��****************
  set device to printer
  set printer to �Ͷ���ͬ.TXT
  set printer on 
endif
use ldht excl
zap
append from ryzk
RS = reccount()
do case
case DH=2
  coun for cjdm='&dycj' to rs
  set filter to CJDM='&DYCJ'
case DH=3
  coun for bmbh='&dybm' to rs
  set filter to BMBH='&DYBM'
endcase
PG = 1
N = 1
go top
do while  not eof()
  set print font "����",16
  @ 3 , 0 say space(18)+'���ּ����Ͷ���ͬ������������'
  set print font "����",8
  @ prow()+3 , 3 say '���λ:'+alltrim(MC111)+space(33)+'��'+str(PG;
,2)+'ҳ'+space(28)+'����:'+str(year(date()),4)+'��'+str(month(date()),2)+'��'+str(day(date());
,2)+'��'
  T1 = '�������ө��������ө��ө��������������ө����ө������������ө��������������ө��������ө������������������ө�������'
  T2 = '��    ��        ����              �����ީ�            ��   ��ͬʱ��   ��        �� ������Ƿ����ʱ�� ��      ��'
  T3 = '�ĩ����੤�������੤�੤�������������੤���੤�����������੤�������������੤�������੤�����������������੤������'
  T4 = '�������۩��������۩��۩��������������۩����۩������������۩��������������۩��������۩������������������۩�������'
  @ prow()+1 , 0 say T1
  @ prow()+0.8 , 0 say '����ũ� ��  �� ���ԩ�   ��������   �������� ��λ(����) ��   ǩ���Ͷ�   ����ͬ���ީ�'+str(year(date()),4)+'��'+str(month(date());
,2)+'��ĩ����ͬ����  ע��'
  @ prow()+0.8 , 0 say T2
  @ prow()+0.8 , 0 say T3
  if PG>1
    N = (PG-1)*30+1
    skip
  endif
  do while  not eof()
    @ prow()+0.8 , 0 say '��'+str(N,4)+'��'+RYXM+'��'+XB1+'��'
    do case
    case len(allt(dtoc(csrq)))=12
         CS=allt(dtoc(csrq))+space(2)
    case len(allt(dtoc(csrq)))=13
         CS=allt(dtoc(csrq))+space(1)
    case len(allt(dtoc(csrq)))=14
         CS=allt(dtoc(csrq))
    endcase             
    @ prow() , pcol() say CS+'��'+str(gl,2)+'�ꩦ'
    if  not empty(GZ1)
      @ prow() , pcol() say left(GZ1,12)+'��'
    else
      @ prow() , pcol() say left(GW,12)+'��'
    endif
IF year(XHTrQ)>0
***********��ǩ����ͬ��Ա**************
       ��ͬ���� = val(xQHTQ)+(year(xHTRQ)-year(date()))+round((month(xHTRQ)-month(date()))/12,2)
    do case
    case len(allt(dtoc(xhtrq)))=12
         CS=allt(dtoc(xhtrq))+space(2)
    case len(allt(dtoc(xhtrq)))=13
         CS=allt(dtoc(xhtrq))+space(1)
    case len(allt(dtoc(xhtrq)))=14
         CS=allt(dtoc(xhtrq))
    endcase    
      @ prow() , pcol() say CS+'��'+XQHTQ1+'��'            
     if !'�޹̶���'$XQHTQ1 and !empty(xqhtq1)
      do case
      case val(left(str(��ͬ����,5,2),2))>0 and val(right(str(��ͬ����,5;
,2),2))/100*12>0
        replace HTNX with left(str(��ͬ����,5,2),2)+'����'+str(val(right(str(��ͬ����;
,5,2),2))/100*12,2)+'����'
        @ prow() , pcol() say space(3)+HTNX+space(3)
      case val(left(str(��ͬ����,5,2),2))>0 and val(right(str(��ͬ����,5;
,2),2))/100*12=0
        replace HTNX with left(str(��ͬ����,5,2),2)+'��'
        @ prow() , pcol() say space(3)+HTNX+space(3)
      case val(left(str(��ͬ����,5,2),2))=0 and val(right(str(��ͬ����,5;
,2),2))/100*12>0
        replace HTNX with str(val(right(str(��ͬ����,5,2),2))/100*12,2)+'����'
        @ prow() , pcol() say space(3)+HTNX+space(3)
         case val(left(str(��ͬ����,5,2),2))=0 and val(right(str(��ͬ����,5;
,2),2))/100*12=0
        replace HTNX with '0����'
         @ prow() , pcol() say space(3)+HTNX+space(3)
      endcase      
else
@ prow() , pcol() say space(18)
endif
@ prow() , pcol() say '��      ��'
ELSE    
***********����ǩ����ͬ��Ա**************
       if !'�޹̶���'$sQHTQ1 and !empty(sqhtq1)
       ��ͬ���� = val(SQHTQ)+(year(SHTRQ)-year(date()))+round((month(SHTRQ)-month(date()))/12,2)
        do case
    case len(allt(dtoc(shtrq)))=12
         CS=allt(dtoc(shtrq))+space(2)
    case len(allt(dtoc(shtrq)))=13
         CS=allt(dtoc(shtrq))+space(1)
    case len(allt(dtoc(shtrq)))=14
         CS=allt(dtoc(shtrq))
    endcase 
       @ prow() , pcol() say CS+'��'+SQHTQ1+'��'
      else
        ��ͬ���� = val(SQHTQ)+(year(SHTRQ)-year(date()))+round((month(SHTRQ)-month(date()))/12,2)
        do case
    case len(allt(dtoc(shtrq)))=12
         CS=allt(dtoc(shtrq))+space(2)
    case len(allt(dtoc(shtrq)))=13
         CS=allt(dtoc(shtrq))+space(1)
    case len(allt(dtoc(shtrq)))=14
         CS=allt(dtoc(shtrq))
    endcase    
        @ prow() , pcol() say CS+'��'+SQHTQ1+'��'
      endif
       if !'�޹̶���'$sQHTQ1 and !empty(sqhtq1)
      do case
      case val(left(str(��ͬ����,5,2),2))>0 and val(right(str(��ͬ����,5;
,2),2))/100*12>0
        replace HTNX with left(str(��ͬ����,5,2),2)+'����'+str(val(right(str(��ͬ����;
,5,2),2))/100*12,2)+'����'
        @ prow() , pcol() say space(3)+HTNX+space(3)
      case val(left(str(��ͬ����,5,2),2))>0 and val(right(str(��ͬ����,5;
,2),2))/100*12=0
        replace HTNX with left(str(��ͬ����,5,2),2)+'��'
        @ prow() , pcol() say space(3)+HTNX+space(3)
      case val(left(str(��ͬ����,5,2),2))=0 and val(right(str(��ͬ����,5;
,2),2))/100*12>0
        replace HTNX with str(val(right(str(��ͬ����,5,2),2))/100*12,2)+'����'
        @ prow() , pcol() say space(3)+HTNX+space(3)
        case val(left(str(��ͬ����,5,2),2))=0 and val(right(str(��ͬ����,5;
,2),2))/100*12=0
        replace HTNX with '0����'
        @ prow() , pcol() say space(3)+HTNX+space(3)
      endcase
         @ prow() , pcol() say '�� ��ǩ ��'
      else
         @ prow() , pcol() say space(18)+'��      ��'
      endif   
ENDIF
    if mod(N,30)=0
      @ prow()+0.8 , 0 say T4
      exit
    else
      @ prow()+0.8 , 0 say T3
    endif     
    skip
    N = N+1
   enddo
  PG = PG+1
enddo
@ prow()+0.8 , 0 say T4
set filter to
close all
@ prow()+1 , 10 say '����������(��ί�д�����):'+ZG111+space(42)+'�Ʊ���:'+ZB111
@ prow()+1 , 0 say ' '
set device to screen
set print off
set print to
close all
use ldht
repl htdqsj with allt(htnx) for !'��'$htnx
replace SQHTQ1 with XQHTQ1 for  not empty(XQHTQ1)
index on CJDM+SQHTQ1+HTNX to ldht
copy to d:\�Ͷ���ͬ type xl5 for !empty(htdqsj)
=messagebox("�����ͬ����Ա�������������D��\�Ͷ���ͬ���͡�D��\�Ͷ���ͬ���顱���ӱ��У���",48,"��ϲ")
total on CJDM+SQHTQ1+HTNX to ��ͬ����
use ��ͬ����
copy to d:\�Ͷ���ͬ���� type xl5
index on CJDM+SQHTQ1 to ldht
total on CJDM+SQHTQ1 to ��ͬ1 
index on SQHTQ1 to ldht
total on SQHTQ1 to ��ͬ2 
use ��ͬ����
copy to htdc
close all
if DH1=1 and DH1A=1
  modi file �Ͷ���ͬ.TXT
endif
if DH1A=2 and DH1=1
  use ��ͬ����
  sum to RS1 A
  index on SQHTQ1 to ��ͬ����
  go top
  do case
  case DH=3
    browse field BMMC :R :H = '��������' , SQHTQ1 :R :H = 'ǩ����ͬ����';
 , HTNX :R :H = '������δ���к�ͬʱ��' , A :R : 4 :H = '����';
 title '��ͬ�����('+alltrim(CJMC)+':'+str(RS1,4)+'��)'
  case DH=2
    browse field CJMC :R :H = '��������' , SQHTQ1 :R :H = 'ǩ����ͬ����';
 , HTNX :R :H = '������δ���к�ͬʱ��' , A :R : 4 :H = '����';
 title '��ͬ�����('+alltrim(CJMC)+':'+str(RS1,4)+'��)'
  case DH=1
    browse field CJMC :R :H = '��������' , SQHTQ1 :R :H = 'ǩ����ͬ����';
 , HTNX :R :H = '������δ���к�ͬʱ��' , A :R : 4 :H = '����';
 title '��ͬ�����('+str(RS1,4)+'��)'
  use ��ͬ1
    browse field CJMC :R :H = '��������' , SQHTQ1 :R :H = 'ǩ����ͬ����';
 , A :R : 4 :H = '����';
 title '��ͬ�����('+str(RS1,4)+'��)'
  use ��ͬ2
    browse field SQHTQ1 :R :H = 'ǩ����ͬ����';
 , A :R : 4 :H = '����';
 title '��ͬ�����('+str(RS1,4)+'��)'
 endcase
endif
close all
set device to screen
if DH1A=2 and DH1=2
   set printer to lpt1
   do htdc   
endif
set printer off
set printer to
close all
clear
return

proc htdc
set talk off
set safety off
tt1='�������ө������������ө������������ө��������������������ө�������'
tt2='����ũ�  ��������  ��ǩ����ͬ���ީ�������δ���к�ͬʱ�䩦 ���� ��'
tt3='�ĩ����੤�����������੤�����������੤�������������������੤������'
tt4='�������۩������������۩������������۩��������������������۩�������'
close all
use htdc
sum a to rs1
go top
M = 1
xh1 = 1
set device to printer
set printer on
set print font "����",12
*********************��****************
 do while !eof()
    @ 3 , 0 say space(26)+alltrim(MC111)+'�Ͷ���ͬ�����'
    @ prow()+1 , 20 say '��'+str(M;
,2)+'ҳ'+space(18)+'����:'+str(year(date()),4)+'��'+str(month(date());
,2)+'��'+str(day(date()),2)+'��'
      @  prow()+1 , 10 say Tt1
      @  prow()+0.8 , 10 say Tt2
      @  prow()+0.8 , 10 say Tt3
  do while !eof()
  @ prow()+0.8 , 10 say '��'
  @ prow() , pcol() say str(xh1,4)
  @ prow() , pcol() say '��'
  @ prow() , pcol() say cjmc
  @ prow() , pcol() say '��  '
  @ prow() , pcol() say sqhtq1
  @ prow() , pcol() say '  ��'
  @ prow() , pcol() say space(4)+htnx+space(4)
  @ prow() , pcol() say '��'
  @ prow() , pcol() say a picture '@Z 999999'
  @ prow() , pcol() say '��'
  skip
  xh1=xh1+1
  if mod(xh1,25)=0
     @ prow()+1 , 10 say tt4
     m=m+1
     exit
  else
     @ prow()+0.8 , 10 say tt3
  endif
  endd
enddo
  @ prow()+0.8 , 10 say '��'
  @ prow() , pcol() say '�ϼ�'
  @ prow() , pcol() say '��'
  @ prow() , pcol() say space(12)
  @ prow() , pcol() say '��'
  @ prow() , pcol() say space(12)
  @ prow() , pcol() say '��'
  @ prow() , pcol() say space(20)
  @ prow() , pcol() say '��'
  @ prow() , pcol() say str(rs1,6)
  @ prow() , pcol() say '��'
@ prow()+0.8 , 10 say tt4
?''
set device to screen
set print off
set print to
close all
clear
retu
