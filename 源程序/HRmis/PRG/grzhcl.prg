*******************************
* .\grzhcl.PRG(Build20060321)
*******************************
close all
set talk off
clear
_SCREEN.PICTURE = ''
ƽ��2=0
use lnk
copy to temp stru
close all
use fox\grzh excl
if file("d:\grzh.dbf")
   zap
   appe from d:\grzh
endif 
go top
loca for aae001=1995
if !eof()
yn = MESSAGEBOX("�Ƿ��롰�����籣ϵͳ���нɷѻ�����������",4+32,"��ʾ")
IF yn = 6
   �˿� = '��Ƭ'
ENDIF   
endif 
*************(1)��ʼ���ɷѻ���������*****************
IF �˿� = '��Ƭ'
******�Զ�����******
use fox\grzh
sort on aae001 to temp1 for allt(grbh)=bh
sele 2
use temp1
go top
sele 1
use temp 
��� = 1995
for I = 1 to val(LY)-1994
    if allt(whtj)='ryzk'
       appe from ryzk for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
    else
       appe from ryzk1 for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
    endif   
    replace ��� with str(���,4)
   ��� = ���+1
endfor
replace �½ɷѻ��� with 0 for val(���)<val(LY)
go top
do while !eof()    
    sele 2
    sum aic079 to ys for aae001=val(a.���)
    go top
    loca for aae001=val(a.���)    
    if !eof()
       replace a.�ɷ����� with ys,a.�½ɷѻ��� with jfjs
    endif
    sele 1
    skip
enddo
ELSE
******�˹�����******
use temp 
��� = 1995
for I = 1 to val(LY)-1994
   if whtj='ryzk'
       appe from ryzk for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
    else
       appe from ryzk1 for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
    endif 
    replace ��� with str(���,4)
    ��� = ���+1
endfor
replace �ɷ����� with 6 for ���='1995'
replace �ɷ����� with 12 for ���<>'1995'
replace �ɷ����� with �ɷ�����-3 for ���='2003'and �ɷ�����>3
replace �½ɷѻ��� with 0 for val(���)<val(LY)
go bott
replace �ɷ����� with month(date()) , ˵�� with ���+'��ʵ���·���'
ENDIF
******��ˡ��ɷ����� , �½ɷѻ�����******
replace ˵�� with str(val(���)-1,4)+'����ƽ������' all
close all
use temp excl
go top
clear
reca all
delete for '-1'$˵��
pack
@ 36 , 90 say '������ʾ' font '����',14
@ 38 , 70 say '(1)����ĩ��ɷ�����,2001��ֱ����2000����ƽ������' font '����',12
@ 40 , 70 say '(2)��ӡ������ҵͳ��ת�ƻ���˵������1998������' font '����',12
@ 42 , 70 say '(3)��ӡ�������ʻ��Ǽǿ����ӱ��˽ɷ���ʼ������' font '����',12
 browse for left(XM,8)=left(RYXM,8) or left(XM,5)=left(RYBH,5) field;
 RYBH :R : 8 :H = '��  ��' , RYXM :R : 8 :H = '��  ��' , ��� :R ;
,�ɷ����� , �½ɷѻ��� , ˵��:R titl '���ϱ����ֲ�'
go top
loca for ���='2002' and (allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH))
if !eof()
   ƽ��2=�½ɷѻ���
endif
delete for �½ɷѻ���=0
pack
use ylbxk
copy to ylgrzh stru 
use ylgrzh 
appe from temp
*************(2)�Զ����˹����������ʻ�(��������)*****************
IF �˿� = '��Ƭ'
******�Զ�����******
close all
use ylgrzh 
go top
����=val(���)
close all
select 2
use ��������
loca for val(���)=����
sele 1
use ylgrzh  
go top
LL = b.�������ʡ�
LNLL = b.�������ʣ�
DW = b.��λ���ɣ�
GR = b.���˽��ɣ�
PJGZ = b.������ƽ��
��=val(���)
if ��<=1997
  SN = b.������ƽ��
else
  SN = 0
endif
 replace �ɷѻ��� with �½ɷѻ���*�ɷ����� , �µ�λ�ɷ� with round(�½ɷѻ���*DW/100,1) , ��λ�ɷ�;
 with �µ�λ�ɷ�*�ɷ����� , ��ƽ5 with round(SN*5/100,1)*�ɷ����� , �¸��˽ɷ�;
 with round(�½ɷѻ���*GR/100,1) , ���˽ɷ� with �¸��˽ɷ�*�ɷ����� , ���� with GR;
 ,��λ��Ϣ with round((��λ�ɷ�+��ƽ5)*(LL/1000)*�ɷ�����/2,1);
 , ������Ϣ with round(���˽ɷ�*(LL/1000)*�ɷ�����/2,1) , ����ϼ�;
 with ��λ�ɷ�+��ƽ5+��λ��Ϣ+���˽ɷ�+������Ϣ,���ۼ� with ����ϼ�
ENDIF
******�˹�����******
close all 
use ylgrzh excl
clear
@ 38 , 70 say '(1)��ӡ������ҵͳ��ת�ƻ���˵������������1998�굱������(������������)' font '����',12
@ 40 , 70 say '(2)��ӡ�������ʻ��Ǽǿ������뱾�˽ɷ���ʼ������' font '����',12
go top
browse part 12 field ���: r ,rybh:r:h='���',RYXM:r:h='����' ;
 , �ɷ����� : 8 :H = '�ɷ�����' , �ɷѻ��� : 8 :H = '�ɷѻ���';
 , ��λ�ɷ� : 12 :H = '���굥λ�ɷ�' , ��ƽ5 : 7 :H = '��ƽ5%';
 , ��λ��Ϣ : 12 :H = '���굥λ��Ϣ' , ���� :h ='���˽ɷѱ���%' , ���˽ɷ� : 12 :H = '������˽ɷ�';
 , ������Ϣ : 12 :H = '���������Ϣ' , ����ϼ� : 10 :H = '����ϼ�';
 , ���굥λ�� : 12 :H = '���굥λ�ɷ�' , ������Ϣ1 : 12 :H =;
 '���굥λ��Ϣ' , ������˽� : 12 :H = '������˽ɷ�' , ������Ϣ2;
 : 12 :H = '���������Ϣ' , ���ۼ� : 10 :H = '��    ��' titl  '���ϱ��ո����ʻ��Ǽǿ� �� [ Ctrl+T ] ɾ��' 
 
*************(3)���������ʻ�(��������)***************** 
��λ��=��λ�ɷ�
��ƽ=��ƽ5
��λϢ=��λ��Ϣ
���굥λ=���굥λ��
��Ϣ1=������Ϣ1
���˽�=���˽ɷ�
����Ϣ=������Ϣ
�������=������˽�
��Ϣ2=������Ϣ2
�ܼ�=���ۼ�
repl ���˱�Ϣ with ���˽ɷ�+������Ϣ+������˽�+������Ϣ2 , ���걾Ϣ with ������˽�+������Ϣ2 all
pack
go top
����=val(���)
loca for ���='1998'
if ����=1998
   �����ܱ�Ϣ=���˱�Ϣ
   �������걾Ϣ=���걾Ϣ-������Ϣ2
endif
close all
use ��������
copy to temp for val(���)>����
select 2
use temp
inde on ��� to temp
go top
sele 1
use ylgrzh
set relation to ��� into 2
go top
skip
do while !eof()
��=val(���)
if ��<=1997
  SN = b.������ƽ��
else
  SN = 0
endif
LL = b.�������ʡ�
LNLL = b.�������ʣ�
DW = b.��λ���ɣ�
GR = b.���˽��ɣ�
PJGZ = b.������ƽ��
 replace �ɷѻ��� with �½ɷѻ���*�ɷ����� , �µ�λ�ɷ� with round(�½ɷѻ���*DW/100,1) , ��λ�ɷ�;
 with �µ�λ�ɷ�*�ɷ����� , ��ƽ5 with round(SN*5/100,1)*�ɷ����� , �¸��˽ɷ�;
 with round(�½ɷѻ���*GR/100,1) , ���˽ɷ� with �¸��˽ɷ�*�ɷ����� , ���� with GR
if ���='2001'
    do case
    case �ɷ�����=1
      replace ��λ�ɷ� with round(ƽ��2*4/100,1) , ���˽ɷ� with round(ƽ��2*7/100,1) , �ɷѻ��� with ƽ��2
    case �ɷ�����=2
      replace ��λ�ɷ� with round(ƽ��2*4/100,1)*2 , ���˽ɷ� with round(ƽ��2*7/100,1)*2 , �ɷѻ��� with;
 ƽ��2*2
    case �ɷ�����=3
      replace ��λ�ɷ� with round(ƽ��2*4/100,1)*3 , ���˽ɷ� with round(ƽ��2*7/100,1)*3 , �ɷѻ��� with;
 ƽ��2*3
    case �ɷ�����>3
      replace ��λ�ɷ� with �µ�λ�ɷ�*(�ɷ�����-3)+round(ƽ��2*4/100,1)*3 , ���˽ɷ� with;
 �¸��˽ɷ�*(�ɷ�����-3)+round(ƽ��2*7/100,1)*3 , �ɷѻ��� with �½ɷѻ���*(�ɷ�����-3)+ƽ��2*3
    endcase
endif
  if ���='2002'
    do case
    case �ɷ�����=1
     replace ��λ�ɷ� with round(�½ɷѻ���*3/100,1) , ���˽ɷ� with round(�½ɷѻ���*8/100,1)
    case �ɷ�����=2
     replace ��λ�ɷ� with round(�½ɷѻ���*3/100,1)*2 , ���˽ɷ� with round(�½ɷѻ���*8/100,1)*2
    case �ɷ�����=3
     replace ��λ�ɷ� with round(�½ɷѻ���*3/100,1)*3 , ���˽ɷ� with round(�½ɷѻ���*8/100,1)*3
    case �ɷ�����>3
     replace ��λ�ɷ� with �µ�λ�ɷ�*(�ɷ�����-3)+round(�½ɷѻ���*3/100,1)*3 , ���˽ɷ� with;
 �¸��˽ɷ�*(�ɷ�����-3)+round(�½ɷѻ���*8/100,1)*3
    endcase
  endif
   replace ��λ��Ϣ with round((��λ�ɷ�+��ƽ5)*(LL/1000)*�ɷ�����/2,1);
 , ������Ϣ with round(���˽ɷ�*(LL/1000)*�ɷ�����/2,1) , ����ϼ�;
 with ��λ�ɷ�+��ƽ5+��λ��Ϣ+���˽ɷ�+������Ϣ , ���굥λ�� with ��λ��+��ƽ+��λϢ+���굥λ+��Ϣ1;
 , ������Ϣ1 with round(���굥λ��*(LNLL/100),1) , ������˽� with;
 ���˽�+����Ϣ+�������+��Ϣ2 , ������Ϣ2 with round(������˽�*(LNLL/100);
,1) , ���ۼ� with ����ϼ�+���굥λ��+������Ϣ1+������˽�+������Ϣ2;
 , ������� with ��λ�ɷ�+���˽ɷ� , �ۼ���Ϣ with ��λ��Ϣ+������Ϣ+������Ϣ1+������Ϣ2;
 , �����ۼ� with �ܼ� , ������ƽ with PJGZ , �������Ͻ� with round(������ƽ*0.2;
,1) , �˻����Ͻ� with round(���ۼ�/120,1) , ���ϴ����� with round(�������Ͻ�+�˻����Ͻ�;
,1) , ���˱�Ϣ with ���˽ɷ�+������Ϣ+������˽�+������Ϣ2 , ���걾Ϣ with ������˽�+������Ϣ2
��λ��=��λ�ɷ�
��ƽ=��ƽ5
��λϢ=��λ��Ϣ
���굥λ=���굥λ��
��Ϣ1=������Ϣ1
���˽�=���˽ɷ�
����Ϣ=������Ϣ
�������=������˽�
��Ϣ2=������Ϣ2
�ܼ�=���ۼ�
skip
enddo
if ����=1998  
   repl ���˱�Ϣ with �����ܱ�Ϣ , ���걾Ϣ with �������걾Ϣ for val(���)=1998
endif
close all
use ylgrzh
go top
*********У��***********
loca for val(���)=2002
if !eof()
gz_21=�½ɷѻ���
skip -1
if val(���)=2001
   gz_20=�½ɷѻ���
   REPLACE �½ɷѻ��� with round((gz_20*9+gz_21*3)/12,0)
endif
endif
************************
go top
clear
******(4)��˸����ʻ�������******
browse part 12 field ���: r ,rybh:r:h='���',RYXM:r:h='����' ;
 , �ɷ����� : 8 :H = '�ɷ�����' , �ɷѻ��� : 8 :H = '�ɷѻ���';
 , ��λ�ɷ� : 12 :H = '���굥λ�ɷ�' , ��ƽ5 : 7 :H = '��ƽ5%';
    , ��λ��Ϣ : 12 :H = '���굥λ��Ϣ' , ���˽ɷ� : 12 :H = '������˽ɷ�';
 , ������Ϣ : 12 :H = '���������Ϣ' , ����ϼ� : 10 :H = '����ϼ�';
 , ���굥λ�� : 12 :H = '���굥λ�ɷ�' , ������Ϣ1 : 12 :H =;
 '���굥λ��Ϣ' , ������˽� : 12 :H = '������˽ɷ�' , ������Ϣ2;
 : 12 :H = '���������Ϣ' , ���ۼ� : 10 :H = '��    ��';
 titl '���ϱ��ո����ʻ��Ǽǿ����'
yn = MESSAGEBOX("��Ҫ��ӡ��ת��ʹ�á������ʻ�����",4+32,"��ʾ")
IF yn = 6
copy to �����ʻ� type xl5 fiel ���,�ɷ�����,�½ɷѻ���,�ɷѻ���,��λ�ɷ�,��λ��Ϣ,����,���˽ɷ�,������Ϣ,����ϼ�,���굥λ��,������Ϣ1,������˽�,������Ϣ2,���ۼ�
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("d:\simis\�����ʻ�.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"   
ENDI    
if !file("�����ʻ�.DBF")
   copy to �����ʻ� type fox2x
else
   use �����ʻ� excl
   go top
   loca for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
   if !eof()
      delete for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
      pack
      appe from ylgrzh
   else
       appe from ylgrzh
   endif
endif        
close all
�˿� = '00'
delete file temp.dbf
delete file temp1.dbf
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
retu


