**************************
* �ɷѹ���.prg(2008.1.7)
**************************
nd=str(year(date()),4)
clear
_SCREEN.PICTURE = ''
@ 1,36 say '�� �� �� ʾ' font '����',20
@ 4,10 say '��1������˳��������깤���ܶ���㹤����' font '����',14
@ 6,10 say '��2���Զ������ϱ��籣�����ӱ�D:\�걨����.XLS�����ӱ�' font '����',14
@ 8,10 say '��3����򿪸õ��ӱ�Ҫ��༭�걨�̻�ERPģ�壡����' font '����',14
yn = MESSAGEBOX("׼��������",4+32,"��ʾ")
if yn=7
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
endif
@ 14,10 say "�����뵱ǰ��Ҫ�걨�ɷѻ����Ľɷ��꣺" font '����',14 get nd font '����',18
read
sn=str(val(nd)-1,4)
if file('zr1'+sn+'.dbf')
use �걨���� excl
zap
appe from zr1&sn
repl ���� with cjmc,���� with bmmc,���� with ryxm,erp��� with erprybh,���˱�� with scbh,���֤�� with sfz,�����ܶ� with hj,��ƽ�� with pj,�ɷѻ��� with jfjs,���ϱ��� with zr3 all

copy to D:\�걨���� fiel rybh,����,����,����,erp���,���˱��,��������,��������,��λ����,���֤��,�����ܶ�,��ƽ��,�ɷѻ���,���ϱ��� type XL5
yn = MESSAGEBOX("��D:\�걨���������ӱ��ѵ������ִ�ʹ����",4+32,"��ʾ")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
*myexcel.sheetsinnewworkbook=1
*myexcel.visible=.t.
*myexcel.workbooks.add
*myexcel.worksheets("sheet1").activate
*next
myexcel.workbooks.open("D:\�걨����")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"  
ENDIF
else
=MESSAGEBOX("���ݵ���ʧ�ܣ����ȴ���"+sn+"�깤���ܶ�",48,"��ʾ")
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
endif
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