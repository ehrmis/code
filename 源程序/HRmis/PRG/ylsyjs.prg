close all
use sbdmk
if empty(allt(dw)) or empty(allt(mc))
   MESSAGEBOX('���Ƚ���ϵͳά��������','��ʾ')
   retu
endif
close all
DO SRNd.SPr
*********����������ꡢǰ��*******************
LY1=str(val(LY)-1,4)
qn=str(val(LY)-2,4)
IF !file('zr1'+ly1+'.dbf')
MESSAGEBOX('���Ƚ���'+LY1+'�깤���ܶ������','��ʾ')
  RETURN
ENDIF
use ��������
go top
loca for ��� = '&LY'
go top
*����:��ʱά����ǰ��������*
loca for ��� = '&LY' 
if ������ƽ��=0 or �������ʣ�=0 or �������ʣ�=0 or ���˽��ɣ�=0 or ʧҵ���գ�=0 or ϵ��=0
   DO ly_gzze
endif
**************
do ylbxjs
**************
wait wind   LY+'�����ϱ��մ���ɹ�......'  nowa
**************
do sybxjs
**************
wait wind   LY+'��ʧҵ���մ���ɹ�......'  nowa
CLOSE ALL
use zr&LY
DBFF1='zr'
WHNY=LY+'�깤���ܶ�'
do Qdir
close all
use zr1&LY
DBFF1='zr1'
WHNY=str(val(LY)+1,4)+'��ɷѻ���'
do Qdir
CLOSE ALL
if �˿�='00'
use YL&LY
DBFF1='YL'
WHNY=LY+'�����ϱ���'
do Qdir
close all
use YL&LY
copy to LY+'�����ϱ���'
endif
close all
use SY&LY
DBFF1='SY'
WHNY=LY+'��ʧҵ����'
do Qdir
close all
use SY&LY
copy to LY+'��ʧҵ����'
close all
return