******************************
* .\YLBXJS.PRG (Build20161026)
******************************
close all
delete file lnk&LY1..dbf
YF0 = month(date())
YF00 = month(date())-1
RQ1 = day(date())
YF = iif(YF0>9,str(YF0,2),'0'+str(YF0,1))
YF1 = iif(YF00>9,str(YF00,2),'0'+str(YF00,1))
RQ111 = iif(RQ1>9,str(RQ1,2),'0'+str(RQ1,1))
close all
use ��������
go top
loca for ���='&LY'
PJGZ = ������ƽ��
LL = �������ʣ�/100
*if ��<=2006
*    LL = �������ʡ�/1000
*endif
LNLL = �������ʣ�
GRSY = ʧҵ���գ�
XS = ϵ��
wait window '��������'+str(val(LY)+1,4)+'�����ϱ���......' nowait
ZDM = space(3)
go top
loca for ���='&LY'
if val(LY)<=1997
  SN = ������ƽ��
else
  SN = 0
endif
DW0 = ��λ���ɣ�
GR0 = ���˽��ɣ�
SN0 = ������ƽ��
close all
IF file('ZR1'+LY+'.DBF')
**************�����ȴ����깤���ܶ�����������ʵ�ʽɷ�����������****20160720*********
sele 3
use zr1&ly1 
inde on rybh to ZR1&ly1
sele 2
use zr1&ly 
inde on rybh to ZR1&ly
sele 1
IF !FILE('yl'+LY+'.dbf')
use ylbxk
COPY TO yl&LY stru
USE yl&LY
appe from zr1&LY for ygxz='01'
***********�������ϱ��ձ����뵱��ʵ�ʽɷ��������������Ǻ�*********20160720********************
inde on rybh to YL&ly
upda on rybh from B repl �ɷ����� with b.mou
upda on rybh from C repl �½ɷѻ��� with c.jfjs
**********************����ȣ�LY��ʵ�ʽɷ���Ա���˶��ٸ��µĹ��ʾʹ��۴����˶��ٸ��µ���ᱣ�գ����ɷѻ���������ȣ�LY1�������ܶ�***********20160628*********************
ELSE
USE yl&LY
ENDIF
***********�����ֹ�ά�����ݣ��´μ��㲻�ظ�ά���������ݿ�����̫�ɲ���ʵ�ʽɷ�����һ�µĻ���������ɾ�����ݿ���������*********20160720********************
brow fiel cjmc:h='��λ����',bmmc:h='��������',ryxm:h='����', �ɷ����� , �½ɷѻ��� for �ɷ�����<12 titl '���������������Ա'+ly+'��ʵ�ʽɷ��������ԭ�������ݿ���������ʵ�ʽɷ���������ȵĻ�����ɾ�����ݿ��������ɣ�' 
ELSE
  MESSAGEBOX('���Ƚ���'+LY+'�깤���ܶ������','��ʾ')
  return
ENDIF
replace ��� with LY , A with 1 all
close all
******************��ʼ�������ʽ��������ʻ�ָ��*****20160720***********
select 2
use yl&ly1
index on RYBH to yl&ly1
select 1
use yl&ly excl
index on RYBH to yl&ly
*************************���������ʻ�ָ��********************
 replace �ɷѻ��� with �½ɷѻ���*�ɷ����� , �µ�λ�ɷ� with round(�½ɷѻ���*DW0/100,2) , ��λ�ɷ�;
 with �µ�λ�ɷ�*�ɷ����� , ��ƽ5 with round(SN*5/100,1)*�ɷ����� , �¸��˽ɷ�;
 with round(�½ɷѻ���*GR0/100,2) , ���˽ɷ� with round(�¸��˽ɷ�*�ɷ�����,2) , a with 1 all
  *****************************������Ա�����ʻ�ָ�����ά��*******************
 COUNT for �½ɷѻ���=0 or �ɷ�����=0 TO rs
IF rs>0
  BROWSE FIELDS cjmc:h='��λ����',bmmc:h='��������',ryxm:h='����', �ɷ����� , �½ɷѻ��� ,�ܼ� for �½ɷѻ���=0 or �ɷ�����=0 TITLE '������'+LY+'��ʵ�ʽɷ�������ʵ�ʽɷѻ�����'+LY1+'������˻��ܼ�'
  replace  �ɷѻ��� with �½ɷѻ���*�ɷ�����;
 , �µ�λ�ɷ� with round(�½ɷѻ���*DW0/100,2) , ��λ�ɷ� with �µ�λ�ɷ�*�ɷ�����;
 , ��ƽ5 with round(SN*5/100,1)*�ɷ����� , �¸��˽ɷ� with round(�½ɷѻ���*GR0/100;
,2) , ���˽ɷ� with round(�¸��˽ɷ�*�ɷ�����,2) all   
ENDIF
*******************���㡢�������걾Ϣ****��᣺����﷨�Ͻ���******************
 replace ������Ϣ with round(���˽ɷ�*LL*XS*1/2,2) , ����ϼ� with ���˽ɷ�+������Ϣ , ���ۼ� with  ����ϼ�  all
 upda on rybh from B replace ���ۼ� with  round(b.���ۼ�*(1+LNLL/100),2)+����ϼ� , �����ۼ� with b.���ۼ� , ���걾Ϣ with round(b.���ۼ�*(1+LNLL/100),2)
                                                                                               ************�������걾Ϣ********20160622*************************
 repl  ������� with ���˽ɷ� , �ۼ���Ϣ with ������Ϣ , ������ƽ with PJGZ , �������Ͻ� with round(������ƽ*0.2;
,1) , �˻����Ͻ� with round(���ۼ�/120,2) , ���ϴ����� with round(�������Ͻ�+�˻����Ͻ�;
,1) , ���˱�Ϣ with ���˽ɷ�+������Ϣ all
****************����Ա��**************
 replace ���ۼ� with  round(�ܼ�*(1+LNLL/100),2)+����ϼ� , �����ۼ� with �ܼ� for �ܼ�>0
**************************************
inde on bmbh+rybh to yl&ly
go top
brow fiel cjmc:h='��λ����',bmmc:h='��������',ryxm:h='����', �ɷ����� , �½ɷѻ��� , ���˽ɷ� , ������Ϣ , ����ϼ� , ���ۼ� , �ܼ�:h=ly1+'���ܼ�' , �����ۼ�  for �ɷ�����<12 titl '���������'+ly+'������ʻ����˽��ۼƴ����' 
pack
go top
brow fiel cjmc:h='��λ����',bmmc:h='��������',ryxm:h='����', �ɷ����� , �½ɷѻ��� , ���˽ɷ� , ������Ϣ , ����ϼ� , ���ۼ� , �����ۼ� , �ܼ�:h=ly1+'���ܼ�' titl '���������'+ly+'������ʻ����˽��ۼƴ����' 
pack
inde on rybh to yl&ly
CLOSE all
return
