***************************
*YLBXCL.prg(Build20151103)
***************************
�˿� = '00'
ƽ��2=0
ly=nd
LY1 = str(val(LY)-1,4)
if not file('yl'+LY1+'.DBF')
    MESSAGEBOX('��������'+LY1+'�����ϱ��գ���','��ʾ')
    return
ENDIF
use yl&ly1
if recc()=0
   MESSAGEBOX('��������'+LY1+'�����ϱ��գ���','��ʾ')
   return
endif
use ylbxk
COPY TO yltemp stru  
 USE yltemp  
 appe from ryzk  
replace ��� with LY , A with 1 all
close all
use grzh excl
lj="&xp"+"\backup\"
if file("&lj"+"grzh.dbf")
   zap
   appe from &lj.grzh
endif 
go top
loca for aae001=val(ly)
if !eof()
yn = MESSAGEBOX("�Ƿ��롰�����籣ϵͳ���нɷѻ�����������",4+32,"��ʾ")
IF yn = 6
   �˿� = '��Ƭ' 
ENDIF   
endif 
use ��������
go top
loca for ���='&LY'
PJGZ = ������ƽ��
LNLL = �������ʣ�
DW = ��λ���ɣ�
GR = ���˽��ɣ�
GRSY = ʧҵ���գ�
��=val(���)
XS = ϵ��
if ��<=1997
  SN = ������ƽ��
else
  SN = 0
endif
LL = �������ʣ�/100
if ��<=2006
    LL = �������ʡ�/1000
endif
close all
IF �˿� <> '��Ƭ' 
sele 3
use zr1&ly excl
delete for pj<1
pack
inde on rybh to ZR1&ly
ENDIF
sele 2
IF �˿� = '��Ƭ'
use grzh
sort on aae001 to temp for aae001=val(LY)
use temp
inde on allt(grbh) to temp
go top
ELSE
if not file('zr1'+LY1+'.DBF')
    MESSAGEBOX('���Ƚ���'+LY1+'�깤���ܶ������','��ʾ')
    return
endif
use zr1&ly1 excl
delete for pj<1
pack
inde on rybh to ZR1&ly1
ENDIF
sele 1
use yltemp
**************************************************
IF �˿� = '��Ƭ'
set relation to allt(grbh) into 2
go top
do while !eof() 
wait window '������д�ɷ�������������'+allt(ryxm)+'��'+allt(str(recn(),5))+'��'+allt(str(recc(),5))+'��......' nowait   
    sele 2
    sum aic079 to ys for aae001=val(LY) and allt(grbh)=allt(a.grbh)
    go top
    loca for aae001=val(LY) and allt(grbh)=allt(a.grbh)   
    if !eof()
       replace a.�ɷ����� with ys,a.�½ɷѻ��� with jfjs
    endif
    continue
    if !eof()
       ƽ��2=jfjs       
    endif
    sele 1
replace �ɷѻ��� with �½ɷѻ���*�ɷ�����;
  , �µ�λ�ɷ� with round(�½ɷѻ���*DW/100,1) , ��λ�ɷ�;
 with �µ�λ�ɷ�*�ɷ����� , ��ƽ5 with round(SN*5/100,1)*�ɷ����� , �¸��˽ɷ�;
 with round(�½ɷѻ���*GR/100,1) , ���˽ɷ� with �¸��˽ɷ�*�ɷ����� 
  if ��=2001 
    do case
    case �ɷ�����=1
      replace ��λ�ɷ� with round(ƽ��2*4/100,1) , ���˽ɷ� with round(ƽ��2*7/100,2) , �ɷѻ��� with ƽ��2
    case �ɷ�����=2
      replace ��λ�ɷ� with round(ƽ��2*4/100,1)*2 , ���˽ɷ� with round(ƽ��2*7/100,2)*2 , �ɷѻ��� with;
 ƽ��2*2
    case �ɷ�����=3
      replace ��λ�ɷ� with round(ƽ��2*4/100,1)*3 , ���˽ɷ� with round(ƽ��2*7/100,2)*3 , �ɷѻ��� with;
 ƽ��2*3
    case �ɷ�����>3
      replace ��λ�ɷ� with �µ�λ�ɷ�*(�ɷ�����-3)+round(ƽ��2*4/100,1)*3 , ���˽ɷ� with;
 �¸��˽ɷ�*(�ɷ�����-3)+round(ƽ��2*7/100,2)*3 , �ɷѻ��� with �½ɷѻ���*(�ɷ�����-3)+ƽ��2*3
    endcase  
  endif
  if ��=2002 
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
  skip
enddo
repl ��λ��Ϣ with round((��λ�ɷ�+��ƽ5)*LL*�ɷ�����/2,1);
 , ������Ϣ with round(���˽ɷ�*LL*�ɷ�����/2,2) , ����ϼ� with ��λ�ɷ�+��ƽ5+��λ��Ϣ+���˽ɷ�+������Ϣ all
ELSE
wait window '���ڴ���'+LY+'������ʻ�......' nowait 
set relation to rybh into 3
repl �ɷ����� with c.MOU all
repl �ɷ����� with 0 for �ɷ�����<0
index on RYBH to yltemp
 update on RYBH from B replace �½ɷѻ��� with b.jfjs;
 , �ɷѻ��� with �½ɷѻ���*�ɷ����� , �µ�λ�ɷ� with b.ZR8 , ��λ�ɷ�;
 with �µ�λ�ɷ�*�ɷ����� , ��ƽ5 with round(SN*5/100,1)*�ɷ����� , �¸��˽ɷ�;
 with b.ZR3 , ���˽ɷ� with �¸��˽ɷ�*�ɷ�����
go top
do while  not eof()
  if val(LY)=2001  
  wait window '���ڴ���2001��ɷѻ�����'+allt(ryxm)+'��'+allt(str(recn(),5))+'��'+allt(str(recc(),5))+'��......' nowait 
    do case
    case �ɷ�����=1
      replace ��λ�ɷ� with c.ZR8 , ���˽ɷ� with c.ZR3 , �ɷѻ��� with c.jfjs
    case �ɷ�����=2
      replace ��λ�ɷ� with c.ZR8*2 , ���˽ɷ� with c.ZR3*2 , �ɷѻ��� with;
 c.jfjs*2
    case �ɷ�����=3
      replace ��λ�ɷ� with c.ZR8*3 , ���˽ɷ� with c.ZR3*3 , �ɷѻ��� with;
 c.jfjs*3
    case �ɷ�����>3
      replace ��λ�ɷ� with �µ�λ�ɷ�*(�ɷ�����-3)+c.ZR8*3 , ���˽ɷ� with;
 �¸��˽ɷ�*(�ɷ�����-3)+c.ZR3*3 , �ɷѻ��� with �½ɷѻ���*(�ɷ�����-3)+c.jfjs*3
    endcase
  else
    exit
  endif
  skip
enddo
go top
do while  not eof()
  if  val(LY)=2002
   wait window '���ڴ���2002��ɷѱ�����'+allt(ryxm)+'��'+allt(str(recn(),5))+'��'+allt(str(recc(),5))+'��......' nowait 
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
  else
    exit
  endif
  skip
enddo
go top
do while  not eof()
  if val(LY)=2003
    do case
    case �ɷ�����=1
     replace ��λ�ɷ� with �µ�λ�ɷ� , ���˽ɷ� with �¸��˽ɷ�
    case �ɷ�����=2
     replace ��λ�ɷ� with �µ�λ�ɷ�*2 , ���˽ɷ� with �¸��˽ɷ�*2
    case �ɷ�����=3
     replace ��λ�ɷ� with �µ�λ�ɷ�*3 , ���˽ɷ� with �¸��˽ɷ�*3
    case �ɷ�����>3
     replace �ɷ����� with �ɷ�����-3, ��λ�ɷ� with �µ�λ�ɷ�*�ɷ����� , ���˽ɷ� with;
 �¸��˽ɷ�*�ɷ�����
    endcase
    repl �ɷѻ��� with �½ɷѻ���*�ɷ�����
  else
    exit
  endif
  skip
enddo
repl ��λ��Ϣ with round((��λ�ɷ�+��ƽ5)*LL*�ɷ�����/2,1);
 , ������Ϣ with round(���˽ɷ�*LL*�ɷ�����/2,1) , ����ϼ� with ��λ�ɷ�+��ƽ5+��λ��Ϣ+���˽ɷ�+������Ϣ all
ENDIF
close all
select 2
use yl&ly1
index on RYBH to yl&ly1
sele 1
use yltemp excl
index on RYBH to yltemp
*******************���㡢�������걾Ϣ****��᣺����﷨�Ͻ���******************
 IF val(LY)>=2007
 replace ������Ϣ with round(���˽ɷ�*LL*XS*1/2,2) , ����ϼ� with ���˽ɷ�+������Ϣ , ���ۼ� with  ����ϼ�  all
 upda on rybh from B replace ���ۼ� with  round(b.���ۼ�*(1+LNLL/100),2)+����ϼ� , �����ۼ� with b.���ۼ� 
 repl  ������� with ���˽ɷ� , �ۼ���Ϣ with ������Ϣ , ������ƽ with PJGZ , �������Ͻ� with round(������ƽ*0.2;
,1) , �˻����Ͻ� with round(���ۼ�/120,2) , ���ϴ����� with round(�������Ͻ�+�˻����Ͻ�;
,1) , ���˱�Ϣ with ���˽ɷ�+������Ϣ all
****************����Ա��**************
 replace ���ۼ� with  round(�ܼ�*(1+LNLL/100),2)+����ϼ� , �����ۼ� with �ܼ� for �ܼ�>0
**************************************
inde on bmbh+rybh to yl&ly
ELSE
upda on rybh from B replace ��λ��Ϣ with round((��λ�ɷ�+��ƽ5)*LL*�ɷ�����/2,2);
 , ������Ϣ with round(���˽ɷ�*LL*�ɷ�����/2,2) , ����ϼ�;
 with ��λ�ɷ�+��ƽ5+��λ��Ϣ+���˽ɷ�+������Ϣ , ���굥λ�� with b.��λ�ɷ�+b.��ƽ5+b.��λ��Ϣ+b.���굥λ��+b.������Ϣ1;
 , ������Ϣ1 with round(���굥λ��*(LNLL/100),2) , ������˽� with;
 b.���˽ɷ�+b.������Ϣ+b.������˽�+b.������Ϣ2 , ������Ϣ2 with round(������˽�*(LNLL/100);
,1) , ���ۼ� with ����ϼ�+���굥λ��+������Ϣ1+������˽�+������Ϣ2;
 , ������� with ��λ�ɷ�+���˽ɷ� , �ۼ���Ϣ with ��λ��Ϣ+������Ϣ+������Ϣ1+������Ϣ2;
 , �����ۼ� with b.���ۼ� , ������ƽ with PJGZ , �������Ͻ� with round(������ƽ*0.2;
,1) , �˻����Ͻ� with round(���ۼ�/120,2) , ���ϴ����� with round(�������Ͻ�+�˻����Ͻ�;
,1) , ���˱�Ϣ with ���˽ɷ�+������Ϣ+������˽�+������Ϣ2 , ���걾Ϣ with ������˽�+������Ϣ2
ENDIF
******��˸����ʻ�������******
go top
browse part 12 field ���: r ,rybh:r:h='���',RYXM:r:h='����' ;
 , �ɷ����� : 8 :H = '�ɷ�����' , �ɷѻ��� : 8 :H = '�ɷѻ���';
 , ��λ�ɷ� : 12 :H = '���굥λ�ɷ�' , ��ƽ5 : 7 :H = '��ƽ5%';
    , ��λ��Ϣ : 12 :H = '���굥λ��Ϣ' , ���� :h ='���˽ɷѱ���%' , ���˽ɷ� : 12 :H = '������˽ɷ�';
 , ������Ϣ : 12 :H = '���������Ϣ' , ����ϼ� : 10 :H = '����ϼ�';
 , ���굥λ�� : 12 :H = '���굥λ�ɷ�' , ������Ϣ1 : 12 :H =;
 '���굥λ��Ϣ' , ������˽� : 12 :H = '������˽ɷ�' , ������Ϣ2;
 : 12 :H = '���������Ϣ' , ���ۼ� : 10 :H = '��    ��'
delete for �ɷ�����=0
pack
if !file("�����ʻ�.DBF")
   copy to �����ʻ� type fox2x
else
  use �����ʻ� excl
  delete for val(���)=val(LY)  
  pack  
  appe from yltemp 
endif 
sort on bmbh,rybh,��� to temp
use temp
*********У��***********
loca for val(���)=2002
if !eof()
gz_21=�½ɷѻ���
skip -1
if val(���)=2001
   gz_20=�½ɷѻ���
   REPLACE �½ɷѻ��� with round((gz_20*9+gz_21*3)/12,0)
endif
continue
endif
************************
copy to �����ʻ� type fox2x
use �����ʻ�
repl ���� with GR all 
yn = MESSAGEBOX("��Ҫ��ӡ��ת��ʹ�á������ʻ�����",4+32,"��ʾ")
IF yn = 6
copy to &bf.\�����ʻ� type xl5
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\�����ʻ�.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"   
ENDI   
close all
�˿� = '00'
retu 