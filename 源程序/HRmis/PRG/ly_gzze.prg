*****************************************
* .\LY_GZZE.PRG(Build20150825)
*****************************************
LY = str(val(nd)-1,4)
******('�ɷ���(LY)'��'��Ȼ��'ת��(ND))
use ��������
go top
loca for ��� = '&nd'
if eof()
  append blank
  replace ��� with nd , ���˽��ɣ� with 8 , ʧҵ���գ� with 0.6 , ҽ�Ʊ��գ� with 2 ,��ҵ��� with 10 , ס�������� with 12
endif
go top
*����:��ʱά����ǰ��������*
loca for ��� = '&nd' 
if ������ƽ��=0
browse for ���=str(year(date()),4) titl '�뼰ʱ������ᱣ�ջ�������(�����������ꡰ��ƽ���ʡ����㣬����ʽ�������ݺ����޸ĵ���ʵ�ʹٷ���ᱣ�����ݣ�'
retu
endif
********************ÿ��6�·ݰ������ƽ�������������ɵ���ɷѻ����͸��˽ɷ�����***************************************
GR = ���˽��ɣ�/100
SN = ������ƽ��
DQ = ������ƽ��
SY = ʧҵ���գ�/100
BQ = sybxbl/100
fds=ROUND(300/100*SN,0)
bds=round(60/100*SN,0)
YB = ҽ�Ʊ��գ�/100
ybsx=ҽ������ֵ
ybxx=ҽ������ֵ
GJ = ס��������/100
gjjsx=����������
gjjxx=����������
on key label F1 do grcx
set safety off
ZDM = space(3)
close all
****��ʼ������
ND1 = val(LY)
YF1 = month(date())
RQ1 = day(date())
YF = iif(YF1>9,str(YF1,2),'0'+str(YF1,1))
RQ111 = iif(RQ1>9,str(RQ1,2),'0'+str(RQ1,1))
close all
dhsr=''
�� = year(date())-val(LY)
IF not file('Zr1'+LY+'.DBF')
************1.���㹤���ܶ�**************
close all
use sbdmk
gxyf=�����·�
use ��������
go top
loca for ��� = '&LY'
DWBL = 3
GRBL = 8
  use gzze1
  COPY TO zr1&LY stru
  USE zr1&LY
  if file('ry'+LY+'01.DBF')
     append from ry&LY.01 FOR 'ְ��'$ygxz1
  else
     append from ryzk FOR 'ְ��'$ygxz1
  endif
  replace MOU with 12 all
  close all
  select 2
  if ��>1
    ���� = str(year(date()),4)+'0'+str(gxyf,1)
    if file('gz'+����+'.dbf')
      use gz&����
    else
      use ryzk
    endif
  else
    use ryzk
  endif
  index on RYBH to ry
  select 1
  use zr1&LY
  inde on rybh to zr1&LY
  if val(LY)<2000
    update on RYBH from b replace zr3 with b.ylbx , PJ with round(zr3*100/GRBL,1),HJ with;
 round(PJ*MOU,1),jfjs with pj
  else
    update on RYBH from b replace zr3 with b.ylbx , PJ with round(zr3*100/GRBL,0),HJ with round(PJ*MOU,0),jfjs with pj
  endif
  close all
dh=''
dycj=''
dybm=''
do dhkdy
use zr1&LY excl
do case
case dh=2
set filter to cjdm='&dycj'
case dh=3
set filter to bmbh='&dybm'
endcase
inde on bmbh+rybh to zr1&LY
go top
************�˹�У��**************
loca for zr3=0
if !eof()
browse field CJMC :H = '  ��λ����  ' : 12 :R , RYXM :H = '��Ա����';
 : 8 :R , MOU :H = '�ɷ�����' : 8 , PJ :r: 8 :H = 'ƽ������' , У�� , zr8 :r: 8 :H = '��λ�ɷ�' , zr3 :r: 8 :H = '���˽ɷ�' for zr3=0 titl '������Ա[�� Esc �˳�]'
endif
pack
go top
browse field CJMC :H = '  ��λ����  ' : 12 :R , RYXM :H = '��Ա����';
 : 8 :R , MOU :H = '�ɷ�����' : 8 , PJ :r: 8 :H = 'ƽ������' , У�� , zr8 :r: 8 :H = '��λ�ɷ�' , zr3 :r: 8 :H = '���˽ɷ�' title '������˶�'+ly+'��ƽ�����ʺͽɷ�����[�� F1 ���Ҹ��ˡ��� Esc �˳�]'
pack
go top
REPL pj with У�� ,jfjs with pj,ZR3 with round(PJ*GRBL/100,1) , ZR8 with round(PJ*DWBL/100,1) FOR У��>0
repl zr5 with 0 , hj with pj*mou all
if val(LY)<=1997
  replace ZR5 with round(SN*5/100,1) , HEJI with ZR8+ZR5+ZR3 all
else
  replace HEJI with ZR8+ZR3 all
endif
set filter to
wait window '�����Զ�����'+LY+'�깤���ܶ'+str(val(ly)+1,4)+'��ɷѻ���......' nowait
close all
**************�˹�У�������ɹ����ܶ��zr&LY********************
IF i=88
use gzze
copy to zr&LY stru
close all
select 2
use zr1&LY
index on rybh to zr1&LY
select 1
use zr&LY
append from zr1&LY
index on rybh to zr&LY
update on RYBH from b replace MOU with b.MOU
go top
do while !eof()
for IJB = 1 to MOU
  CMJB = iif(IJB>9,str(IJB,2),'0'+str(IJB,1))
 upda on rybh from B replace M&cmjb with b.PJ 
****************ע���롰��������������**********
endfor
skip
enddo
replace A with 1 all
ENDIF
ELSE
close all
********2.���������ʹ���Ѵ���Ĺ����ܶ���´����߲�����������ɷѻ���jfjs����ᱣ��(��������)���˽ɷѽ��***********
use zr1&LY
REPLACE jfjs with pj all
REPL jfjs with У��  FOR У��>0
*************����������*****************************
repl zr8 with jfjs all
repl zr8 with ybsx for zr8>ybsx
repl zr8 with ybxx for zr8<ybxx
REPLACE ybx WITH ROUND(zr8*YB,2) all
***********������ҽ�Ʊ�������ִ�������޷ⶥ��������ͳһ������һ��ɷѻ���**************************
repl zr8 with jfjs all
repl zr8 with gjjsx for zr8>gjjsx
repl zr8 with gjjxx for zr8<gjjxx
REPLACE gjj WITH ROUND(zr8*GJ,0) all
***********������ס���������������ִ�������޷ⶥ��������ͳһ������һ��ɷѻ���**************************
repl jfjs with fds for jfjs>fds
repl jfjs with bds for jfjs<bds
repl zr3 with round(jfjs*GR,2),sybx with round(jfjs*SY,2) all
repl sybx with round(jfjs*BQ,2) FOR 'B��'$zzjg
*****************A����B���ֿ�����*************20150825********************************
repl heji WITH zr3+qynj+sybx+ybx+gjj all
go top
browse field CJMC :H = '  ��λ����  ' : 12 :R , RYXM :H = '��Ա����';
 : 8 :R , MOU :H = '�ɷ�����' , PJ :r:H = ly+'����ƽ������' , ���������� , jfjs :r:H =str(val(ly)+1,4)+'��ɷѻ���',У��, zr3 :r:H = '���ϱ���',sybx :r:H = 'ʧҵ����',ybx :r:H = 'ҽ�Ʊ���',gjj :r:H = 'ס��������' title '������˶�'+LY+'����ƽ�����ʺͽɷ�����[�� F1 ���Ҹ��ˡ��� Esc �˳�]'
ENDIF
close all
select 2
use ryzk
index on rybh to ryzk
select 1
use zr1&LY
index on rybh to zr1&LY
update on RYBH from b replace cjdm with b.cjdm,cjmc with b.cjmc,bmbh with b.bmbh,bmmc with b.bmmc
sort on bmbh,zw,rybh to temp
use temp
copy to zr1&LY
CLOSE ALL
use zr1&LY
DBFF1='zr1'
WHNY=LY+'�깤���ܶ�'
do Qdir
close all
delete file temp.dbf
retu

