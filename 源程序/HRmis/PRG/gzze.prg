***************************
* .\gzze.PRG (Buld201405)
***************************
close all
dhsr=''
set safety off
on key label F1 do grcx
**************˫λ�����Զ�����*******************
nd1=year(date())
yf1=month(date())
do srny2.spx
NY = LY+YF
LY1=str(val(LY)-1,4)
YF1 = iif(val(yf)>10,str(val(yf)-1,2),'0'+str(val(yf)-1,1))
NY1 = LY+YF1
if val(yf)=1
   NY1 = LY1+'12'
   YF1 = '12'
endi
use sbdmk
��ʼ�� = ��ʼ�·�
*************************************************
use gzze
COPY TO zr&LY stru
USE zr&LY
if file('RY'+LY+'01.DBF')
***************************���������Ա********************
 yn = MESSAGEBOX("��Ҫ����һ�·��ڲ�ְ���������ѵ���ְ���������ܶ���",4+32,"��ʾ")
IF yn = 6
 append from ry&LY.01 
ELSE
 append from ryzk
ENDIF 
else
 append from ryzk
endif
************************ͨ�ù��ܣ��غż��***********
use zr&LY
repl a with 1 all
sort on rybh to temp
use temp excl
go top
do while !eof()
   bh=rybh
   skip
   if bh=rybh
      repl a with 0
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a=0 to aaa
if aaa>0
   inde on bmbh+rybh to temp
   go top
   brow for a=0 titl '����'+allt(str(aaa,4))+'���غţ����޸���ȷ��[ Ctrl+T ]ɾ�������¼'
   pack
   delete file zr&LY..dbf
   sort on bmbh,rybh to zr&LY
endif
close all
select 1
use zr&LY
inde on rybh to zr&LY
M1 = 1
����=0
repl gzhj with 0,sfje with 0,ylbx with 0,qynj with 0,sybx with 0,ybx with 0,gjj with 0,sds with 0 all
***********�﷨�Ͻ��ԣ�ѭ���ݼӳ�ʼ��*****������ͻ�Խ��Խ���ظ���Ӷ��ٱ���************
go top
do while M1<=val(yf)
  MGZ = iif(M1>9,str(M1,2),'0'+str(M1,1))
  wait wind  '������д'+allt(str(val(mgz)))+'��ְ�������ܶ�...'  nowa
  if  not file('GZ'+LY+MGZ+'.dbf') 
  *or not file('TY'+LY+MGZ+'.dbf')
    ***************ERP****************************
    MESSAGEBOX("�뵼��ERPϵͳ��"+LY+"��"+allt(str(val(MGZ)))+"�¹��ʽ�������",48,"��ʾ")
    retu
  endif  
***********���˲�����Ա�����޹��ʵ���ylbx**********
    select 2
    use gzk
    copy to temp stru
    use temp
    appe from gz&LY.&Mgz
    repl a with 1 all
    sum a to rs
    ����=����+rs 
    if file('TY'+LY+Mgz+'.dbf')
       appe from TY&LY.&Mgz
    endif
    index on RYBH to temp
    select 1
    upda on rybh from B replace sfje with sfje+b.sfje,M&Mgz with b.hj-b.jang,gzhj with gzhj+M&Mgz;
ylbx with ylbx+b.ylbx,qynj with qynj+b.qynj,ybx with ybx+b.ybx,sybx with sybx+b.sybx,gjj with gjj+b.gjj,sds with sds+b.sds 
    M1 = M1+1
enddo
����=round(����/(M1-1),0)
close all
M1 = 1
select 1
use zr&LY
inde on rybh to zr&LY
repl jjhj with 0 all
go top
do while M1<=val(yf)
  MGZ = iif(M1>9,str(M1,2),'0'+str(M1,1))
  wait wind '������д'+allt(str(M1,2))+'�·ݵ��... ....' nowa
  if  not file('gz'+LY+MGZ+'.dbf')
    exit
  else
    select 2
    use gz&LY.&Mgz
    index on RYBH to gz&LY.&Mgz
    select 1
    upda on rybh from B replace J&Mgz with b.jang,jjhj with jjhj+J&Mgz
    M1 = M1+1
  endif
enddo
wait wind  '��������ʵ�ʽɷ�����...'  nowa
go top
do while !eof()
   M111=1
   for MJB1 = 1 to 12
    J111 = iif(MJB1>9,str(MJB1,2),'0'+str(MJB1,1))
    if (M&J111>0 or J&J111>0) and MOU<>12
       replace MOU with M111
       M111=M111+1
    endif    
  endfor
  skip
enddo
index on bmbh+rybh to zr&LY
********************���ɹ����ܶ�(HJ)����ƽ��(PJ)**************
repl gzhj with m01+m02+m03+m04+m05+m06+m07+m08+m09+m10+m11+m12,jjhj with j01+j02+j03+j04+j05+j06+j07+j08+j09+j10+j11+j12,; 
hj with gzhj+jjhj,PJ with round(HJ/MOU,0) for mou>0
USE gzze1
COPY TO zr1&ly stru
APPEND FROM zr&ly
********���ɽɷѻ���*******************
use zr&LY
DBFF1='zr'
WHNY=LY+'�깤���ܶ�'
do Qdir
use zr&LY
pf='d:\'+LY
copy to &pf.�깤���ܶ� type xl5 
 =messagebox("�ѳɹ�������D��\�����ܶ���ӱ�",48,"��ϲ")
  yn = MESSAGEBOX("��Ҫʹ�õ��ӱ�༭��ӡ������",4+32,"��ʾ")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           pf1="&pf"
           myexcel.workbooks.open(pf1+"�깤���ܶ�.xls")
           myexcel.visible=.t.
           myexcel.caption="ʹ�õ��ӱ�༭��ӡ����" 
           release oleapp
           WAIT CLEAR
ENDIF     
close all
do dwgzze
