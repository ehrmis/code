***************************
* .\lwgz.PRG (Buld201405)
***************************
close all
YF0 = month(date())
YF00 = month(date())-1
RQ1 = day(date())
YF2 = iif(YF0>9,str(YF0,2),'0'+str(YF0,1))
YF1 = iif(YF00>9,str(YF00,2),'0'+str(YF00,1))
LY = iif(YF00=12,str(year(date()-1),4),str(year(date()),4))
YF = iif(RQ1>10,YF2,YF1)
use lwze
COPY TO lw&LY stru
USE lw&LY
if file('RY'+LY+'01.DBF')
***************************���������Ա********************
 yn = MESSAGEBOX("��Ҫ����һ�·��ڲ�ְ���������ѵ���ְ���������ܶ���",4+32,"��ʾ")
IF yn = 6
 append from ry&LY.01 
ELSE
 append from ryzk for '������ǲ��'$ygxz1
ENDIF 
else
 append from ryzk for '������ǲ��'$ygxz1
endif
close all
************************ͨ�ù��ܣ��غż��***********
use lw&LY
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
   delete file lw&LY..dbf
   sort on bmbh,rybh to lw&LY
endif
close all
select 1
use lw&LY
inde on rybh to lw&LY
M1 = 1
����=0
repl gzhj with 0,ʵ����� with 0,��Ч���� with 0,�����ν� with 0,��� with 0 all
***********�﷨�Ͻ��ԣ�ѭ���ݼӳ�ʼ��*****������ͻ�Խ��Խ���ظ���Ӷ��ٱ���************
go top
do while M1<=val(yf)
  MGZ = iif(M1>9,str(M1,2),'0'+str(M1,1))
  wait wind  '������д'+allt(str(val(mgz)))+'��������ǲ�������ܶ�...'  nowa
  if  not file('LW'+LY+MGZ+'.dbf')   
    MESSAGEBOX("δ����"+LY+"��"+allt(str(val(MGZ)))+"�¹�������",48,"��ʾ")
    retu
  endif  
    select 2
    use lwpqgzk
    copy to temp stru
    use temp
    appe from lw&LY.&Mgz
    repl a with 1 all
    sum a to rs
    ����=����+rs 
    index on RYBH to temp
    select 1
    upda on rybh from B replace ʵ����� with ʵ�����+b.ʵ�����,M&Mgz with b.ʵ�����-b.��Ч����,gzhj with gzhj+M&Mgz;
��Ч���� with ��Ч����+b.��Ч����,�����ν� with �����ν�+b.�����ν�,��� with ���+b.���
   M1 = M1+1
enddo
����=round(����/(M1-1),0)
close all
M1 = 1
select 1
use lw&LY
inde on rybh to lw&LY
repl jjhj with 0 all
go top
do while M1<=val(yf)
  MGZ = iif(M1>9,str(M1,2),'0'+str(M1,1))
  wait wind '������д'+allt(str(M1,2))+'�·ݼ�Ч����... ....' nowa
  if  not file('lw'+LY+MGZ+'.dbf')
    exit
  else
    select 2
    use lw&LY.&Mgz
    index on RYBH to lw&LY.&Mgz
    select 1
    upda on rybh from B replace J&Mgz with b.��Ч����,jjhj with jjhj+J&Mgz
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
index on bmbh+rybh to lw&LY
********************���ɹ����ܶ�(HJ)����ƽ��(PJ)**************
repl gzhj with m01+m02+m03+m04+m05+m06+m07+m08+m09+m10+m11+m12,jjhj with j01+j02+j03+j04+j05+j06+j07+j08+j09+j10+j11+j12,; 
hj with gzhj+jjhj,PJ with round(HJ/MOU,0) for mou>0
brow
use LW&LY
DBFF1='lw'
WHNY=LY+'�������'
do Qdir
use lw&LY
pf='d:\'+LY
copy to &pf.������� type xl5 
 =messagebox("�ѳɹ�������D��\����ѵ��ӱ�",48,"��ϲ")
  yn = MESSAGEBOX("��Ҫʹ�õ��ӱ�༭��ӡ������",4+32,"��ʾ")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           pf1="&pf"
           myexcel.workbooks.open(pf1+"�������.xls")
           myexcel.visible=.t.
           myexcel.caption="ʹ�õ��ӱ�༭��ӡ����" 
           release oleapp
           WAIT CLEAR
ENDIF     
close all

