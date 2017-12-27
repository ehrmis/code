***************************
* .\GRGW.PRG
***************************
close all
CJBH = ''
DH = ''
DH1 = ''
DYCJ = ''
DYBM = ''
dwh=3
yjl=30
zt="font '宋体',11"
zt1="font '宋体',12"
YF1=month(date())
ND1=year(date())
ND = str(ND1,4)
if YF1>9
  YF = str(YF1,2)
else
  YF = '0'+str(YF1,1)
endif
NY = ND+YF
********当前系统年月*************
NYR = Right(NY,4)
YFs = iif(val(yf)>9,str(val(yf)-1,2),'0'+str(val(yf)-1,1))
IF val(yf)=1
   NY1 = str(year(date())-1,4)+'12'
   YFs = '12'
else   
   if val(yf)=10
      NY1 = str(year(date()),4)+'09'
   ELSE
      NY1 = str(year(date()),4)+YFs
    endif   
ENDIF
********当前系统年月[上月表达式]*************
if !file('gw'+ny+'.dbf')
=MESSAGEBOX('请先审核维护本月工种岗位！', 16,"提示")
 return
endif
store 0 to gj1 , gj2 , gj3 , gj4 , gj5 , gj6 , gj7 , gj8 , gj9 , sf
select * from 岗位管理 where gwfl1='操作岗位' into dbf ccrygwzk
close all
do DHKDY
select 3
use ccrygwzk
select 2
use cjk
go top
do dhkdy1
wait wind "正在打印，请稍候······" nowa
if dh1=2
  do srdwh.spr
  set device to printer
  set printer on
else
  set device to printer
  set printer to 操作岗位.TXT &&预览
  set printer on 
endi
do while  not eof()
  CJBH = alltrim(cjDM)
  do case
  case DH=1
    select * from ccrygwzk into dbf rygw where CJDM='&CJBH'
  case DH=2
    select * from ccrygwzk into dbf rygw where CJDM='&DYCJ'
  case DH=3
    select * from ccrygwzk into dbf rygw where BMBH='&DYBM'
  endcase
  use
  select 1
  use rygw
  count for allt(gwjb1)='九岗' to gj9
  count for allt(gwjb1)='八岗' to gj8
  count for allt(gwjb1)='七岗' to gj7
  count for allt(gwjb1)='六岗' to gj6
  count for allt(gwjb1)='五岗' to gj5
  count for allt(gwjb1)='四岗' to gj4
  count for allt(gwjb1)='三岗' to gj3
  count for allt(gwjb1)='二岗' to gj2
  count for allt(gwjb1)='一岗' to gj1
  ZRS = reccount()
  count for sfgj>0 to sf
  PG = 1
  go top
  do while  not eof()
    @ dwh , 0 say space(10)+'&MC111'+'职工岗位调查明细表' font '宋体',20
    @ prow()+2 , 0 say space(45)+'( 操作岗位 )' font '宋体',12
    @ prow()+2 , 8 say '单 位:'+alltrim(cjmc)+space(30)+'第'+str(PG,2)+'页'+space(25)+'日期:'+str(year(date()),4)+'年'+str(month(date()),2)+'月'+str(day(date()),2)+'日' &zt1
    T1 = '━━━━━━┳━━━━┯━━━━┯━━━┯━━━━━━┯━━┯━━━━┯━━━━━━┯━━━━━━━━┯━━━'
    T2 = ' 班组名称   ┃ 姓  名 │浮前岗级│档  级│职业资格培训│班长│技术等级│ 现从事工种 │ 现 所 在 岗 位 │ 工龄 '
    T3 = '──────╂────┼────┼───┼──────┼──┼────┼──────┼────────┼───'
    T4 = '━━━━━━┻━━━━┷━━━━┷━━━┷━━━━━━┷━━┷━━━━┷━━━━━━┷━━━━━━━━┷━━━'
    @ prow()+1 , 0 say T1 &zt
    @ prow()+1 , 0 say T2 &zt
    @ prow()+1 , 0 say T3 &zt
    N = 1
    do while N<yjl
     if  not eof()
      @ prow()+1 , 0 say bmmc+'┃' &zt
      @ prow() , pcol() say ryxm &zt
      @ prow() , pcol() say '│'+space(1) &zt
      @ prow() , pcol() say gwjb1 &zt
      @ prow() , pcol() say space(1)+'│' &zt
      @ prow() , pcol() say gwdc1 &zt
      @ prow() , pcol() say '│' &zt
      @ prow() , pcol() say left(pxmc,12)+'│' &zt
      if !empty(bzf)
         @ prow() , pcol() say ' √ │' &zt
      else
         @ prow() , pcol() say '    │' &zt
      endif
      @ prow() , pcol() say 岗等 &zt 
      @ prow() , pcol() say '│' &zt     
      @ prow() , pcol() say left(gz1,12)+'│'+left(gw,16) &zt
      @ prow() , pcol() say '│'+str(GL,2,0)+'年' &zt
      @ prow()+1 , 0 say T3 &zt
     endif
     if mod(N,yjl)=0
        @ prow()+1 , 0 say T4 &zt
      endif
      if  not eof()
        N = N+1
        skip
      else       
        @ prow()+1 , 0 say ' 小  计(人) ' &zt
        @ prow() , pcol() say '┃' &zt
        @ prow() , pcol() say ZRS picture '@ 99999999' &zt
        @ prow() , pcol() say '│' &zt
@ prow() , pcol() say '平均岗：'+str(round((gj9*9+gj8*8+gj7*7+gj6*6+gj5*5+gj4*4+gj3*3+gj2*2+gj1)/ZRS;
,1),4,1)+'岗   上浮：'+str(sf,3)+'人' &zt
        @ prow()+1 , 0 say T4 &zt
        exit
      endif
    enddo
    PG = PG+1
  enddo
  if DH<>1
    exit
  endif
  select 2
  skip
enddo
select 3
  count for allt(gwjb1)='九岗' to gj9
  count for allt(gwjb1)='八岗' to gj8
  count for allt(gwjb1)='七岗' to gj7
  count for allt(gwjb1)='六岗' to gj6
  count for allt(gwjb1)='五岗' to gj5
  count for allt(gwjb1)='四岗' to gj4
  count for allt(gwjb1)='三岗' to gj3
  count for allt(gwjb1)='二岗' to gj2
  count for allt(gwjb1)='一岗' to gj1
  count for sfgj>0 to sf
ZRS = reccount()
@ prow()+1 , 0 say T1 &zt
@ prow()+1 , 0 say ' 合      计 ┃' &zt
@ prow() , pcol() say ZRS picture '@ 99999999' &zt
@ prow() , pcol() say '│' &zt
@ prow() , pcol() say '平均岗:'+str(round((gj9*9+gj8*8+gj7*7+gj6*6+gj5*5+gj4*4+gj3*3+gj2*2+gj1)/ZRS;
,1),4,1)+'岗   上浮：'+str(sf,3)+'人' &zt
@ prow()+1 , 0 say T4 &zt
@ prow()+1 , 0 say space(20)+'劳资负责人:'+'&RS111'+space(45)+'制表人:'+'&ZB111' &zt1
set device to screen
if dh1=2
   set printer to lpt1
endif
set printer off
set printer to
close all
if dh1=1
   clear
   modi file 操作岗位.txt
   *run notepad 操作岗位.txt
endi
clear windows
delete file rygw.dbf
return
