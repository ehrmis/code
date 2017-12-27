***************************
* .\GBGW.PRG
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
select * from 岗位管理 into dbf rygwzk where gwfl1<>'操作岗位'
do dhkDY
select 3
use rygwzk
repl gwfl1 with '管理人员' for ryfl='04'
repl gwfl1 with '专业人员' for ryfl<>'04'
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
  set printer to 管理岗位.TXT &&预览
  set printer on 
endi
do while not eof()
  CJBH = alltrim(cjDM)
  do case
  case DH=1
    select * from rygwzk into table rygw where CJDM='&CJBH'
  case DH=2
    select * from rygwzk into table rygw where CJDM='&DYCJ'
  case DH=3
    select * from rygwzk into table rygw where BMBH='&DYBM'
  endcase
  use
  select 1
  use rygw
  count for RYFL='04' to GLRY
  count for RYFL<>'04' to zyry
  count for allt(gwjb1)='十七岗' to gj17
  count for allt(gwjb1)='十六岗' to gj16
  count for allt(gwjb1)='十四岗' to gj14
  count for allt(gwjb1)='十三岗' to gj13
  count for allt(gwjb1)='十一岗' to gj11
  count for allt(gwjb1)='十岗' to gj10
  count for allt(gwjb1)='九岗' to gj9
  count for allt(gwjb1)='八岗' to gj8
  count for allt(gwjb1)='七岗' to gj7
  count for allt(gwjb1)='六岗' to gj6
  count for allt(gwjb1)='五岗' to gj5
  count for allt(岗等)='初职(员)' to YJ
  count for allt(岗等)='初职(助)' to ZJ
  count for allt(岗等)='中职' to ZZ
  count for allt(岗等)='副高职' to FGZ
  count for allt(岗等)='正高职' to ZGZ
  count for allt(岗等)='正处级' to zCJ
  count for allt(岗等)='副处级' to fCJ
  count for allt(岗等)='正科级' to ZK
  count for allt(岗等)='副科级' to FK
  count for allt(岗等)='一级科员' to YJKY
  count for allt(岗等)='二级科员' to LJKY
  count for allt(岗等)='三级科员' to SJKY
  ZRS = reccount()
  PG = 1
  go top
do while  not eof()
    @ dwh , 0 say space(10)+'&MC111'+'职工岗位调查明细表'  font '宋体',20
    @ prow()+2 , 0 say space(37)+'( 管理岗位、专业技术岗位 )' font '宋体',12
    @ prow()+2 , 8 say '单 位:'+alltrim(cjmc)+space(28)+'第'+str(PG,2)+'页'+space(22)+'日期:'+str(year(date()),4)+'年'+str(month(date()),2)+'月'+str(day(date()),2)+'日' &zt1
    T1 = '━━━━━━┳━━━━┯━━━━━━━━━━━━━━━━━━━━━━┯━━━━━━━┯━━━━┯━━━'
    T2 = '            ┃        │           行  政  技  术  职  务           │              │        │      '
    T3 = ' 部门(班组) ┃ 姓  名 ├──┬──┬──┬─┬─┬─┬─┬─┬─┬─┤现 所 在 岗 位│人员分类│岗级  '
    T4 = '            ┃        │ 员 │ 助 │ 中 │高│处│正│副│主│主│科│              │        │      '
    T5 = '  名   称   ┃        │ 职 │ 职 │ 职 │职│级│科│科│管│办│员│  (  工种  )  │        │      '
    T6 = '──────╂────┼──┼──┼──┼─┼─┼─┼─┼─┼─┼─┼───────┼────┼───'
    T7 = '──────╂────┼──┴──┴──┴─┴─┴─┴─┴─┴─┴─┼───────┼────┼───'
    T8 = '━━━━━━┻━━━━┷━━━━━━━━━━━━━━━━━━━━━━┷━━━━━━━┷━━━━┷━━━'
    T9 = '━━━━━━┳━━━━┯━━┯━━┯━━┯━┯━┯━┯━┯━┯━┯━┯━━━━━━━┯━━━━┯━━━'
    @ prow()+1 , 0 say T1 &zt
    @ prow()+1 , 0 say T2 &zt
    @ prow()+1 , 0 say T3 &zt
    @ prow()+1 , 0 say T4 &zt
    @ prow()+1 , 0 say T5 &zt
    @ prow()+1 , 0 say T6 &zt
    N = 1
    do while N<yjl
    IF  not eof()
        @ prow()+1 , 0 say bmmc+'┃' &zt
        @ prow() , pcol() say ryxm &zt
      if allt(岗等)='初职(员)' 
         @ prow() , pcol() say '│ √ │    │    │  │  │  │  │  │  │  │' &zt
      endif
      if allt(岗等)='初职(助)'
         @ prow() , pcol() say '│    │ √ │    │  │  │  │  │  │  │  │' &zt
      endif
      if allt(岗等)='中职'
         @ prow() , pcol() say '│    │    │ √ │  │  │  │  │  │  │  │' &zt
      endif
      if '高职'$岗等
         @ prow() , pcol() say '│    │    │    │√│  │  │  │  │  │  │' &zt
      endif
      if '处级'$岗等
         @ prow() , pcol() say '│    │    │    │  │√│  │  │  │  │  │' &zt
      endif
      if allt(岗等)='正科级'
         @ prow() , pcol() say '│    │    │    │  │  │√│  │  │  │  │' &zt
      endif
      if allt(岗等)='副科级'
         @ prow() , pcol() say '│    │    │    │  │  │  │√│  │  │  │' &zt
      endif
      if allt(岗等)='业务主管'
         @ prow() , pcol() say '│    │    │    │  │  │  │  │√│  │  │' &zt
      endif
      if allt(岗等)='业务主办'
         @ prow() , pcol() say '│    │    │    │  │  │  │  │  │√│  │' &zt
      endif
      if allt(岗等)='业务员'
         @ prow() , pcol() say '│    │    │    │  │  │  │  │  │  │√│' &zt
      endif
      if !empty(zw)
         @ prow() , pcol() say left(gw,14)+'│' &zt
      else
         @ prow() , pcol() say gz1+'  │' &zt
      endif
      @ prow() , pcol() say left(gwfl1,8)+'│' &zt
      @ prow() , pcol() say gwjb1 &zt
      @ prow()+1 , 0 say T6 &zt
     ENDIF
      if mod(N,yjl)=0
        @ prow()+1 , 0 say T7 &zt
      endif
      if  not eof()
        N = N+1
        skip
      else
        @ prow()+1 , 0 say ' 小  计(人) ' &zt
        @ prow() , pcol() say '┃'+str(glry+zyry,3) &zt
        @ prow() , pcol() say '     │' &zt
        @ prow() , pcol() say yJ picture '@Z 9999' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say zJ picture '@Z 9999' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say zz picture '@Z 9999' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say zgz+fgz picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say Zcj+fcj picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say Zk picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say fk picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say yjKY picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say ljky picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say sjky picture '@Z 99' &zt
        @ prow() , pcol() say '│管理人员:' &zt
        @ prow() , pcol() say GLRY picture '@Z 999' &zt
        @ prow() , pcol() say ' 专业技术人员:' &zt
        @ prow() , pcol() say zyry picture '@Z 999' &zt
        @ prow()+1 , 0 say T7 &zt
@ prow()+1 , 0 say ' 总  计(人) ┃'+str(yj+zj+zz+fgz+zgz+zcj+fcj+zk+fk+yjky+ljky+sjky,3)+'     │'+space(9)+'技术岗位:'+str(yj+zj+zz+fgz+zgz,3)+'  管理岗位:'+str(zcj+fcj+zk+fk+yjky+ljky+sjky,3)+'         │' &zt
@ prow() , pcol() say '平均岗:'+str(round((gj5*5+gj6*6+gj7*7+gj8*8+gj9*9+gj10*10+gj11*11+gj13*13+gj14*14+gj16*16+gj17*17)/ZRS,1),4,1)+'岗' &zt
@ prow()+1 , 0 say T8 &zt
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
  count for RYFL='04' to GLRY
  count for RYFL<>'04' to zyry
  count for allt(gwjb1)='十七岗' to gj17
  count for allt(gwjb1)='十六岗' to gj16
  count for allt(gwjb1)='十四岗' to gj14
  count for allt(gwjb1)='十三岗' to gj13
  count for allt(gwjb1)='十一岗' to gj11
  count for allt(gwjb1)='十岗' to gj10
  count for allt(gwjb1)='九岗' to gj9
  count for allt(gwjb1)='八岗' to gj8
  count for allt(gwjb1)='七岗' to gj7
  count for allt(gwjb1)='六岗' to gj6
  count for allt(gwjb1)='五岗' to gj5
  count for allt(岗等)='初职(员)' to YJ
  count for allt(岗等)='初职(助)' to ZJ
  count for allt(岗等)='中职' to ZZ
  count for allt(岗等)='副高职' to FGZ
  count for allt(岗等)='正高职' to ZGZ
  count for allt(岗等)='正处级' to zCJ
  count for allt(岗等)='副处级' to fCJ
  count for allt(岗等)='正科级' to ZK
  count for allt(岗等)='副科级' to FK
  count for allt(岗等)='业务主管' to YJKY
  count for allt(岗等)='业务主办' to LJKY
  count for allt(岗等)='业务员' to SJKY
ZRS = reccount()
@ prow()+1 , 0 say T9 &zt
@ prow()+1 , 0 say '合   计(人) ┃' &zt
        @ prow() , pcol() say str(glry+zyry,3) &zt
        @ prow() , pcol() say '     │' &zt
        @ prow() , pcol() say yJ picture '@Z 9999' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say zJ picture '@Z 9999' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say zz picture '@Z 9999' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say zgz+fgz picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say Zcj+fcj picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say Zk picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say fk picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say yjKY picture '@Z 99' &zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say ljky picture '@Z 99'&zt
        @ prow() , pcol() say '│' &zt
        @ prow() , pcol() say sjky picture '@Z 99' &zt
        @ prow() , pcol() say '│管理人员:' &zt
        @ prow() , pcol() say GLRY picture '@Z 999' &zt
        @ prow() , pcol() say ' 专业技术人员:' &zt
        @ prow() , pcol() say zyry picture '@Z 999' &zt
        @ prow()+1 , 0 say T7 &zt
@ prow()+1 , 0 say '合   计(人) ┃'+str(yj+zj+zz+fgz+zgz+zcj+fcj+zk+fk+yjky+ljky+sjky,3)+'     │'+space(9)+'技术岗位:'+str(yj+zj+zz+fgz+zgz,3)+'  管理岗位:'+str(zcj+fcj+zk+fk+yjky+ljky+sjky,3)+'         │' &zt
@ prow() , pcol() say '平均岗:'+str(round((gj5*5+gj6*6+gj7*7+gj8*8+gj9*9+gj10*10+gj11*11+gj13*13+gj14*14+gj16*16+gj17*17)/ZRS,1),4,1)+'岗' &zt
@ prow()+1 , 0 say T8 &zt
@ prow()+1 , 0 say space(10)+'劳资负责人:'+'&RS111'+space(45)+'制表人:'+'&ZB111' &zt1
set device to screen
if dh1=2
   set printer to lpt1
endif
set printer off
set printer to
close all
if dh1=1
   clear
   modi file 管理岗位.txt
endi
clear windows
delete file rygw.dbf
clear
return
