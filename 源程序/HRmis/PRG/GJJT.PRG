*****************************
* GJJT.prg( Build 20050822 )
*****************************
close all
use dmk
jzy=对帐基准月
定位=压感纸定位
close all
ny=nd+yf
sy=iif(month(date())>9,str(month(date()),2),'0'+str(month(date()),1))
sn=str(year(date())-1,4)+"年"+allt(jzy)+"月"
if file('kg'+dm111+right(nd,2)+yf+'.dbf')
   jn=nd+"年"+allt(str(val(yf)))+"月"
   if file('kg'+ny+'.dbf')
      gjjny=ny
   else
      gjjny=dm111+right(nd,2)+yf
   endif 
else
   jn=nd+"年"+allt(str(val(yf)-1))+"月"
    if file('kg'+nd+sy+'.dbf')
       gjjny=nd+sy
    else
       gjjny=dm111+right(nd,2)+sy
    endif   
endif
S0="                                           住房公积金个人结存核对单"
S2="┌──────────┬────┬───────┬───────┬───────┬───────┬─────────┐"
S3="│      个人帐号      │人员姓名│ 期初余额(元) │ 期间缴纳(元) │ 本期利息(元) │ 累计本息(元) │    备     注     │"
S4="├──────────┼────┼───────┼───────┼───────┼───────┼─────────┤"
S5="└──────────┴────┴───────┴───────┴───────┴───────┴─────────┘"
***************************************
dh=''
dycj=''
dybm=''
do fjdy
***面向命令报表程序分级打印解决方案****
do case
 case dh=1
 use kg&gjjny
 copy to temp for gjjz<>0
 case dh=2
 use kg&gjjny
 copy to temp for cjdm=dycj.and.gjjz<>0
 case dh=3
 use kg&gjjny
 copy to temp for bmbh=dybm.and.gjjz<>0
endcase
set device to print
set print on
set print font "宋体",8
*********************▲****************
clear
set colo to 0/7
***针对fjdy.prg颜色设置问题***********
 sele 2
 use ryzk
 inde on rybh to ryzk
 sele 1
 use temp
 set relation to rybh into 2
 go top
 xh=0
 do while !eof()
  do while !eof()
  do case
  case dh=1
       WAIT WINDOW NOWAIT '正在打印住房公积金个人结存核对单<共'+allt(str(recc()))+'人>... ...' 
  case dh=2
       WAIT WINDOW NOWAIT '正在打印'+cjdm+'：'+allt(cjmc)+'<共'+allt(str(recc()))+'人>... ...' 
  case dh=3
       WAIT WINDOW NOWAIT '正在打印'+cjdm+'：'+allt(cjmc)+'、'+allt(bmmc)+'<共'+allt(str(recc()))+'人>... ...'       
  endcase
  if mod(xh,(22-定位)/2)=0
     @ 0,0 say S0
  else
     @prow()+2,0 say S0
  endif  
  S1=" 车间："+cjmc+"   部门："+bmmc+"                              核对期间："+sn+"至"+jn
  @prow()+1,0 say S1
  @prow()+1,0 say S2
  @prow()+0.8,0 say S3
  @prow()+0.8,0 say S4
  @prow()+0.8,0 say "│"
  @prow(),pcol() say gjjzh
  @prow(),pcol() say "│"
  @prow(),pcol() say RYxm
  @prow(),pcol() say "│"
  @prow(),pcol() say b.ydhj1 PICT "@Z 999,999,999.99"
  @prow(),pcol() say "│"
  @prow(),pcol() say snjn+bnjn PICT "@Z 999,999,999.99"
  @prow(),pcol() say "│"
  @prow(),pcol() say bnlx PICT "@Z 999,999,999.99"
  @prow(),pcol() say "│"
  @prow(),pcol() say gjjz+bnlx PICT "@Z 999,999,999.99"
  @prow(),pcol() say "│"
  @prow(),pcol() say "本次利息结至"+allt(jzy)+"月份 │"
  @prow()+0.8,0 say S5
  skip
  xh=xh+1
  if mod(xh,(22-定位)/2)=0
    exit
  endif
  enddo
 enddo
close all
delete file temp.dbf
set device to screen
?''
set print off
set print to
clear
retu
�