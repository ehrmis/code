***********
* 岗位统计
***********
YF1=month(date())
ND1=year(date())
do srnd.spx
NY = ND+YF
NYR=right(NY,4)
NYR = Right(NY,4)
YFs = iif(month(date())>9,str(month(date())-1,2),'0'+str(month(date())-1,1))
NY1 = str(year(date()),4)+YFs
if month(date())=1
   NY1 = str(year(date())-1,4)+'12'
   YFs = '12'
endi
if month(date())=10
   NY1 = str(year(date()),4)+'09'
endi
if !file('gw'+ny+'.dbf')
=MESSAGEBOX('请先审核维护本月工种岗位！', 16,"提示")
 return
endif
clear
zt="font '宋体',11"
zt1="font '宋体',10"
set device to printer
set printer on
YN = MESSAGEBOX('需要打印岗位等级审核表吗？！',4+32,'提示')
if yn=6
t1='┏━━┯━━━━┯━━━┯━━━┯━━━━┯━━━━┯━━━━┯━━┓'
t2='┃序号│岗位分类│ 级别 │ 档级 │岗位等级│浮后岗资│岗资级别│人数┃'
t3='┠──┼────┼───┼───┼────┼────┼────┼──┨'
t4='┗━━┷━━━━┷━━━┷━━━┷━━━━┷━━━━┷━━━━┷━━┛'
use 岗位统计
sum 人数 for '操作'$gwfl1 to cz
sum 人数 for '技术'$gwfl1 to js
sum 人数 for '管理'$gwfl1 to glry
sum 人数 to rs
sum val(gwdm)*人数 to gs
go top
xh=1
do while  not eof()
      @ 0 , 0 say space(5)+'&MC111'+'岗位等级审核表(表一)' font  '宋体',11
      @ 2 , 0 say T1 &zt
      @ 3 , 0 say T2 &zt
      @ 4 , 0 say T3 &zt
  do while  not eof()
  @ prow()+1 , 0 say '┃' &zt
  @ prow() , pcol() say str(xh,4) &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwfl1 &zt
  @ prow() , pcol() say '│'&zt
  @ prow() , pcol() say gwjb1 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwdc1 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say 岗等 &zt
  @ prow() , pcol() say '│  ' &zt
  @ prow() , pcol() say gwgz &zt
  @ prow() , pcol() say '  │ ' &zt
  @ prow() , pcol() say 岗级 &zt
  @ prow() , pcol() say ' │' &zt
  @ prow() , pcol() say 人数 picture '@Z 9999' &zt
  @ prow() , pcol() say '┃' &zt
  skip
  xh=xh+1
  if mod(xh,26)=0
     @ prow()+1 , 0 say t4 &zt
     exit
  else
     @ prow()+1 , 0 say T3 &zt
  endif
  enddo
enddo
  @ prow()+1 , 0 say '┃' &zt
  @ prow() , pcol() say '合计│' &zt
  @ prow() , pcol() say '操作岗位:'+str(cz,3)+'人 技术岗位:'+str(js,3)+'人 管理岗位:'+str(glry,2) &zt
  @ prow() , pcol() say '人 平均' &zt
  @ prow() , pcol() say str(round(gs/rs,1),4,1) &zt
  @ prow() , pcol() say '岗│' &zt
  @ prow() , pcol() say rs picture '@Z 9999' &zt
  @ prow() , pcol() say '┃' &zt
  @ prow()+1 , 0 say t4 &zt
  @ prow()+1 , 10 say '填报日期:'+alltrim(str(year(date()),4))+'年'+alltrim(str(month(date());
,2))+'月'+alltrim(str(day(date()),2))+'日' &zt1
@ prow() , pcol() say '                                填报人员:&ZB111' &zt1
set printer to lpt1
endif
?
YN = MESSAGEBOX('需要打印持证操作人员岗位分布表吗？！',4+32,'提示')
if yn=6
t1='┏━━┯━━━━┯━━━┯━━━┯━━━━┯━━━━┯━━━━┯━━┯━━┓'
t2='┃序号│岗位分类│ 级别 │ 档级 │岗位等级│浮后岗资│岗资级别│班长│人数┃'
t3='┠──┼────┼───┼───┼────┼────┼────┼──┼──┨'
t4='┗━━┷━━━━┷━━━┷━━━┷━━━━┷━━━━┷━━━━┷━━┷━━┛'
use 操作分类
sum 人数 to rs
sum val(gwdm)*人数 to gs
sum 人数 for '操作'$gwfl1 to cz
sum 人数 for '技术'$gwfl1 to js
sum 人数 for !empty(bzf) to bz
go top
xh=1
do while  not eof()
      @ 0 , 0 say space(5)+'&MC111'+'岗位等级统计表（表二）' font '黑体',20
      @ 2 , 0 say '                              ( 持证操作人员岗位分布 )' &zt1
      @ 3 , 0 say T1 &zt
      @ 4 , 0 say T2 &zt
      @ 5 , 0 say T3 &zt
do while  not eof()
  @ prow()+1 , 0 say '┃' &zt
  @ prow() , pcol() say str(xh,4) &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwfl1 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwjb1 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwdc1 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say 岗等 &zt
  @ prow() , pcol() say '│  ' &zt
  @ prow() , pcol() say gwgz &zt
  @ prow() , pcol() say '  │ ' &zt
  @ prow() , pcol() say 岗级 &zt
  @ prow() , pcol() say ' │' &zt
if !empty(bzf)
   @ prow() , pcol() say ' √ │' &zt
else
   @ prow() , pcol() say '    │' &zt
endif
  @ prow() , pcol() say 人数 picture '@Z 9999' &zt
  @ prow() , pcol() say '┃' &zt
  skip
  xh=xh+1
  if mod(xh,26)=0
     @ prow()+1 , 0 say t4 &zt
     exit
  else
     @ prow()+1 , 0 say T3 &zt
  endif
  enddo
enddo
  @ prow()+1 , 0 say '┃' &zt
  @ prow() , pcol() say '合计' &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say '操作岗位:'+str(cz,3)+'人 技术岗位:'+str(js,3) &zt
  @ prow() , pcol() say '人 平均:'+str(round(gs/rs,1),4,1)+'  岗' &zt
  @ prow() , pcol() say ' │注册班长: '+str(bz,2)+'人│' &zt
  @ prow() , pcol() say rs picture '@Z 9999' &zt
  @ prow() , pcol() say '┃' &zt
@ prow()+1 , 0 say t4 &zt
@ prow()+1 , 10 say '填报日期:'+alltrim(str(year(date()),4))+'年'+alltrim(str(month(date());
,2))+'月'+alltrim(str(day(date()),2))+'日' &zt1
@ prow() , pcol() say '                                填报人员:&ZB111' &zt1
set printer to lpt1
endif
?
YN = MESSAGEBOX('需要打印无证人员岗位分布表吗？！',4+32,'提示')
if yn=6
t1='┏━━┯━━━━┯━━━┯━━━┯━━━━┯━━━┯━━┯━━┓'
t2='┃序号│岗位分类│ 级别 │ 档级 │浮后岗资│ 岗级 │班长│人数┃'
t3='┠──┼────┼───┼───┼────┼───┼──┼──┨'
t4='┗━━┷━━━━┷━━━┷━━━┷━━━━┷━━━┷━━┷━━┛'
use 无证分类
sum 人数 to rs
sum val(gwdm)*人数 to gs
go top
xh=1
      @ 0 , 0 say space(1)+'&MC111'+'岗位等级统计表（表三）' font '黑体',20
      @ 2 , 0 say '                         ( 无证人员岗位分布 )' &zt1
      @ 3 , 0 say T1 &zt
      @ 4 , 0 say T2 &zt
      @ 5 , 0 say T3 &zt
do while  not eof()
  @ prow()+1 , 0 say '┃' &zt
  @ prow() , pcol() say str(xh,4) &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwfl1 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwjb1 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwdc1 &zt
  @ prow() , pcol() say '│  ' &zt
  @ prow() , pcol() say gwgz &zt
  @ prow() , pcol() say '  │' &zt
  @ prow() , pcol() say 岗级+'│' &zt
if !empty(bzf)
   @ prow() , pcol() say ' √ │' &zt
else
   @ prow() , pcol() say '    │' &zt
endif
  @ prow() , pcol() say 人数 picture '@Z 9999' &zt
  @ prow() , pcol() say '┃' &zt
  @ prow()+1 , 0 say t3 &zt
  skip
  xh=xh+1
enddo
  @ prow()+1 , 0 say '┃' &zt
  @ prow() , pcol() say '合计' &zt
  @ prow() , pcol() say '│        │      │      │        │      │    │' &zt
  @ prow() , pcol() say rs picture '@Z 9999' &zt
  @ prow() , pcol() say '┃' &zt
@ prow()+1 , 0 say t4 &zt
@ prow()+1 , 10 say '填报日期:'+alltrim(str(year(date()),4))+'年'+alltrim(str(month(date());
,2))+'月'+alltrim(str(day(date()),2))+'日' &zt1
@ prow() , pcol() say '                                填报人员:&ZB111' &zt1
set printer to lpt1
endif
?
YN = MESSAGEBOX('需要打印管理人员岗位分布表吗？！',4+32,'提示')
if yn=6
t1='┏━━┯━━━━┯━━━┯━━━┯━━━━┯━━━┯━━━━┯━━━━┯━━┓'
t2='┃序号│岗位分类│ 级别 │ 档级 │浮后岗资│ 岗级 │行政职务│技术职称│人数┃'
t3='┠──┼────┼───┼───┼────┼───┼────┼────┼──┨'
t4='┗━━┷━━━━┷━━━┷━━━┷━━━━┷━━━┷━━━━┷━━━━┷━━┛'
use 管理分类
sum 人数 for '技术'$gwfl1 to js
sum 人数 for '管理'$gwfl1 to glry
sum 人数 to rs
sum val(gwdm)*人数 to gs
go top
xh=1
      @ 0 , 0 say space(7)+'&MC111'+'岗位等级统计表（表四）' font '黑体',20
      @ 2 , 0 say '                                      ( 管理人员岗位分布 )' &zt1
      @ 3 , 0 say T1 &zt
      @ 4 , 0 say T2 &zt
      @ 5 , 0 say T3 &zt
do while  not eof()
  @ prow()+1 , 0 say '┃' &zt
  @ prow() , pcol() say str(xh,4) &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwfl1 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwjb1 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gwdc1 &zt
  @ prow() , pcol() say '│  ' &zt
  @ prow() , pcol() say gwgz &zt
  @ prow() , pcol() say '  │' &zt
  @ prow() , pcol() say 岗级 &zt
  @ prow() , pcol() say '│' &zt
if '技'$gwfl1
  @ prow() , pcol() say '        │' &zt
  @ prow() , pcol() say 岗等 &zt
  @ prow() , pcol() say '│' &zt
else
  @ prow() , pcol() say 岗等 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say '        │' &zt
endif
  @ prow() , pcol() say 人数 picture '@Z 9999' &zt
  @ prow() , pcol() say '┃' &zt
  @ prow()+1 , 0 say t3 &zt
  skip
  xh=xh+1
enddo
  @ prow()+1 , 0 say '┃' &zt
  @ prow() , pcol() say '合计' &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say '管理岗位:'+str(glry,3)+'人 技术岗位:'+str(js,3) &zt
  @ prow() , pcol() say '人 平均:'+str(round(gs/rs,1),4,1)+'岗' &zt
  @ prow() , pcol() say '  共:' &zt
  @ prow() , pcol() say rs picture '@Z 9999' &zt
  @ prow() , pcol() say '人'+space(16)+'┃' &zt
@ prow()+1 , 0 say t4 &zt
@ prow()+1 , 10 say '填报日期:'+alltrim(str(year(date()),4))+'年'+alltrim(str(month(date());
,2))+'月'+alltrim(str(day(date()),2))+'日' &zt1
@ prow() , pcol() say '                                填报人员:&ZB111' &zt1
set printer to lpt1
endif
?
YN = MESSAGEBOX('需要打印工种统计表吗？！',4+32,'提示')
if yn=6
t1='┏━━┯━━━━┯━━━━━━┯━━━━━━━━┯━━━┯━━━━┓'
t2='┃序号│ 代  码 │  工    种  │ 岗          位 │ 人数 │岗位等级┃'
t3='┠──┼────┼──────┼────────┼───┼────┨'
t4='┗━━┷━━━━┷━━━━━━┷━━━━━━━━┷━━━┷━━━━┛'
use 工种统计
sum a to rs
inde on 岗等 to 工种统计
go top
xh=1
M = 1
  do while  not eof()
      @ 0 , 0 say space(1)+'&MC111'+str(year(date()),4)+'年'+allt(str(month(date()),2))+'月份工种统计表' font '黑体',20
      @ 2 , 45 say '第'+str(M,2)+'页' &zt1
      @ 3 , 0 say T1 &zt
      @ 4 , 0 say T2 &zt
      @ 5 , 0 say T3 &zt
  do while !eof()
  @ prow()+1 , 0 say '┃' &zt
  @ prow() , pcol() say str(xh,4) &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gz &zt
  @ prow() , pcol() say ' │' &zt
  @ prow() , pcol() say gz1 &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say gw &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say a picture '@Z 999999' &zt
  @ prow() , pcol() say '│' &zt
  if gwfl1='操作岗位'
     @ prow() , pcol() say space(1)+gwjb1+' ┃' &zt
  else
     @ prow() , pcol() say 岗等+'┃' &zt
  endi
  skip
  xh=xh+1
  if mod(xh,26)=0
     @ prow()+1 , 0 say t4 &zt
     m=m+1
     exit
  else
     @ prow()+1 , 0 say t3 &zt
  endif
  endd
enddo
  @ prow()+1 , 0 say '┃' &zt
  @ prow() , pcol() say '合计' &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say space(8) &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say space(12) &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say space(16) &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say str(rs,6) &zt
  @ prow() , pcol() say '│' &zt
  @ prow() , pcol() say space(8)+'┃' &zt
@ prow()+1 , 0 say t4 &zt
?
@ prow()+1 , 0 say '填报日期:'+alltrim(str(year(date()),4))+'年'+alltrim(str(month(date());
,2))+'月'+alltrim(str(day(date()),2))+'日' &zt1
@ prow() , pcol() say '          填报人员:&ZB111' &zt1
?''
set printer to lpt1
endif
use
??chr(13)+chr(10)
set device to screen
set printer off
set console off
set confirm off
set talk off
retu
