close all
set safety off
use 党员 excl
zap
appe from ryzk for '党'$zzmm1
*repl qt1 with bzgz-85-ylbx-ybx-sybx-gjj all
repl qt2 with round(bzgz/100,0) all
*go top
*do while !eof()
*do case
*case qt1<=400
 *    repl qt2 with round(qt1*0.5/100,0)
*case qt1>400 and qt1<=600
 *    repl qt2 with round(qt1*1/100,0)
*case qt1>600 and qt1<=800
 *    repl qt2 with round(qt1*1.5/100,0)
*case qt1>800
 *    repl qt2 with round(qt1*2/100,0)
*endcase
*if qt2>10 and qt2<15
 *  repl qt2 with 10
*endi
*skip
*enddo
repl qt2 with 10 for qt2>10 and qt2<15 and empty(zw1)
go top
brow fiel cjmc:h='车间',bmmc:h='部门(科室)',ryxm:h='姓名',qt2:h='月缴党费'
pd='Y'
@ 10,30 say '需要打印吗(Y/N)?' get pd
read
if upper(pd)='Y'
   list to prin fiel cjmc,bmmc,ryxm,qt2
endi
retu
                                  