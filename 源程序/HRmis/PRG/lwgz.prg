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
***************************保留年初人员********************
 yn = MESSAGEBOX("需要保留一月份在册职工（含现已调出职工）工资总额吗？",4+32,"提示")
IF yn = 6
 append from ry&LY.01 
ELSE
 append from ryzk for '劳务派遣工'$ygxz1
ENDIF 
else
 append from ryzk for '劳务派遣工'$ygxz1
endif
close all
************************通用功能：重号检查***********
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
   brow for a=0 titl '发现'+allt(str(aaa,4))+'人重号！请修改正确或按[ Ctrl+T ]删除多余记录'
   pack
   delete file lw&LY..dbf
   sort on bmbh,rybh to lw&LY
endif
close all
select 1
use lw&LY
inde on rybh to lw&LY
M1 = 1
人数=0
repl gzhj with 0,实发金额 with 0,绩效工资 with 0,生产嘉奖 with 0,单项奖 with 0 all
***********语法严谨性：循环递加初始化*****不清零就会越加越多重复相加多少遍了************
go top
do while M1<=val(yf)
  MGZ = iif(M1>9,str(M1,2),'0'+str(M1,1))
  wait wind  '正在填写'+allt(str(val(mgz)))+'月劳务派遣工工资总额...'  nowa
  if  not file('LW'+LY+MGZ+'.dbf')   
    MESSAGEBOX("未发现"+LY+"年"+allt(str(val(MGZ)))+"月工资数据",48,"提示")
    retu
  endif  
    select 2
    use lwpqgzk
    copy to temp stru
    use temp
    appe from lw&LY.&Mgz
    repl a with 1 all
    sum a to rs
    人数=人数+rs 
    index on RYBH to temp
    select 1
    upda on rybh from B replace 实发金额 with 实发金额+b.实发金额,M&Mgz with b.实发金额-b.绩效工资,gzhj with gzhj+M&Mgz;
绩效工资 with 绩效工资+b.绩效工资,生产嘉奖 with 生产嘉奖+b.生产嘉奖,单项奖 with 单项奖+b.单项奖
   M1 = M1+1
enddo
人数=round(人数/(M1-1),0)
close all
M1 = 1
select 1
use lw&LY
inde on rybh to lw&LY
repl jjhj with 0 all
go top
do while M1<=val(yf)
  MGZ = iif(M1>9,str(M1,2),'0'+str(M1,1))
  wait wind '正在填写'+allt(str(M1,2))+'月份绩效工资... ....' nowa
  if  not file('lw'+LY+MGZ+'.dbf')
    exit
  else
    select 2
    use lw&LY.&Mgz
    index on RYBH to lw&LY.&Mgz
    select 1
    upda on rybh from B replace J&Mgz with b.绩效工资,jjhj with jjhj+J&Mgz
    M1 = M1+1
  endif
enddo
wait wind  '正在修正实际缴费月数...'  nowa
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
********************生成工资总额(HJ)与月平均(PJ)**************
repl gzhj with m01+m02+m03+m04+m05+m06+m07+m08+m09+m10+m11+m12,jjhj with j01+j02+j03+j04+j05+j06+j07+j08+j09+j10+j11+j12,; 
hj with gzhj+jjhj,PJ with round(HJ/MOU,0) for mou>0
brow
use LW&LY
DBFF1='lw'
WHNY=LY+'年劳务费'
do Qdir
use lw&LY
pf='d:\'+LY
copy to &pf.年劳务费 type xl5 
 =messagebox("已成功导出在D：\劳务费电子表！",48,"恭喜")
  yn = MESSAGEBOX("需要使用电子表编辑打印报表吗？",4+32,"提示")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           pf1="&pf"
           myexcel.workbooks.open(pf1+"年劳务费.xls")
           myexcel.visible=.t.
           myexcel.caption="使用电子表编辑打印报表" 
           release oleapp
           WAIT CLEAR
ENDIF     
close all

