***************************
* .\gzze.PRG (Buld201405)
***************************
close all
dhsr=''
set safety off
on key label F1 do grcx
**************双位月数自动处理*******************
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
起始月 = 起始月份
*************************************************
use gzze
COPY TO zr&LY stru
USE zr&LY
if file('RY'+LY+'01.DBF')
***************************保留年初人员********************
 yn = MESSAGEBOX("需要保留一月份在册职工（含现已调出职工）工资总额吗？",4+32,"提示")
IF yn = 6
 append from ry&LY.01 
ELSE
 append from ryzk
ENDIF 
else
 append from ryzk
endif
************************通用功能：重号检查***********
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
   brow for a=0 titl '发现'+allt(str(aaa,4))+'人重号！请修改正确或按[ Ctrl+T ]删除多余记录'
   pack
   delete file zr&LY..dbf
   sort on bmbh,rybh to zr&LY
endif
close all
select 1
use zr&LY
inde on rybh to zr&LY
M1 = 1
人数=0
repl gzhj with 0,sfje with 0,ylbx with 0,qynj with 0,sybx with 0,ybx with 0,gjj with 0,sds with 0 all
***********语法严谨性：循环递加初始化*****不清零就会越加越多重复相加多少遍了************
go top
do while M1<=val(yf)
  MGZ = iif(M1>9,str(M1,2),'0'+str(M1,1))
  wait wind  '正在填写'+allt(str(val(mgz)))+'月职工工资总额...'  nowa
  if  not file('GZ'+LY+MGZ+'.dbf') 
  *or not file('TY'+LY+MGZ+'.dbf')
    ***************ERP****************************
    MESSAGEBOX("请导入ERP系统中"+LY+"年"+allt(str(val(MGZ)))+"月工资奖金数据",48,"提示")
    retu
  endif  
***********工伤产假人员本月无工资但缴ylbx**********
    select 2
    use gzk
    copy to temp stru
    use temp
    appe from gz&LY.&Mgz
    repl a with 1 all
    sum a to rs
    人数=人数+rs 
    if file('TY'+LY+Mgz+'.dbf')
       appe from TY&LY.&Mgz
    endif
    index on RYBH to temp
    select 1
    upda on rybh from B replace sfje with sfje+b.sfje,M&Mgz with b.hj-b.jang,gzhj with gzhj+M&Mgz;
ylbx with ylbx+b.ylbx,qynj with qynj+b.qynj,ybx with ybx+b.ybx,sybx with sybx+b.sybx,gjj with gjj+b.gjj,sds with sds+b.sds 
    M1 = M1+1
enddo
人数=round(人数/(M1-1),0)
close all
M1 = 1
select 1
use zr&LY
inde on rybh to zr&LY
repl jjhj with 0 all
go top
do while M1<=val(yf)
  MGZ = iif(M1>9,str(M1,2),'0'+str(M1,1))
  wait wind '正在填写'+allt(str(M1,2))+'月份单项奖... ....' nowa
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
index on bmbh+rybh to zr&LY
********************生成工资总额(HJ)与月平均(PJ)**************
repl gzhj with m01+m02+m03+m04+m05+m06+m07+m08+m09+m10+m11+m12,jjhj with j01+j02+j03+j04+j05+j06+j07+j08+j09+j10+j11+j12,; 
hj with gzhj+jjhj,PJ with round(HJ/MOU,0) for mou>0
USE gzze1
COPY TO zr1&ly stru
APPEND FROM zr&ly
********生成缴费基数*******************
use zr&LY
DBFF1='zr'
WHNY=LY+'年工资总额'
do Qdir
use zr&LY
pf='d:\'+LY
copy to &pf.年工资总额 type xl5 
 =messagebox("已成功导出在D：\工资总额电子表！",48,"恭喜")
  yn = MESSAGEBOX("需要使用电子表编辑打印报表吗？",4+32,"提示")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           pf1="&pf"
           myexcel.workbooks.open(pf1+"年工资总额.xls")
           myexcel.visible=.t.
           myexcel.caption="使用电子表编辑打印报表" 
           release oleapp
           WAIT CLEAR
ENDIF     
close all
do dwgzze
