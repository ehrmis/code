***************************
*ERPkq.PRG(Build2005.08.01)
***************************
close all
on key label F1 do grcx
YF1=month(date())
ND1=year(date())
do srny.spx
clear
NY = ND+YF
close all
***********************Ａ.初始化*****************
IF !file("GZ"+NY+".txt")
    =MESSAGEBOX("ERP工资文件GZ"+ny+".TXT未下载...",0+48,"错误")
    retu
ELSE   
use erpgz excl
zap
APPEND FROM gz&ny..txt DELIMITED WITH CHAR TAB
REPLACE erprybh WITH ALLTRIM(STR(d)) all
REPLACE erprybh WITH '00'+ALLTRIM(erprybh) FOR LEN(ALLTRIM(erprybh))=6
REPLACE erprybh WITH '0'+ALLTRIM(erprybh) FOR LEN(ALLTRIM(erprybh))=7
ENDIF
close all
sele 2
use ryzk
inde on erprybh to erpgz
sele 1
if !file('kq'+ny+'.dbf')
   use kqk excl
   zap
   appe from erpgz
   set relation to erprybh into 2
repl rybh with b.rybh,cjmc with b.cjmc,cjdm with b.cjdm,bmbh WITH b.bmbh,bmmc WITH b.bmmc all 
   sort on bmbh,rybh to kq&ny
use kq&ny excl  
repl 中班 with zbgr,夜班 with ybgr,节日加班 with jrjb,病假 with bj,事假 with sj,产假 with cj,工伤 with gs,探亲假 with tqj,;
婚丧假 with hsj,拘留 with jl,旷工 with kg,年休假 with gj all
go top
brow noedit part 40 titl '请认真审核维护“ERP工资数据导入情况”'
nn=MESSAGEBOX("数据导入成功...",0+64,"信息")
else
nn=MESSAGEBOX(nd+"年"+allt(str(val(yf)))+"月份考勤已存在，数据导入失败！",0+64,"信息")
*****************安全性******************
endif
clear
close all
delete file erp*.*
