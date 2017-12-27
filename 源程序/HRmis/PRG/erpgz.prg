***************************
*ERPgz.PRG(Build2011.3.1)
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
IF !file("GZ"+NY+".txt") or !file("BQ"+NY+".txt")
    =MESSAGEBOX("ERP工资文件GZ"+ny+".TXT未下载...",0+48,"错误")
    retu
else 
use erpgz excl
zap
APPEND FROM gz&ny..txt DELIMITED WITH CHAR TAB
APPEND FROM bq&ny..txt DELIMITED WITH CHAR TAB
REPLACE erprybh WITH ALLTRIM(STR(d)) all
REPLACE erprybh WITH '00'+ALLTRIM(erprybh) FOR LEN(ALLTRIM(erprybh))=6
REPLACE erprybh WITH '0'+ALLTRIM(erprybh) FOR LEN(ALLTRIM(erprybh))=7
use gzk 
copy to temp stru
use temp
appe from erpgz
ENDIF
close all
*sele 3
*use kq&ny
*inde on erprybh to kq&ny
sele 2
use ryzk
copy to temp1 for ygxz='01'
use temp1
inde on erprybh to temp1
sele 1
USE temp excl
delete for val(erprybh)=0
pack
inde on erprybh to temp
upda on erprybh from B repl  rybh with b.rybh,dwdm with b.dwdm,dwmc with b.dwmc,cjmc with b.cjmc,cjdm with b.cjdm,bmbh WITH b.bmbh,bmmc WITH b.bmmc,gl WITH b.gl,zw WITH b.zw,zw1 with b.zw1
*upda on erprybh from C repl  jang2 with c.jang2
**************导入单项奖*************************
repl 车间 with cjmc,部门 with bmmc,姓名 with ryxm,绩效工资 with jang-jang2,单项奖 with jang2,工资总额 with hj,实发金额 with sfje all
copy to gztemp for val(rybh)>0
use gztemp
sort on bmbh,rybh to gz&ny
use gz&ny
repl a with 1 all
repl hj with ylbx+qynj+ybx+sybx+gjj+sds+fzsd+sfje for hj=0
go top
brow noedit part 40 titl '请认真审核维护“ERP工资数据导入情况”'
nn=MESSAGEBOX("数据导入成功...",0+64,"信息")
use cwcsyk excl
ny2=dm111+right(left(ny,4),2)+right(ny,2)
zap
copy to kg&ny2
USE KG&NY2
append from GZ&NY
if file('ty'+ny+'.dbf')
   append from ty&NY
endif
REPL ALL GJJ1 WITH GJJ,HJ WITH GJJ+GJJ1,ydhj with gjjz
clear
close all
delete file erp*.*
