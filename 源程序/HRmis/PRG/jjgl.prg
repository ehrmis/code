*************************
* jjgl.prg(Build20090117)
*************************
close all
set safety off
ND1 = year(date())
YF1 = month(date())
clear
do srny.spx
ly=nd
ly1=str(val(nd)-1,4)
gzsrk=sblj+'zr1'+ly
if  not file(gzsrk+'.DBF') 
    MESSAGEBOX('请先进行'+ly+'年工资总额处理！！！','提示')
    retu
endif
wait window '正在数据处理!请稍候......' nowait
close all
sele 2
use &sblj.zr&ly
inde on rybh to &sblj.zr&ly
sele 1
use &sblj.zr&ly1
inde on rybh to &sblj.zr&ly1
upda on rybh from B repl cjdm with b.cjdm , cjmc with b.cjmc , bmbh with b.bmbh , bmmc with b.bmmc
copy to &sblj.zr11&ly1
copy to &sblj.zr11temp
close all
*********与上年同期比较工资总额************
  use &sblj.zr11&ly1
  replace MOU with 0 , HJ with 0 all
  for IJB = 1 to yf1
    CMJB = iif(IJB>9,str(IJB,2),'0'+str(IJB,1))
    if  not file('GZ'+ly+CMJB+'.DBF')
      exit
    endif
    replace HJ with round(HJ+J&cmjb,0),MOU with val(CMJB) all
    repl PJ with round(HJ/mou,0) , A with 1 for mou>0
  endfor 
  sum hj,mou to hjtem,moutem
  if hjtem=0 or moutem=0 
     use &sblj.zr11temp
     copy to &sblj.zr11&ly1
  endif    
close all
sele 5
use ryzk
index on RYBH to ryzk
select 4
use &sblj.zr11&ly1 excl 
repl A with 1 all
index on rybh to &sblj.zr11&LY1
set relation to rybh into 5
repl cjdm with e.cjdm,m03 with val(e.zw) for rybh=e.rybh
repl cjdm with '88',cjmc with '管理技术骨干' for m03=7
repl cjdm with '99',cjmc with '科级干部' for m03=5 or m03=6
delete for jjhj=0
pack
index on CJDM to temp
total on CJDM to &sblj.zr11 
select 3
use &sblj.zr&LY
copy to temp
use temp
repl hj with jjhj,PJ with round(HJ/mou,0) , A with 1 for mou>0
set relation to rybh into 5
repl cjdm with e.cjdm,m03 with val(e.zw) for rybh=e.rybh
repl cjdm with '88',cjmc with '管理技术骨干' for m03=7
repl cjdm with '99',cjmc with '科级干部' for m03=5 or m03=6
index on CJDM to &sblj.zr&LY
total on CJDM to &sblj.zr1 
select 2
use &sblj.zr11
replace PJ with round(HJ/A,0) all
sum to LJ1,RS1 HJ,A
index on CJDM to &sblj.zr11
select 1
use &sblj.zr1 excl
delete for jjhj=0
pack
sum jjHJ,A to LJ,RS
set relation to cjdm into 2
replace PJ with round(jjHJ/A,0) , m01 with round(PJ-b.PJ,0) all
replace m02 with round(m01/b.PJ*100,1) all
appe blank
go bott
repl cjmc with '全厂奖金累计',hj with LJ,pj with round(lj/rs,2),m01 with round(LJ/RS,1)-round(LJ1/RS1,1),m02 with (round(LJ/RS,1)-round(LJ1/RS1,1))/round(LJ1/RS1,1)*100 
copy to jjtemp fiel cjmc,hj,pj,m01,m02
use 奖金统计 excl
zap
appe from jjtemp
repl 车间 with cjmc,绩效工资 with hj,人均 with pj,增资额 with m01,增幅 with m02 all
select 4
index on RYBH to zr11temp
select 3
set relation to RYBH into 4 
  clear
  for zdz=1 to fcount()
  zd=fiel(zdz)
  zdz1=(fsize(zd)+fsize(zd))*2.4
  if uppe(allt(zd))='RYXM' 
     exit
  endif
  endfor 
  repl m01 with d.hj ,m02 with d.pj,m03 with round((PJ-m02)/m02*100,1) for d.pj>0  
  copy to  jjtemp fiel cjmc,bmmc,ryxm,hj,m01,pj,m02,m03
use 奖金增幅 excl
zap
appe from jjtemp
repl 车间 with cjmc,部门 with bmmc,绩效工资 with hj,平均 with pj,平均增资额 with pj-m02,增幅 with m03 all
close all
use gzjj excl
zap
appe from &gzsrk for jjhj>0 and mou>0
sele 3
use &sblj.zr&LY
inde on rybh to zr&LY
sele 2
use ryzk
inde on rybh to ryzk
sele 1
use gzjj
inde on rybh to gzjj
upda on rybh from C repl ylbx with c.ylbx,ybx with c.ybx,sybx with c.sybx,gjj with c.gjj,sds with c.sds,sfje with c.sfje 
repl pj with round(hj/mou,0),jfjs with round(jjhj/mou,0),a with 1 all
inde on bmbh to gzjj
tota on bmbh to bmjj
use bmjj
sum hj,jjhj,a,ylbx,ybx,sybx,gjj,sds,sfje to gzsr,jjsr,rs,yl,yb,sy,zf,ss,sf 
repl pj with round(hj/a,0),jfjs with round(jjhj/a,0) all
appe blan
go bott
repl cjmc with '　合  计    ',hj with gzsr,jjhj with jjsr,;
pj with round(hj/rs,0),jfjs with round(jjhj/rs,0),ylbx with yl,ybx with yb,sybx with sy,gjj with zf,sds with ss,sfje with sf
***********************************************************
use gzjj
inde on cjdm to gzjj
tota on cjdm to cjjj
use cjjj
sum hj,jjhj,a,ylbx,ybx,sybx,gjj,sds,sfje to gzsr,jjsr,rs,yl,yb,sy,zf,ss,sf 
repl pj with round(hj/a,0),jfjs with round(jjhj/a,0) all
appe blan
go bott
repl cjmc with '　合  计    ',hj with gzsr,jjhj with jjsr,;
pj with round(hj/rs,0),jfjs with round(jjhj/rs,0),ylbx with yl,ybx with yb,sybx with sy,gjj with zf,sds with ss,sfje with sf
***********************************************************
use gzjj
set relation to rybh into 2
repl a with 0 for val(b.zw)>4 and val(b.zw)<7
copy to kjjj for a=0
use kjjj
set relation to rybh into 2
repl cjmc with b.zw1,a with 1 all
sum hj,jjhj,a,ylbx,ybx,sybx,gjj,sds,sfje to gzsr,jjsr,rs,yl,yb,sy,zf,ss,sf 
appe blan
go bott
repl cjmc with '　合  计    ',hj with gzsr,jjhj with jjsr,;
pj with round(hj/rs,0),jfjs with round(jjhj/rs,0),ylbx with yl,ybx with yb,sybx with sy,gjj with zf,sds with ss,sfje with sf
sort on hj to 科级奖金
***********************************************************
close all
select 1
use jj excl
pack
replace 批回工资 with 当月 for 类别='工资' and 批回否=.T.
replace 批回单项奖 with 当月 for !'工资'$类别 and 批回否=.T.
replace 当月 with 0-当月,发出工资 with 当月 for 类别='工资' and 批回否=.F. and  当月>0
replace 当月 with 0-当月,发出单项奖 with 当月 for !'工资'$类别 and 批回否=.F. and  当月>0
inde on cjdm+bmbh to jj
repl 发出总计 with 0 all
tota on cjdm+bmbh to cjhz field 发出单项奖,发出工资,当月
use cjhz excl
delete for val(cjdm)=0
pack
repl 发出总计 with 当月 all
sum 发出单项奖,发出工资,发出总计 to dxj,zhj,zj
appe blan
go bott
repl 车间名称 with '　合  计    ',发出单项奖 with dxj,发出工资 with zhj,发出总计 with zj 
use 单位汇总 excl
zap
appe from cjhz
use jj
inde on allt(月份) to jj
tota on allt(月份) to dyhj field 当月,批回工资,批回单项奖,发出工资,发出单项奖
select 2
use dyhj
repl 当月发出合 with 发出工资+发出单项奖 all
repl 当月批回合 with 批回工资+批回单项奖 all
go top
do while !eof()
select 1
repl 当月批回工资 with b.批回工资 for allt(月份)=allt(b.月份)
repl 当月批回单项奖 with b.批回单项奖 for allt(月份)=allt(b.月份)
repl 当月发出工资 with b.发出工资 for allt(月份)=allt(b.月份)
repl 当月发出单项奖 with b.发出单项奖 for allt(月份)=allt(b.月份)
repl 当月发出合计 with b.当月发出合 for allt(月份)=allt(b.月份)
repl 当月批回合计 with b.当月批回合 for allt(月份)=allt(b.月份)
select 2  
skip
enddo
close all
use jj
calculate for 批回否=.T. and 类别='工资' to PHZHJ sum(当月)
calculate for 批回否=.F. and 类别='工资' to FCZHJ sum(当月)
calculate for 批回否=.T. and !'工资'$类别 to PHDXJ sum(当月)
calculate for 批回否=.F. and !'工资'$类别 to FCDXJ sum(当月)
calculate for 批回否=.T. to PHZJ sum(当月)
calculate for 批回否=.F. to FCZJ sum(当月)
calculate for 名称='年终奖' to NZJ sum(当月)
calculate for 名称='季度奖' to JDJ sum(当月)
repl 累计批回工资 with PHZHJ,累计批回单项奖 with PHDXJ,累计发出工资 with FCZHJ,累计发出单项奖 with FCDXJ,批回总计 with PHZJ,发出总计 with FCZJ all
repl 累计 with 0 all
inde on allt(名称) to jj
repl 累计 with 0 all
tota on allt(名称) to lj1 field 当月,累计
use lj1
repl 累计 with 当月 all
use jj
inde on allt(cjdm)+allt(bmbh)+allt(名称) to jj
repl 累计 with 0 all
tota on allt(cjdm)+allt(bmbh)+allt(名称)+allt(月份) to lj field 当月,累计
close all
sele 2
use lj
repl 累计 with 当月 all
inde on allt(cjdm)+allt(bmbh)+allt(名称)+allt(月份) to lj
sele 1
use jj
set relation to allt(cjdm)+allt(bmbh)+allt(名称)+allt(月份) into 2
go top
do while !eof()
repl 累计 with b.当月 
skip
enddo
close all
IF !DBUSED('JJGL') 
      OPEN DATABASE JJGL 
 ENDIF   
  IF !DBUSED('DMK') 
      OPEN DATABASE DMK 
 ENDIF   
 use jj 
 sort on 月份,cjdm,bmbh to 奖金明细 field 月份,车间名称,部门名称,名称,类别,批回否,当月,累计
 sort on cjdm,bmbh,名称,类别,月份 to 单位明细 field 车间名称,部门名称,名称,类别,当月,累计
 inde on allt(cjdm)+allt(bmbh)+allt(月份)+allt(名称)+allt(类别) to jj
 sort on cjdm,bmbh,月份,名称,类别 to str(ND1,4)+'年薪酬'
 use lj1 
 sort on 名称,类别 to 奖金累计 field 名称,类别,批回否,累计
 use lj 
 sort on cjdm,bmbh,名称 to 各单位薪酬 field 车间名称,部门名称,名称,累计
 close all
 delete file lj.dbf
 delete file lj1.dbf
  IF !DBUSED('JJGL') 
      OPEN DATABASE JJGL 
 ENDIF   
  IF !DBUSED('DMK') 
      OPEN DATABASE DMK 
 ENDIF   
do form jj
