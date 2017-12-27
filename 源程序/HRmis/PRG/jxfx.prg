*************************
* jxfx.prg(Build20080104)
*************************
close all
set safety off
dhsr=''
ND1 = year(date())
YF1 = month(date())
clear
do srny.spx
ly=nd
ly1=str(val(nd)-1,4)
gzsrk=sblj+'zr'+ly1
if  not file(gzsrk+'.DBF')
  =MESSAGEBOX("请先处理"+LY1+"年工资总额!",48,"提示")
    close all
    return
else
   use &sblj.zr&LY1
   if recc()=0
    =MESSAGEBOX("请先处理"+LY1+"年工资总额!",48,"提示")
    close all
    return
   endif    
endif    
gzsrk=sblj+'zr1'+ly
if  not file(gzsrk+'.DBF') or !file('gz'+LY+'.dbf')
    MESSAGEBOX('请先进行'+ly+'年工资总额处理！！！','提示')
    return
else
   use &sblj.zr1&LY
   if recc()=0
    =MESSAGEBOX("请先处理"+LY+"年工资总额!",48,"提示")
    close all
    return
   endif    
endif    
wait window '正在数据处理!请稍候......' nowait
close all
dhsr=''
set safety off
on key label F1 do grcx
close all
sele 2
use &sblj.zr&LY
copy to temp stru
copy to &sblj.zr11&LY1 stru
copy to &sblj.zr11temp stru
inde on rybh to &sblj.zr&LY
sele 1
use &sblj.zr11&LY1
yn = MESSAGEBOX("处级干部不参加工资总额增长比较吗？",4+32,"提示")
IF yn = 6
appe from &sblj.zr&LY1 for (zw<'03' or zw>'04') and ygxz='01'
else
appe from &sblj.zr&LY1 for ygxz='01'
ENDI
inde on rybh to &sblj.zr11&LY1
upda on rybh from B repl cjdm with b.cjdm , cjmc with b.cjmc , bmbh with b.bmbh , bmmc with b.bmmc, zw with b.zw,;
zw1 with b.zw1 , gwfl with b.gwfl , gwfl1 with b.gwfl1,zcdj with b.zcdj,zcdj1 with b.zcdj1,ygxz with b.ygxz,ygxz1 with b.ygxz1
use &sblj.zr11temp
IF yn = 6
appe from &sblj.zr&LY1 for (zw<'03' or zw>'04') and ygxz='01'
else
appe from &sblj.zr&LY1 for ygxz='01'
ENDI
inde on rybh to &sblj.zr11temp
upda on rybh from B repl cjdm with b.cjdm , cjmc with b.cjmc , bmbh with b.bmbh , bmmc with b.bmmc, zw with b.zw,;
zw1 with b.zw1 , gwfl with b.gwfl , gwfl1 with b.gwfl1,zcdj with b.zcdj,zcdj1 with b.zcdj1,ygxz with b.ygxz,ygxz1 with b.ygxz1
do dhksr 
close all
if dhsr=1 
*********与上年同期比较工资总额************
  use  &sblj.zr11&LY1 &&(上年同期明细)
  replace MOU with 0 , HJ with 0 all
  for IJB = 1 to val(yf)
    CMJB = iif(IJB>9,str(IJB,2),'0'+str(IJB,1))
    if  not file('GZ'+LY+CMJB+'.DBF')
      exit
    endif
    replace HJ with round(HJ+M&cmjb+J&cmjb,0),MOU with val(CMJB) all
    repl PJ with round(HJ/mou,0) , A with 1 for mou>0
  endfor  
endif
close all
sele 5
use ryzk
index on RYBH to ryzk
select 4
if dhsr=1 
  use  &sblj.zr11&LY1 excl 
  **************上年同期明细
  repl A with 1 all
  index on rybh to  &sblj.zr11&LY1
else
  use  &sblj.zr11temp excl
  repl hj with gzhj+jjhj,PJ with round(HJ/mou,0) , A with 1 for mou>0
  index on rybh to  &sblj.zr11temp 
endif
set relation to rybh into 5
repl cjdm with e.cjdm  for rybh=e.rybh
repl cjdm with '99',cjmc with '科级干部' for zw='05' or zw='06'
repl cjdm with '88',cjmc with '处级干部' for zw='03' or zw='04'
index on CJDM to temp
total on CJDM to &sblj.zr11 
*******************上年同期汇总
select 3
use temp 
IF yn = 6
appe from  &sblj.zr&LY for (zw<'03' or zw>'04') and ygxz='01'
else
appe from  &sblj.zr&LY for ygxz='01'
ENDI
set relation to rybh into 5
repl cjdm with e.cjdm  for rybh=e.rybh
repl cjdm with '99',cjmc with '科级干部' for zw='05' or zw='06'
repl cjdm with '88',cjmc with '处级干部' for zw='03' or zw='04'
index on CJDM to &sblj.zr&LY
total on CJDM to &sblj.zr1  &&(今年汇总)
select 2
use &sblj.zr11
replace PJ with round(HJ/A,0) all
sum to LJ1,RS1 HJ,A
index on CJDM to &sblj.zr11
select 1
use &sblj.zr1 excl
IF yn = 6
delete for cjmc = '处级干部'
pack
ENDI
sum HJ,A to LJ,RS
set relation to cjdm into 2
replace PJ with round(HJ/A,0) , m01 with round(PJ-b.PJ,0) all
replace m02 with round(m01/b.PJ*100,1) all
appe blank
go bott
  repl cjmc with '全厂累计总额',hj with LJ,pj with round(lj/rs,2),m01 with round(LJ/RS,1)-round(LJ1/RS1,1),m02 with (round(LJ/RS,1)-round(LJ1/RS1,1))/round(LJ1/RS1,1)*100 
  go top
    browse field CJMC:H = ' 单位名称 ' , HJ :P = '@z 99,999;
,999.99' : 14 :H = '自年初累计(元)' , PJ :P =;
 '@z 999,999' : 15 :H = '平均(元/人)' , m01;
 : 20 :H = '比上年增长(元/人)' , m02;
 : 14 :H = '增幅（%）' noedit titl '一、全厂自年初累计工资总额统计表（单位：元）'
close all
***********************
use temp 
copy to temp1 stru
use temp1
yn = MESSAGEBOX("科级干部不参加工资总额增长比较吗？",4+32,"提示")
IF yn = 6
appe from temp for (zw<'05' or zw>'06') and ygxz='01'
else
appe from temp for ygxz='01'
endif
repl a with 1 all
inde on gwfl to temp1
tota on gwfl to hztemp
use hztemp
repl pj with round(hj/a,0) all
go top
brow fiel gwfl1:h='人员分类',HJ :P = '@z 99,999;
,999.99' : 14 :H = '自年初累计(元)' , PJ :P =;
 '@z 999,999' : 15 :H = '平均(元/人)' ,a:h='人数'  noedit titl '各岗位人员自年初累计工资总额统计表（单位：元）'
 copy to D:\人员分类收入分析 fiel gwfl1,jjhj,gzhj,hj,pj,a type xls
 =messagebox("各岗位人员收入分析数据已导出到“D：\人员分类收入分析”电子表中！！",48,"恭喜") 
use temp
copy to temp1 stru
use temp1
IF yn = 6
appe from temp for (zw<'05' or zw>'06') and ygxz='01'
else
appe from temp for ygxz='01'
endif
repl a with 1 all
inde on zcdj to temp
tota on zcdj to hztemp
use hztemp
repl pj with round(hj/a,0) all
go top
brow fiel zcdj1:h='职称（职业资格）等级',HJ :P = '@z 99,999;
,999.99' : 14 :H = '自年初累计(元)' , PJ :P =;
 '@z 999,999' : 15 :H = '平均(元/人)' ,a:h='人数'  noedit titl '按职称（职业资格）等级统计自年初累计工资总额（单位：元）'
  copy to D:\职称等级收入分析 fiel zcdj1,jjhj,gzhj,hj,pj,a type xls
 =messagebox("职工收入按职称（职业资格）等级统计分析数据已导出到“D：\职称等级收入分析”电子表中！！",48,"恭喜")  
***********************  
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
if dhsr=1 
*********与上年同期比较绩效工资************
  use &sblj.zr11&ly1 &&(上年同期明细)
  replace MOU with 0 , HJ with 0 all
  for jJB = 1 to val(yf)
    jjJB = iif(jJB>9,str(jJB,2),'0'+str(jJB,1))
    if  not file('GZ'+ly+jjJB+'.DBF')
      exit
    endif
    replace HJ with round(HJ+J&jjjb,0),MOU with val(jjJB) all
    repl PJ with round(HJ/mou,0) , A with 1 for mou>0
  endfor 
  sum hj,mou to hjtem,moutem
  if hjtem=0 or moutem=0 
     use &sblj.zr11temp
     copy to &sblj.zr11&ly1
  endif    
endif
close all
sele 5
use ryzk
index on RYBH to ryzk
select 4
if dhsr=1 
  use &sblj.zr11&ly1 excl &&(上年同期明细)
  repl A with 1 all
  index on rybh to &sblj.zr11&LY1
else
  use &sblj.zr11temp excl
  repl hj with jjhj,PJ with round(HJ/mou,0) , A with 1 for mou>0
  index on rybh to &sblj.zr11temp 
endif
set relation to rybh into 5
repl cjdm with e.cjdm,m03 with val(e.zw) for rybh=e.rybh
repl cjdm with '99',cjmc with '科级干部' for m03=5 or m03=6
repl cjdm with '88',cjmc with '处级干部' for m03=3 or m03=4
IF yn = 6
delete for jjhj=0 or cjmc = '处级干部'
pack
ENDI
index on CJDM to temp
total on CJDM to &sblj.zr11 &&(上年同期汇总)
select 3
use &sblj.zr&LY
copy to temp
use temp excl
repl hj with jjhj,PJ with round(HJ/mou,0) for mou>0
repl A with 1 all
set relation to rybh into 5
repl cjdm with e.cjdm,m03 with val(e.zw) for rybh=e.rybh
repl cjdm with '99',cjmc with '科级干部' for m03=5 or m03=6
repl cjdm with '88',cjmc with '处级干部' for m03=3 or m03=4
IF yn = 6
delete for jjhj=0 or cjmc = '处级干部'
pack
ENDI
index on CJDM to &sblj.zr&LY
total on CJDM to &sblj.zr1  &&(今年汇总)
select 2
use &sblj.zr11
replace PJ with round(HJ/A,0) for a>0
sum to LJ1,RS1 HJ,A
index on CJDM to &sblj.zr11
select 1
use &sblj.zr1 excl
delete for jjhj=0
pack
sum jjHJ,A to LJ,RS
set relation to cjdm into 2
replace PJ with round(jjHJ/A,0) for a>0 
repl m01 with round(PJ-b.PJ,0)  all
replace m02 with round(m01/b.PJ*100,1) for b.pj>0
appe blank
go bott
repl cjmc with '全厂奖金累计',hj with LJ,pj with round(lj/rs,2),;
m01 with round(LJ/RS,1)-round(LJ1/RS1,1),m02 with (round(LJ/RS,1)-round(LJ1/RS1,1))/round(LJ1/RS1,1)*100 
copy to jjtemp fiel cjmc,hj,pj,m01,m02
go top
    browse field CJMC:H = ' 单位名称 ' , HJ :P = '@z 99,999;
,999.99' : 14 :H = '自年初累计(元)' , PJ :P =;
 '@z 999,999' : 15 :H = '平均(元/人)' , m01;
 : 20 :H = '比上年增长(元/人)' , m02;
 : 14 :H = '增幅（%）' noedit titl '二、全厂自年初累计绩效工资统计表（单位：元）'
 copy to D:\绩效工资统计分析 fiel cjmc,jjhj,gzhj,hj,pj,m01,m02,a type xls
 =messagebox("绩效工资统计数据已导出到“D：\绩效工资统计分析”电子表中！！",48,"恭喜")  
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
  go top
  browse for d.PJ>0 field  CJMC : 12 :H = '车 间',bmmc:12:h='部 门',RYXM : 8 :H = '姓  名  ';
 , HJ :P = '@z 99,999,999.99' : 10 :H = ' 绩效工资 ' , d.HJ :P =;
 '@z 99,999,999.99' : 15 :H = '上年同期奖金 ' , PJ :P = '@z 999,999' :;
 10 :H = ' 平均奖金 ' , d.PJ :P = '@z 99,999' : 17 :H = ' 上年同期平均;
 ' , ZR5 = round((PJ-d.PJ)/d.PJ*100,1) : 17 :H = '比上年增长（%）';
 part zdz1 noedit titl '三、职工绩效工资分配及增长情况 [按 F1 键搜索某一职工]'
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
use gz&LY
inde on rybh to gz&LY
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
close all
retu
