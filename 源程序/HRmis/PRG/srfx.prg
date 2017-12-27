*****************************
* .\SRFX.PRG(Build2017.1.10)
*****************************
close all
USE sbdmk
统计月份=更新月份
**********可系统设置自定义月份*******20170110************
USE deset
LOCATE FOR ALLTRIM(oop)='px'
fbl=seted
USE dmk
大职务=VAL(大职务代码)
小职务=VAL(小职务代码)
set safety off
on key label F1 do grcx
**************双位月数自动处理*******************
IF month(DATE())<=统计月份
   jn=STR(YEAR(DATE()),4)
   nd=STR(YEAR(DATE())-1,4)
   YF1 = 12
   MESSAGEBOX("系统设置缴费基数更新月份为"+ALLTRIM(STR(统计月份))+"月，在此月之前统计分析"+nd+"年工资数据；需要统计分析"+jn+"年工资数据,请先系统设置缴费基数更新月份在本月之前。",64,"提示")
*******对上年全年收入统计分析时间的月份最大限期到“缴费更新月份”，之后默认自然年*******20170110***********
ELSE
   nd=STR(YEAR(DATE()),4)
   YF1 = month(DATE())
ENDIF 
LY1=str(val(ND)-1,4)
*年均人数=平均人数
*************************************************
if  not file('zr'+LY1+'.dbf')
  =MESSAGEBOX("请先处理"+LY1+"年工资总额!",48,"提示")
    close all
    return
endif    
if  not file('zr'+ND+'.dbf')
  =MESSAGEBOX("请先处理"+ND+"年工资总额!",48,"提示")
    close all
    return
endif
close all
delete file zr11&LY1.dbf
delete file zr11temp.dbf
delete file temp?.dbf
sele 2
use ryzk
inde on rybh to ryzk
sele 1
use zr&ND
inde on rybh to zr&ND
upda on rybh from B repl zzdm with b.zzdm,zzjg with b.zzjg,cjdm with b.cjdm , cjmc with b.cjmc , bmbh with b.bmbh , bmmc with b.bmmc, zw with b.zw,;
zw1 with b.zw1 , gwfl with b.gwfl , gwfl1 with b.gwfl1,zcdj with b.zcdj,zcdj1 with b.zcdj1,ygxz with b.ygxz,ygxz1 with b.ygxz1,zgqk with b.zgqk,zgqk1 with b.zgqk1
copy to zr11&ND for pj>0 AND  '职工'$ygxz1 and  '在岗'$zgqk1 
*****创建今年职工收入分析临时库，新上岗无工资收入职工除外*********
copy to zr11&LY1 stru
use zr11&LY1 excl
appe from ryzk for '职工'$ygxz1 and  '在岗'$zgqk1 
**** &&使用现在的ryzk.dbf在岗职工信息生成上年收入分析数据库********主要针对组织机构变动的特殊情况而言，用目前的ryzk.dbf人员状况信息可以虚拟出上年不存在的新增组织机构，但必须使用上年真实工资数据，不含今年新增人员************20160121*************************
repl ylbx with 0,qynj with 0,sybx with 0,ybx with 0,gjj with 0,sds with 0 all
***********语法严谨性：循环递加初始化*****清零ryzk.dbf中原有数据************
yesno = MESSAGEBOX('需要与'+LY1+'年'+ALLTRIM(STR(YF1))+'月同期比较吗？否则与'+LY1+'年月平均收入比较',291,'提示')
*****选择与上年同期比较或月平均比较
 DO CASE 
CASE yesno=6
   M1=1
  do while M1<=yf1
  MGZ = iif(M1>9,str(M1,2),'0'+str(M1,1))
  wait wind  '正在填写'+LY1+'年'+allt(str(val(mgz)))+'月职工工资总额...'  nowa
  if  not file('GZ'+LY1+MGZ+'.dbf')
 ***************ERP****************************
    MESSAGEBOX("请导入ERP系统中"+LY1+"年"+allt(str(val(MGZ)))+"月工资数据",48,"提示")
  endif  
***********工伤产假人员本月无工资但缴ylbx**********
    close all
    select 2
    use gzk excl
    zap
    appe from gz&LY1.&Mgz
    ***********************法1：同期比较用上年月工资库重新累计*********20160121**************************
    index on RYBH to gzk
    select 1
    use zr11&LY1
    index on RYBH to zr11&LY1
    upda on rybh from B replace ylbx with ylbx+b.ylbx,qynj with qynj+b.qynj,ybx with ybx+b.ybx,sybx with sybx+b.sybx,gjj with gjj+b.gjj,sds with sds+b.sds,sfje with sfje+b.sfje,m&Mgz with b.hj-b.jang ,j&Mgz with b.jang 
    M1 = M1+1   
 enddo
close all
wait wind  '同期比较用自年初至'+LY1+'年'+ALLTRIM(STR(yf1))+'月工资库重新累计（实际月数）...'  nowa
use zr11&LY1 excl
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
delete for mou=0
pack
repl gzhj with m01+m02+m03+m04+m05+m06+m07+m08+m09+m10+m11+m12,jjhj with j01+j02+j03+j04+j05+j06+j07+j08+j09+j10+j11+j12,hj with gzhj+jjhj,pj with round(hj/mou,0) all
************************************************************重新计算上年同期收入水平，生成新的月平均数*************20160121***********************************************************************************************
CASE yesno=7
****是否同比？（7）默认按上年月平均比较*******20160121********************
close all
sele 2
use zr&LY1
inde on rybh to zr&LY1
***********法2：上年年终原始数据与今年同期收入比较******20160121********************
sele 1
use zr11&LY1 excl
inde on rybh to zr11&LY1
upda on rybh from B repl mou with b.mou,sfje with b.sfje,jjhj with b.jjhj,gzhj with b.gzhj,hj with b.HJ,pj with b.pj;
ylbx with b.ylbx,qynj with b.qynj,ybx with b.ybx,sybx with b.sybx,gjj with b.gjj,sds with b.sds
CASE yesno=2
     RETURN
ENDCASE
delete for mou=0
pack
**************************初始化完毕***生成上年同期工资总额收入分析临时数据库（分两种情况）**********20160121*********************
close all
use zr11&LY1 excl 
index on rybh to zr11&LY1
repl cjdm with '77',cjmc with '部门负责人' for val(zw)>大职务 and val(zw)<=小职务
repl cjdm with '88',cjmc with '单位领导' for val(zw)<=大职务
index on CJDM+rybh to zr11&LY1
DELETE FOR pj=0
****&&(关键细节：各单位上年总额与各单位上年相加总额的人数一致，不含今年增减人员，必须得到各单位上年真实的人均收入水平，才能计算出今年的实际增跌幅%******20160121*****************************************)
GO top
BROWSE TITLE '请认真审核自年初至'+LY1+'年'+ALLTRIM(STR(yf1))+'月收入情况，接下来将以此表数据与'+ND+'年同期收入比较，并计算出各单位实际增跌幅％' 
pack
index on CJDM to zr11&LY1
total on CJDM to zr11
USE zr11
sum to LJ1,RS1 HJ,A
appe blank
go bott
repl cjmc with '合　　计',hj with LJ1,a WITH rs1
if yesno=6 
   replace PJ with round(HJ/A,0) all      
  else 
   replace PJ with round(HJ/A/yf1,0) all   
endif
GO top
browse field CJMC:H = ' 单位名称 ' , A :P = '@z 9,999' : 15 :H = '人数', HJ :P = '@z 999,999,999.99' : 14 :H = '自年初累计(元)' , PJ :P = '@z 999,999' : 15 :H = '平均' TITLE '请认真审核'+LY1+'年同期累计情况'
******** &&(上年同期汇总生成*收入分析*表中各单位的人数必须是上年有收入的人员，否则会降低上年同期各单位的真实收入水平**20160121**************)
use zr11&ND excl 
inde on rybh to zr11&ND
repl cjdm with '77',cjmc with '部门负责人' FOR val(zw)>大职务 and val(zw)<=小职务
repl cjdm with '88',cjmc with '单位领导' for val(zw)<= 大职务
index on CJDM+rybh to zr11&ND
DELETE FOR pj=0
****&&(关键细节：直接删除上年库中无工资总额的人员，确保各单位上年总额与各单位上年相加总额的人数一致，不含今年增减人员，必须得到各单位上年真实的人均收入水平，才能计算出今年的实际增跌幅%******20160121*****************************************)
GO top
BROWSE TITLE '请认真审核自年初至'+ND+'年'+ALLTRIM(STR(yf1))+'月收入情况，接下来将以此表数据与'+LY1+'年同期收入比较，并计算出各单位实际增跌幅％' 
pack
index on CJDM to zr11&ND
total on CJDM to zr1  &&(本年自年初统计汇总)
USE zr1 EXCLUSIVE
sum to LJ,RS HJ,A
appe blank
go bott
repl cjmc with '合　　计',hj with LJ,a WITH rs
if yesno=6 
   replace PJ with round(HJ/A,0) all      
  else 
   replace PJ with round(HJ/A/yf1,0) all   
endif
GO top
browse field CJMC:H = ' 单位名称 ' , A :P = '@z 9,999' : 15 :H = '人数', HJ :P = '@z 999,999,999.99' : 14 :H = '自年初累计(元)' , PJ :P = '@z 999,999' : 15 :H = '平均' TITLE '请认真审核'+ND+'年累计情况'
DELETE FOR  '合　　计'$cjmc
pack
close all
select 2
use zr11
repl m12 with ylbx+qynj+sybx+ybx+gjj all
sum HJ,m12,sds,sfje,A to LJ1,sb1,gs1,sf1,RS1
index on CJDM to zr11
if yesno=6 
   replace PJ with round(HJ/A,0),jang with round(sfje/A,0) all      
 else 
   replace PJ with round(HJ/A/yf1,0),jang with round(sfje/A/yf1,0) all   
endif
select 1
use zr1 
repl m12 with ylbx+qynj+sybx+ybx+gjj all
sum HJ,m12,sds,sfje,A to LJ,sb,gs,sf,RS
index on CJDM to zr1
if yesno=6 
   replace PJ with round(HJ/A,0),jang with round(sfje/A,0) all      
 else 
   replace PJ with round(HJ/A/yf1,0),jang with round(sfje/A/yf1,0) all   
endif
upda on cjdm from B repl m01 with round(PJ-b.PJ,0),m03 with round(jang-b.jang,0),m02 with round(m01/b.PJ*100,2),m04 with round(m03/b.jang*100,2) 
appe blank
repl cjdm with '99' for empty(cjdm)
go bott
if yesno=6 
   repl cjmc with '合　　计',hj with LJ,pj with round(lj/rs,2),m01 with round(LJ/rs,1)-round(LJ1/rs1,1),;
m02 with m01/round(LJ1/rs1,1)*100,jang with round(sf/rs,2),m04 with (round(sf/rs,1)-round(sf1/rs1,1))/round(sf1/rs1*100,2),m12 with sb,sds with gs,sfje with sf,a with rs      
 else 
   repl cjmc with '全年工资总额',hj with LJ,pj with round(lj/rs/yf1,2),m01 with round(LJ/rs/yf1,1)-round(LJ1/rs1/12,1),;
m02 with m01/round(LJ1/rs1/12,1)*100,jang with round(sf/rs/yf1,2),m04 with (round(sf/rs/yf1,1)-round(sf1/rs1/12,1))/round(sf1/rs1/12*100,2),m12 with sb,sds with gs,sfje with sf,a with rs  
ENDIF
   go top
    browse field CJMC:H = ' 单位名称 ' , HJ :P = '@z 999,999;
,999.99' : 14 :H = '自年初累计(元)' , PJ :P =;
 '@z 999,999' : 15 :H = '平均' , m01;
 :P = '@z 999,999.99':20 :H = '增长' , m02;
 : 14 :H = '增幅（%）', m12:P = '@z 999,999,999.99':h='代扣代缴三险两金',sds:P = '@z 999,999,999.99':h='个人所得税',sfje:P = '@z 99,999,999.99':h='实发金额',jang:P = '@z 9999,999.99':h='税后人均收入',m04;
 : 14 :H = '实增（%）' noedit titl '一、自年初累计工资总额统计表（单位：元）'
 copy to &bf.\收入分析 fiel cjmc,jjhj,gzhj,hj,pj,m01,m02,m12,sds,sfje,jang,m04,a type xl5
 close all
 sele 2
 use zr11&LY1 &&上年累计明细表
 inde on rybh to zr11&LY1
 sele 1
  use zr11&ND &&今年累计明细表
  inde on RYBH to zr11&ND
  upda on rybh from B repl m01 with b.hj,m02 with b.pj
 **********************将上年数据更新到今年工资总额中比较*******真实性*****************
  repl jang with PJ- m02,m04 with round(jang/m02*100,2) for m02>0
  inde on bmbh+rybh to zr11&ND
  go top
  browse field  CJMC : 12 :H = '车 间',bmmc:12:h='部 门',RYXM : 8 :H = '姓  名  ';
 , HJ :P = '@z 99,999,999.99' : 10 :H = ' 工资总额 ' , m01 :P =;
 '@z 99,999,999.99' : 15 :H = '上年总额 ' , PJ :P = '@z 999,999' :;
 10 :H = ' 平均工资 ' , m02 :P = '@z 999,999' : 17 :H = ' 上年月平均;
 ' , jang :P = '@z 999,999' : 17 :H  = '比上年增长（元）',m04 : 17 :H = '增幅（%）';
 part 20 noedit titl '二、职工收入情况一览表 [按 F1 键搜索某一职工]'    
  sort on hj to temp DESC
  use temp 
  go top
  browse field  CJMC : 12 :H = '车 间',bmmc:12:h='部 门',RYXM : 8 :H = '姓  名  ';
 , HJ :P = '@z 99,999,999.99' : 10 :H = ' 工资总额 ' , m01 :P =;
 '@z 99,999,999.99' : 15 :H = '上年总额 ' , PJ :P = '@z 999,999' :;
 10 :H = ' 平均工资 ' , m02 :P = '@z 999,999' : 17 :H = ' 上年月平均;
 ' , jang :P = '@z 999,999' : 17 :H  = '比上年增长（元）',m04 : 17 :H = '增幅（%）',ylbx:h='养老保险',ybx:h='医疗保险',sybx:h='失业保险',gjj:h='公积金',qynj:h='企业年金',sds:h='所得税',sfje:P = '@z 999,999.99':h='实发金额' ;
 part 20 noedit titl '三、工资累计排序一览表 [按 F1 键搜索某一职工]'
 copy to &bf.\工资累计排序 fiel rybh,cjmc,bmmc,ryxm,hj,m01,pj,m02,jang,m04,ylbx,ybx,sybx,gjj,qynj,sds,sfje
 copy to &bf.\工资累计排序 fiel rybh,cjmc,bmmc,ryxm,hj,m01,pj,m02,jang,m04,ylbx,ybx,sybx,gjj,qynj,sds,sfje type xl5
 close all
***********************
use zr11&ND excl
yn3 = MESSAGEBOX("管理骨干不参加工资总额增长比较吗？",4+32,"提示")
IF yn3 = 6
delete for val(zw)=小职务
pack
ENDI
inde on gwfl to zr11&ND
tota on gwfl to hztemp
use hztemp
repl pj with round(hj/a,0) all
inde on pj to hztemp
go top
brow fiel gwfl1:h='人员分类',HJ :P = '@z 99,999;
,999.99' : 14 :H = '自年初累计(元)' , PJ :P =;
 '@z 999,999' : 15 :H = '平均(元/人)' ,a:h='人数'  noedit titl '各岗位人员自年初累计工资总额统计表（单位：元）'
     copy to &bf.\人员分类收入分析 fiel gwfl1,jjhj,gzhj,hj,pj,a type xl5
use zr11&ND
inde on zcdj to zr11&ND
tota on zcdj to hztemp
use hztemp
repl pj with round(hj/a,0) all
inde on pj to hztemp
go top
brow fiel zcdj1:h='职称（职业资格）等级',HJ :P = '@z 99,999;
,999.99' : 14 :H = '自年初累计(元)' , PJ :P =;
 '@z 999,999' : 15 :H = '平均(元/人)' ,a:h='人数'  noedit titl '按职称（职业资格）等级统计自年初累计工资总额（单位：元）'
 copy to &bf.\职称等级收入分析 fiel zcdj1,jjhj,gzhj,hj,pj,a type xl5
close all
use zr11&LY1 excl
IF yn3 = 6
delete for val(zw)=小职务
pack
ENDI
repl PJ with round(jjHJ/mou,0) all
******收入分析**调换成**绩效工资分析*****
repl cjdm with '77',cjmc with '部门负责人' for val(zw)>大职务 and val(zw)<=小职务
repl cjdm with '88',cjmc with '单位领导' for val(zw)<=大职务
index on CJDM to zr11&LY1 
total on CJDM to zr11 &&(上年同期汇总生成*绩效工资分析*表)
use zr11&ND
repl PJ with round(jjHJ/mou,0) all
******收入分析**调换成**绩效工资分析*****
repl cjdm with '77',cjmc with '部门负责人' for val(zw)>大职务 and val(zw)<=小职务
repl cjdm with '88',cjmc with '单位领导' for val(zw)<=大职务
index on CJDM to zr11&ND
total on CJDM to zr1  &&(今年汇总)
close all
select 2
use zr11
if yesno=6
   replace PJ with round(jjHJ/A,0) all      
 else 
   replace PJ with round(jjHJ/A/yf1,0) all   
endif
sum to LJ1,RS1 jjHJ,A
index on CJDM to zr11
select 1
use zr1
if yesno=6
   replace PJ with round(jjHJ/A,0) all      
 else 
   replace PJ with round(jjHJ/A/yf1,0) all   
endif 
sum jjHJ,A to LJ,RS
index on CJDM to zr1
upda on cjdm from B repl m02 with b.PJ
repl m01 with round(PJ-m02,0) , m03 with round(m01/m02*100,2) FOR m02>0
*********************************************新进员工上年度无平均收入除外***************
appe blank
repl cjdm with '99' for empty(cjdm)
go bott
repl cjmc with '合　　　计',jjHJ with LJ,pj with round(lj/rs,2),;
m01 with round(LJ/RS,1)-round(LJ1/RS1,1),m03 with (round(LJ/RS,1)-round(LJ1/RS1,1))/round(LJ1/RS1,1)*100 
go top
    browse field CJMC:H = ' 单位名称 ' , jjHJ :P = '@z 99,999;
,999.99' : 14 :H = '自年初累计(元)' , PJ :P =;
 '@z 999,999' : 15 :H = '平均(元/人)' , m01:P = '@z 999,999';
 : 20 :H = '比上年增长(元/人)' , m03;
 : 14 :H = '增幅（%）' noedit titl '四、自年初累计绩效工资统计表（单位：元）'
 copy to &bf.\绩效工资统计分析 fiel cjmc,jjhj,pj,m02,m01,m03,a type xl5
copy to jjtemp fiel cjmc,jjHJ,pj,m02,m01,m03
use 绩效工资统计 excl
zap
appe from jjtemp
repl 车间 with cjmc,绩效工资 with jjHJ,人均 with pj,增资额 with m01,增幅 with m03 all
close all
 sele 2
 use zr11&LY1 &&上年累计明细表
 inde on rybh to zr11&LY1
 sele 1
 use zr11&ND &&今年累计明细表
 inde on RYBH to zr11&ND
 upda on rybh from B repl m01 with b.jjHJ,m02 with b.pj
 **********************将上年数据更新到今年工资总额中比较*******真实性*****************
  repl jang with round((PJ-m02)/m02*100,2) for m02>0
  inde on bmbh+rybh to zr11&ND
  for zdz=1 to fcount()
  zd=fiel(zdz)
  zdz1=(fsize(zd)+fsize(zd))*2.4
  if uppe(allt(zd))='RYXM' 
     exit
  endif
  endfor 
  go top
  browse field  CJMC : 12 :H = '单　位',bmmc:12:h='部 门',RYXM : 8 :H = '姓  名  ';
 , jjHJ :P = '@z 99,999,999.99' : 10 :H = ' 绩效工资 ' , m01 :P =;
 '@z 99,999,999.99' : 15 :H = '上年绩效工资 ' , PJ :P = '@z 999,999' :;
 10 :H = ' 平均绩效工资 ' , m02 :P = '@z 99,999' : 17 :H = ' 上年月平均;
 ' ,  jang :H   =  '比上年增长（%）';
 part zdz1 noedit titl '四、员工绩效工资分配及增长情况 [按 F1 键搜索某一职工]'
 sort on jjHJ to temp DESC
  use temp
  go top
  browse field  CJMC : 12 :H = '单　位',bmmc:12:h='部 门',RYXM : 8 :H = '姓  名  ';
 , jjHJ :P = '@z 99,999,999.99' : 10 :H = ' 绩效工资 ' ,m01 :P =;
 '@z 99,999,999.99' : 15 :H = '上年绩效工资 ' , PJ :P = '@z 999,999' :;
 10 :H = ' 平均绩效工资 ' , m02 :P = '@z 99,999' : 17 :H = ' 上年月平均;
 ' , jang :H   = '比上年增长（%）';
 part zdz1 noedit titl '四、员工绩效工资累计排序 [按 F1 键搜索某一职工]'
 copy to  jjtemp fiel cjmc,bmmc,ryxm,jjHJ,m01,pj,m02,jang
use 绩效工资增幅 excl
zap
appe from jjtemp
repl 车间 with cjmc,部门 with bmmc,绩效工资 with jjHJ,平均 with pj,平均增资额 with pj-m02,增幅 with jang all
close all
use ryzk
copy to temp for '职工'$ygxz1 and  '在岗'$zgqk1 
close all
select 3
use zr&ND
copy to tempfx
use tempfx excl
yxyn = MESSAGEBOX("部门负责人以上领导参加络伦茨曲线收入分析吗？",4+32,"提示")
IF yxyn = 7
delete FOR val(zw)>大职务 and val(zw)<=小职务
pack
ENDI
index on RYBH to zr&ND
sele 2
use temp
index on RYBH to temp
select 1
use ry0 excl
zap
appe from temp
repl hj with 1 all
index on RYBH to ry0
upda on rybh from C repl hj with c.hj,sfje with c.sfje
set relation to RYBH into 2
replace 序号 with '00001' , NLD with '35岁以下' for   b.NLD='20岁以下' or b.NLD='21—25岁' or b.NLD='26—30岁' or;
 b.NLD='31—35岁'
replace 序号 with '00002' , NLD with '36—50岁' for b.NLD='36—40岁' or b.NLD='41—45岁' or b.NLD='46—50岁'
replace 序号 with '00003' , NLD with '51—60岁' for b.NLD='51—55岁' or b.NLD='56—60岁'
close all
use ry0
replace 序号 with '  1  ' for NLD='35岁以下'
replace 序号 with '  2  ' for NLD='36—50岁'
replace 序号 with '  3  ' for NLD='51—60岁'
count for ('小'$WHCD2 and NLD='35岁以下') or ('初'$WHCD2 and NLD='35岁以下');
 to XC1
count for ('科'$WHCD2 and NLD='35岁以下') or ('硕士'$WHCD2 and NLD='35岁以下');
 to DZ1
count for ('小'$WHCD2 and NLD='36—50岁') or ('初'$WHCD2 and NLD='36—50岁');
 to XC2
count for ('科'$WHCD2 and NLD='36—50岁') or ('硕士'$WHCD2 and NLD='36—50岁');
 to DZ2
count for ('小'$WHCD2 and NLD='51—60岁') or ('初'$WHCD2 and NLD='51—60岁');
 to XC3
count for ('科'$WHCD2 and NLD='51—60岁') or ('硕士'$WHCD2 and NLD='51—60岁');
 to DZ3
count for '小'$WHCD2 or '初'$WHCD2 to XC
count for '科'$WHCD2 or '硕士'$WHCD2 to DZ
sum to 总收入 , 总人数 , ZNL , ZGL  hj , A , NL , GL 
总平均 = round(总收入/总人数,2)
count for hj<总平均*0.75 and NLD='35岁以下' to Z101
count for hj>=总平均*0.75 and hj<总平均*1.25 and NLD='35岁以下';
 to Z102
count for hj>=总平均*1.25 and hj<总平均*1.75 and NLD='35岁以下';
 to Z103
count for hj>=总平均*1.75 and NLD='35岁以下' to Z104
***********************************
count for hj<总平均*0.75 and NLD='36—50岁' to Z201
count for  hj>=总平均*0.75 and hj<总平均*1.25 and NLD='36—50岁';
 to Z202
count for hj>=总平均*1.25 and hj<总平均*1.75 and NLD='36—50岁';
 to Z203
count for hj>=总平均*1.75 and NLD='36—50岁' to Z204
***********************************
count for hj<总平均*0.75 and NLD='51—60岁' to Z301
count for hj>=总平均*0.75 and hj<总平均*1.25 and NLD='51—60岁';
 to Z302
count for hj>=总平均*1.25 and hj<总平均*1.75 and NLD='51—60岁';
 to Z303
count for hj>=总平均*1.75 and NLD='51—60岁' to Z304
***********************************
count for hj<总平均*0.75 to Z01
count for hj>=总平均*0.75 and hj<总平均*1.25 to Z02
count for hj>=总平均*1.25 and hj<总平均*1.75 to Z03
count for hj>=总平均*1.75 to Z04
************************************
index on NLD to ry0
total on NLD to ry1
use ry1
sort on 序号 to ry2
use ry3 excl
zap
append from ry2
close all
select 2
use ry3
index on 序号 to ry3
go top
select 1
use srfx excl
zap
append from ry3
index on 序号 to srfx
upda on 序号 from B  replace 年龄结构 with b.NLD , 总人数比例 with str(round(b.A/总人数*100;
,1),2,1)+'%' , 平均年龄 with round(b.NL/b.A,1) , 平均工龄 with round(b.GL/b.A;
,1) , 平均薪酬 with b.hj/b.A , 总收入比例;
 with str(b.hj/总收入*100,2)+'%'
delete for empty(allt(年龄结构))
pack
append blank
replace 序号 with '合 计'
replace 占平均75％ with str(Z101/总人数*100,3,1)+'%' for 序号='  1  '
replace 占平均75％ with str(Z201/总人数*100,3,1)+'%' for 序号='  2  '
replace 占平均75％ with str(Z301/总人数*100,3,1)+'%' for 序号='  3  '
replace 占平均75％ with str(Z01/总人数*100,3,1)+'%' for 序号='合 计'
replace 占75_125％ with str(Z102/总人数*100,3,1)+'%' for 序号='  1  '
replace 占75_125％ with str(Z202/总人数*100,3,1)+'%' for 序号='  2  '
replace 占75_125％ with str(Z302/总人数*100,3,1)+'%' for 序号='  3  '
replace 占75_125％ with str(Z02/总人数*100,3,1)+'%' for 序号='合 计'
replace 占125_175 with str(Z103/总人数*100,3,1)+'%' for 序号='  1  '
replace 占125_175 with str(Z203/总人数*100,3,1)+'%' for 序号='  2  '
replace 占125_175 with str(Z303/总人数*100,3,1)+'%' for 序号='  3  '
replace 占125_175 with str(Z03/总人数*100,3,1)+'%' for 序号='合 计'
replace 占175％上 with str(Z104/总人数*100,3,1)+'%' for 序号='  1  '
replace 占175％上 with str(Z204/总人数*100,3,1)+'%' for 序号='  2  '
replace 占175％上 with str(Z304/总人数*100,3,1)+'%' for 序号='  3  '
replace 占175％上 with str(Z04/总人数*100,3,1)+'%' for 序号='合 计'
replace 高中以下 with str(round(XC1/总人数*100,1),3,1)+'%' , 高中文化 with;
 str(round((val(left(总人数比例,2))/100*总人数-XC1-DZ1)/总人数*100,1),3,1)+'%';
 , 大专文化 with str(round(DZ1/总人数*100,1),3,1)+'%' for 序号='  1  '
replace 高中以下 with str(round(XC2/总人数*100,1),3,1)+'%' , 高中文化 with;
 str(round((val(left(总人数比例,2))/100*总人数-XC2-DZ2)/总人数*100,1),3,1)+'%';
 , 大专文化 with str(round(DZ2/总人数*100,1),3,1)+'%' for 序号='  2  '
replace 高中以下 with str(round(XC3/总人数*100,1),3,1)+'%' , 高中文化 with;
 str(round((val(left(总人数比例,2))/100*总人数-XC3-DZ3)/总人数*100,1),3,1)+'%';
 , 大专文化 with str(round(DZ3/总人数*100,1),3,1)+'%' for 序号='  3  '
replace 年龄结构 with '全体员工' , 总收入比例 with '100%' , 总人数比例 with;
 str(总人数,3)+'人' , 平均年龄 with round(ZNL/总人数,2) , 平均工龄 with round(ZGL/总人数;
,2) , 平均薪酬 with round(总收入/总人数;
,2) , 高中以下 with str(XC/总人数*100,3,1)+'%' , 高中文化 with str((总人数-XC-DZ)/总人数*100;
,3,1)+'%' , 大专文化 with str(DZ/总人数*100,3,1)+'%' for 序号='合 计'
browse noedit titl '五、收入分配公平分析（“络伦茨曲线”数据表）'
wjm='&bf'+'\收入公平分析'
copy to &wjm type xl5 
close all
sele 2
use ry0
inde on rybh to ry0
sele 1
use temp
repl gjj with 1 all
inde on rybh to temp
upda on rybh from B repl gjj with b.hj
close all
on key label F1
IF fbl=1
DEFINE WINDOW temp FROM INT((SROWS()-65)/2)  , INT((SCOLS()-310)/2)  TO INT((SROWS()+65)/2) , INT((SCOLS()+310)/2)  FLOAT CLOSE SHADOW DOUBLE title '六、员工收入按税前应发薪酬比较 [从“文件”菜单中点“查找”选项查找个人]'
ELSE
DEFINE WINDOW temp FROM INT((SROWS()-40)/2)  , INT((SCOLS()-160)/2)  TO INT((SROWS()+40)/2) , INT((SCOLS()+160)/2)  FLOAT CLOSE SHADOW DOUBLE title '六、员工收入按税前应发薪酬比较 [从“文件”菜单中点“查找”选项查找个人]'
ENDIF
activate window temp
sele gjj as 薪酬 , cjmc as 车间,bmmc as 部门,rybh as 编号,ryxm as 姓名,whcd2 as 学历,zyfl as 专业,zc as 职称,岗级,;
gz1 as 工种,gw as 岗位,岗等,zw1 as 职务,nl as 年龄,nld as 年龄段,gl as 工龄,gld as 工龄段 from temp ORDER BY gjj DESC 
clear window temp
close all
sele 2
use ry0
inde on rybh to ry0
sele 1
use temp
repl gjj with 1 all
inde on rybh to temp
upda on rybh from B repl gjj with b.sfje
close all
on key label F1
IF fbl=1
DEFINE WINDOW temp FROM INT((SROWS()-65)/2),INT((SCOLS()-310)/2)  TO INT((SROWS()+65)/2) , INT((SCOLS()+310)/2)  FLOAT CLOSE SHADOW DOUBLE title '六、员工收入按税后应发薪酬比较 [从“文件”菜单中点“查找”选项查找个人]'
ELSE
DEFINE WINDOW temp FROM INT((SROWS()-40)/2),INT((SCOLS()-160)/2)  TO INT((SROWS()+40)/2) , INT((SCOLS()+160)/2)  FLOAT CLOSE SHADOW DOUBLE title '六、员工收入按税后实发薪酬比较 [从“文件”菜单中点“查找”选项查找个人]'
ENDIF
activate window temp
sele gjj as 薪酬 , cjmc as 车间,bmmc as 部门,rybh as 编号,ryxm as 姓名,whcd2 as 学历,zyfl as 专业,zc as 职称,岗级,;
gz1 as 工种,gw as 岗位,岗等,zw1 as 职务,nl as 年龄,nld as 年龄段,gl as 工龄,gld as 工龄段 from temp ORDER BY gjj DESC 
clear window temp
close all
use ryzk
copy to temp
if file('zr'+LY1+'.dbf')
   sele 2
   use zr&LY1
   copy to zrtemp
   inde on rybh to zr&LY1
   BSND=LY1+'年应得收入'
   do bs
   sele 2
   use zrtemp  
   inde on rybh to zrtemp
   sum sfje to sf
   if sf>0
      repl hj with sfje all
      BSND=LY1+'年实发上银行帐户收入'
      do bs
   endif    
endif
close all
use ryzk
copy to temp
if file('zr'+ND+'.dbf')
   sele 2
   use zr&ND
   copy to zrtemp
   inde on rybh to zr&ND
   BSND=ND+'年应得收入'
   do bs
   sele 2
   use zrtemp
   inde on rybh to zrtemp
   sum sfje to sf
   if sf>0
      repl hj with sfje all
      BSND=ND+'年实发上银行帐户收入'
      do bs
   endif    
endif
clear
yn = MESSAGEBOX("数据导出成功！现打开电子表编辑使用吗？",4+32,"提示")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\收入分析.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\工资累计排序.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\人员分类收入分析.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\职称等级收入分析.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\绩效工资统计分析.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\收入公平分析.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
ENDI  
close all
delete files zr11temp.idx
delete files temp*.*
delete files srzf.dbf




proc bs
sele 1
use temp
repl gjj with 1 all
inde on rybh to temp
upda on rybh from B repl gjj with b.hj
close all
IF fbl=1
DEFINE WINDOW temp FROM INT((SROWS()-65)/2),INT((SCOLS()-310)/2)  TO INT((SROWS()+65)/2) , INT((SCOLS()+310)/2)  FLOAT CLOSE SHADOW DOUBLE title '员工收入按年薪比较('+BSND+')[从“文件”菜单中点“查找”选项查找个人]'
ELSE
DEFINE WINDOW temp FROM INT((SROWS()-40)/2),INT((SCOLS()-160)/2)  TO INT((SROWS()+40)/2) , INT((SCOLS()+160)/2)  FLOAT CLOSE SHADOW DOUBLE title '员工收入按年薪比较('+BSND+')[从“文件”菜单中点“查找”选项查找个人]'
ENDIF
activate window temp
sele gjj as 年薪 , cjmc as 车间,bmmc as 部门,rybh as 编号,ryxm as 姓名,whcd2 as 学历,zyfl as 专业,zc as 职称,岗级,;
gz1 as 工种,gw as 岗位,岗等,zw1 as 职务,nl as 年龄,nld as 年龄段,gl as 工龄,gld as 工龄段 from temp ORDER BY gjj DESC 
clear window temp
on key label F1 do grcx.fxp