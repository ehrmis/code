******************************
* .\YLBXJS.PRG (Build20161026)
******************************
close all
delete file lnk&LY1..dbf
YF0 = month(date())
YF00 = month(date())-1
RQ1 = day(date())
YF = iif(YF0>9,str(YF0,2),'0'+str(YF0,1))
YF1 = iif(YF00>9,str(YF00,2),'0'+str(YF00,1))
RQ111 = iif(RQ1>9,str(RQ1,2),'0'+str(RQ1,1))
close all
use 基础数据
go top
loca for 年份='&LY'
PJGZ = 上年社平￥
LL = 当年利率％/100
*if 年<=2006
*    LL = 当年利率‰/1000
*endif
LNLL = 历年利率％
GRSY = 失业保险％
XS = 系数
wait window '正在生成'+str(val(LY)+1,4)+'年养老保险......' nowait
ZDM = space(3)
go top
loca for 年份='&LY'
if val(LY)<=1997
  SN = 上年社平￥
else
  SN = 0
endif
DW0 = 单位缴纳％
GR0 = 个人缴纳％
SN0 = 上年社平￥
close all
IF file('ZR1'+LY+'.DBF')
**************必须先处理当年工资总额来修正当年实际缴费月数、人数****20160720*********
sele 3
use zr1&ly1 
inde on rybh to ZR1&ly1
sele 2
use zr1&ly 
inde on rybh to ZR1&ly
sele 1
IF !FILE('yl'+LY+'.dbf')
use ylbxk
COPY TO yl&LY stru
USE yl&LY
appe from zr1&LY for ygxz='01'
***********当年养老保险必须与当年实际缴费月数、人数相吻合*********20160720********************
inde on rybh to YL&ly
upda on rybh from B repl 缴费月数 with b.mou
upda on rybh from C repl 月缴费基数 with c.jfjs
**********************本年度（LY）实际缴费人员领了多少个月的工资就代扣代缴了多少个月的社会保险，而缴费基数是上年度（LY1）工资总额***********20160628*********************
ELSE
USE yl&LY
ENDIF
***********保留手工维护数据，下次计算不重复维护，但数据库生成太旧不与实际缴费人数一致的话，必须先删除数据库重新生成*********20160720********************
brow fiel cjmc:h='单位名称',bmmc:h='部门名称',ryxm:h='姓名', 缴费月数 , 月缴费基数 for 缴费月数<12 titl '请认真审核下列人员'+ly+'年实际缴费情况，若原生成数据库中人数与实际缴费人数不相等的话，请删除数据库重新生成！' 
ELSE
  MESSAGEBOX('请先进行'+LY+'年工资总额处理！！！','提示')
  return
ENDIF
replace 年份 with LY , A with 1 all
close all
******************初始化完毕正式计算个人帐户指标*****20160720***********
select 2
use yl&ly1
index on RYBH to yl&ly1
select 1
use yl&ly excl
index on RYBH to yl&ly
*************************正常个人帐户指标********************
 replace 缴费基数 with 月缴费基数*缴费月数 , 月单位缴费 with round(月缴费基数*DW0/100,2) , 单位缴费;
 with 月单位缴费*缴费月数 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费;
 with round(月缴费基数*GR0/100,2) , 个人缴费 with round(月个人缴费*缴费月数,2) , a with 1 all
  *****************************新增人员个人帐户指标审核维护*******************
 COUNT for 月缴费基数=0 or 缴费月数=0 TO rs
IF rs>0
  BROWSE FIELDS cjmc:h='单位名称',bmmc:h='部门名称',ryxm:h='姓名', 缴费月数 , 月缴费基数 ,总计 for 月缴费基数=0 or 缴费月数=0 TITLE '请输入'+LY+'年实际缴费月数、实际缴费基数；'+LY1+'年个人账户总计'
  replace  缴费基数 with 月缴费基数*缴费月数;
 , 月单位缴费 with round(月缴费基数*DW0/100,2) , 单位缴费 with 月单位缴费*缴费月数;
 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费 with round(月缴费基数*GR0/100;
,2) , 个人缴费 with round(月个人缴费*缴费月数,2) all   
ENDIF
*******************结算、计算历年本息****体会：编程语法严谨性******************
 replace 当年利息 with round(个人缴费*LL*XS*1/2,2) , 当年合计 with 个人缴费+当年利息 , 总累计 with  当年合计  all
 upda on rybh from B replace 总累计 with  round(b.总累计*(1+LNLL/100),2)+当年合计 , 历年累计 with b.总累计 , 历年本息 with round(b.总累计*(1+LNLL/100),2)
                                                                                               ************结算历年本息********20160622*************************
 repl  当年缴纳 with 个人缴费 , 累计利息 with 当年利息 , 上年社平 with PJGZ , 基础养老金 with round(上年社平*0.2;
,1) , 账户养老金 with round(总累计/120,2) , 养老待遇Ⅰ with round(基础养老金+账户养老金;
,1) , 个人本息 with 个人缴费+当年利息 all
****************新增员工**************
 replace 总累计 with  round(总计*(1+LNLL/100),2)+当年合计 , 历年累计 with 总计 for 总计>0
**************************************
inde on bmbh+rybh to yl&ly
go top
brow fiel cjmc:h='单位名称',bmmc:h='部门名称',ryxm:h='姓名', 缴费月数 , 月缴费基数 , 个人缴费 , 当年利息 , 当年合计 , 总累计 , 总计:h=ly1+'年总计' , 历年累计  for 缴费月数<12 titl '请认真审核'+ly+'年个人帐户记账金额及累计储存额' 
pack
go top
brow fiel cjmc:h='单位名称',bmmc:h='部门名称',ryxm:h='姓名', 缴费月数 , 月缴费基数 , 个人缴费 , 当年利息 , 当年合计 , 总累计 , 历年累计 , 总计:h=ly1+'年总计' titl '请认真审核'+ly+'年个人帐户记账金额及累计储存额' 
pack
inde on rybh to yl&ly
CLOSE all
return
