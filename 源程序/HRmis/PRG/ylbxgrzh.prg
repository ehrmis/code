*******************************
* .\ylbxgrzh.PRG(Build20161028)
*******************************
close all
***********调用DOS程序需要清除底图***************************
STJFN=''
STJFY=''
人可 = '00'
人可1 = ''
计发比例1 = ''
计发比例2 = ''
计发比例3 = ''
计发比例4 = ''
平均2=0
bf1=ALLTRIM(bf)+'\'
use ryzk
locate for allt(XM)=allt(RYXM) or val(bh)=val(RYBH)
    grbh1=val(grbh)
    abc='ryzk'     
    ********数据源************    
    nn=ALLTRIM(STR(YEAR(cjgzrq)))
    mm=IIF(MONTH(cjgzrq)>9,ALLTRIM(STR(MONTH(cjgzrq))),'0'+ALLTRIM(STR(MONTH(cjgzrq))))
    dd=IIF(day(cjgzrq)>9,ALLTRIM(STR(day(cjgzrq))),'0'+ALLTRIM(STR(day(cjgzrq))))
    gzrq=nn+mm+dd
    xm=ALLTRIM(ryxm)
    bh=ALLTRIM(rybh)
    *****姓名、编号标准查找方法*****以姓名显示、以编号精准查找*********
ly=nd
LY1=str(val(LY)-1,4)
use lnk
copy to &bf1.&xm stru
use 基础数据
go top
loca for 年份='&LY'
snspn = 上年社平￥
use sbdmk
精度值 = 指数化精度
交贴=jtf
书报洗理费=sbf+xlf
煤贴=mt
厂内=chn
close all
  sele 1
  use &bf1.&xm
  年代 = 1995
  for I = 1 to val(LY)-1994
    append from ryzk for allt(XM)=allt(RYXM) or val(bh)=val(RYBH)
    replace 年份 with str(年代,4) for empty(年份)
    年代 = 年代+1
  endfor
************遍历该人历年缴费基数及个人帐户后应重新生成临时数据库********20160713*****************
close all
IF file("ylgrzh.dbf")
   USE ylgrzh
   YNKK=RECCOUNT()
   IF YNKK=0  &&空库重输
      人可1='重输'
   ELSE   
      GO top   
      LOCATE FOR XM=allt(RYXM)
      IF !EOF()
         yn = MESSAGEBOX("需要重新输入"+XM+"养老保险手册中月缴费基数和缴费月数？",4+32,"提示")
         IF yn=6
           人可1='重输'
         ENDIF
      ENDIF
    ENDIF      
ENDIF 
*****************************是否重新刚刚输入的数据***二次输入快速处理*************************************     
use grzh excl
if file("grzhtemp.dbf") 
   zap
   appe from grzhtemp
*****************更新“云南社保系统”导出数据**************************
endif 
go top
loca for aae001=1995
if !eof() 
yn = MESSAGEBOX("是否导入“云南社保系统”中缴费基数和月数？",4+32,"提示")
IF yn = 6
   人可 = '卡片'
ENDIF   
endif 
*************(1)手册：初始化缴费基数、月数、养老金计算参数*****************
IF 人可1='重输'
   USE lnk
   COPY TO &bf1.&xm stru  
IF 人可 = '卡片' 
******从云南社保系统中自动提取历年数据输入******
use grzh
sort on aae001 to temp for val(grbh)=grbh1
**************汇总历年缴费数据*************************************
sele 2
use temp
go top
sele 1
use &bf1.&xm excl
年代 = 1995
for I = 1 to val(LY)-1994
    if abc='ryzk'
       appe from ryzk for BH=allt(RYbh)
    else
       appe from ryzk1 for BH=allt(RYbh)
    endif   
    replace 年份 with str(年代,4)
   年代 = 年代+1
endfor
replace 月缴费基数 with 0 for val(年份)<val(LY)
go top
do while !eof()    
    sele 2
    sum aic079 to ys for aae001=val(a.年份)
    go top
    loca for aae001=val(a.年份)    
    if !eof()
       replace a.缴费月数 with ys,a.月缴费基数 with jfjs
    endif
    sele 1
    skip
ENDDO
ELSE
******人工输入：历年个人养老保险手册记录月缴费基数＋历年个人帐户登记卡首年记录数据******
use &bf1.&xm excl
年代 = 1995
for I = 1 to val(LY)-1994
   if abc='ryzk'
       appe from ryzk for BH=allt(RYbh)
    else
       appe from ryzk1 for BH=allt(RYbh)
    endif 
    replace 年份 with str(年代,4)
    年代 = 年代+1
ENDFOR
ENDIF
SUM 缴费月数 TO 杨雷
IF 杨雷=0
replace 缴费月数 with 6 for 年份='1995'
replace 缴费月数 with 12 for 年份<>'1995'
replace 缴费月数 with 缴费月数-3 for 年份='2003'and 缴费月数>3
replace 说明 with str(val(年份)-1,4)+'年月平均工资' all
go bott
replace 缴费月数 with month(date()) , 说明 with 年份+'年实缴月份数'
******审核自动或人工输入“缴费月数 , 月缴费基数”******
ENDIF
go top
 browse for BH=allt(RYbh) field;
 RYBH :R : 8 :H = '编  号' , RYXM :R : 8 :H = '姓  名' , 年份 :R ;
,缴费月数 , 月缴费基数 , 说明:R titl '养老保险手册从本人缴费起始年输入月数、基数；请输末年缴费月数，2001年直接输2000年月平均工资'
delete for 月缴费基数=0
PACK
ELSE
    USE &bf1.&xm EXCLUSIVE
    ******审核自动或人工输入“缴费月数 , 月缴费基数”******
SUM 缴费月数 TO 杨雷
***********首次输入，自动识别（被杨雷诈骗）****20161031***********
IF 杨雷=0
replace 缴费月数 with 6 for 年份='1995'
replace 缴费月数 with 12 for 年份<>'1995'
replace 缴费月数 with 缴费月数-3 for 年份='2003'and 缴费月数>3
replace 说明 with str(val(年份)-1,4)+'年月平均工资' all
go bott
replace 缴费月数 with month(date()) , 说明 with 年份+'年实缴月份数'
******审核自动或人工输入“缴费月数 , 月缴费基数”******
ENDIF
go top
 browse field;
 RYBH :R : 8 :H = '编  号' , RYXM :R : 8 :H = '姓  名' , 年份 :R ;
,缴费月数 , 月缴费基数 , 说明:R titl '养老保险手册从本人缴费起始年输入月数、基数；请输末年缴费月数，2001年直接输2000年月平均工资'
delete for 月缴费基数=0
PACK    
ENDIF
*************云政发[2004]45号*****指数化工资计算式****************************************
go top
loca for 年份='2002'
if !eof()
   平均2=月缴费基数
endif
sum 月缴费基数 to yjshj
count for 月缴费基数>0 to yjscc
实缴平均=round(yjshj/yjscc,2)
close all
select 2
use 基础数据 EXCLUSIVE
delete for val(年份)>year(date())
PACK
if val(GZRQ)<19951001
     缴费起始年='1995'
  else
     缴费起始年=left(GZRQ,4)
endif
if val(GZRQ)<19951001
   缴费周年 = YEAR(DATE())-1996+round((MONTH(DATE())+3)/12,2)
                                           ******12-10+1=3(yfxs)*****
   缴费起始年='1995'
else
   yfxs=12-VAL(mm)+1
   缴费周年 = YEAR(DATE())-(val(缴费起始年)+1)+round((MONTH(DATE())+yfxs)/12,2)
   缴费起始年=left(GZRQ,4)
endif
if snspn=1996
*****？？*****************
   sum to 平均社平 上年社平￥ for val(年份)>val(缴费起始年)
   count to 年数 for val(年份)>val(缴费起始年)
else
   sum to 平均社平 上年社平￥ for val(年份)>=val(缴费起始年)
   count to 年数 for val(年份)>=val(缴费起始年)
endif
缴费虚年 = val(LY)-val(缴费起始年)+1
平均社平 = round(平均社平/年数,1)
index on 年份 to jcsj
go bott
上年社平=上年社平￥
社平养老金=上年养老金
go top
select 1
use &bf1.&xm
inde on 年份 to &bf1.&xm
update on 年份 from B replace 上年社平￥ with b.上年社平￥,上年养老金 with b.上年养老金,指数 with round(月缴费基数/上年社平￥,4)
*****************************本人指数化月平均工资*************************
sum 指数 to 个人指数1 
brow fiel 年份,ryxm:h='姓名',月缴费基数,上年社平￥,指数,上年养老金 noedit titl '指数合计='+str(个人指数,7,4)+';'+'平均月缴费基数='+str(实缴平均,7,4) 
****************************************审核结果手册记录参保缴费情况*****************************************************
REPLACE 个人指数 with 个人指数1,缴费年数 with 缴费虚年,上年社平 with snspn,指数化系数 with 1.3,缴费年限 with 缴费周年,平均基数 with 实缴平均,;
指数化工资 with round(round(个人指数/缴费年数,4)*上年社平,精度值) FOR VAL(年份)=VAL(nd) 
*'--月过渡性养老金和月缴费性调节金计算提示--'
*缴费年数(虚年减中断缴费年)
*上年社平 = 退休时上年社平
*指数化系数 = 指数化工资系数(%)
*缴费年限 = 实际缴费年限(月对月)
*实缴平均 = 平均月缴费基数
*指数化工资 = round(round(个人指数/缴费年数,4)*平均社平,精度值)
*指数化工资 = round(round(个人指数/缴费年数,4)*上年社平,精度值)
***********本人指数化月平均工资=（A1/B1+A2/B2+···+An/Bn）/n×退休上年度全省职工月平均工资*********A1=首年月缴费基数*******B1=首年上年社平*****
BROWSE fiel 年份,ryxm:h='姓名',个人指数,缴费年数,上年社平,指数化系数,缴费年限,平均基数,指数化工资,视同缴费年,几年,几个月,累计缴费年 TITLE  '请确认“月过渡性养老金和月缴费性调节金计算过程内容”（提示:参见个人帐户登记卡）'
REPLACE 指数化工资 with round(round(个人指数/缴费年数,4)*上年社平,精度值),视同缴费年 with round(几年+几个月/12,2),累计缴费年 WITH round(缴费年限+视同缴费年,2) FOR VAL(年份)=VAL(nd)
GO BOTTOM 
累计缴费年限=累计缴费年
use ylbxk
copy to ylgrzh stru 
use ylgrzh 
appe from &bf1.&xm
*******以上完成养老金计算参数初始化：基础养老金指数化工资＋过渡性养老金视同缴费年及手册填写情况****(2)卡片：自动或人工输入首年帐户(当年数据)******
IF 人可 = '卡片'
******自动输入******
close all
use ylgrzh 
go top
首年=val(年份)
close all
select 2
use 基础数据
loca for val(年份)=首年
sele 1
use ylgrzh
go top
LNLL = b.历年利率％
DW = b.单位缴纳％
GR = b.个人缴纳％
PJGZ = b.上年社平￥
年=val(年份)
if 年<=1997
  SN = b.上年社平￥
else
  SN = 0 
endif
LL = b.当年利率％/100
if 年<=2006
    LL = b.当年利率‰/1000
     replace 缴费基数 with 月缴费基数*缴费月数 , 月单位缴费 with round(月缴费基数*DW/100,1) , 单位缴费;
 with 月单位缴费*缴费月数 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费;
 with round(月缴费基数*GR/100,1) , 个人缴费 with  round(月个人缴费*缴费月数,1) , 比例 with GR;
 ,单位利息 with round((单位缴费+社平5)*LL*缴费月数/2,1);
 , 当年利息 with round(个人缴费*LL*缴费月数/2,1) , 当年合计;
 with 单位缴费+社平5+单位利息+个人缴费+当年利息,总累计 with 当年合计
ELSE
if 年<2012
 replace 缴费基数 with 月缴费基数*缴费月数 , 月单位缴费 with round(月缴费基数*DW/100,1) , 单位缴费;
 with 月单位缴费*缴费月数 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费;
 with round(月缴费基数*GR/100,1) , 个人缴费 with  round(月个人缴费*缴费月数,1) , 比例 with GR;
 ,单位利息 with round((单位缴费+社平5)*LL*缴费月数/2,1);
 , 当年利息 with round(个人缴费*LL*缴费月数/2,1) , 当年合计;
 with 单位缴费+社平5+单位利息+个人缴费+当年利息,总累计 with 当年合计
else
********2012年后个人缴费保留到分************************
 replace 缴费基数 with 月缴费基数*缴费月数 , 月单位缴费 with round(月缴费基数*DW/100,1) , 单位缴费;
 with 月单位缴费*缴费月数 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费;
 with round(月缴费基数*GR/100,2) , 个人缴费 with round(月个人缴费*缴费月数,2) , 比例 with GR;
 ,单位利息 with round((单位缴费+社平5)*LL*缴费月数/2,2);
 , 当年利息 with round(个人缴费*LL*缴费月数/2,2) , 当年合计;
 with 单位缴费+社平5+单位利息+个人缴费+当年利息,总累计 with 当年合计
ENDIF
endif
ENDIF
******人工输入******
close all 
use ylgrzh excl
go top
browse part 12 field 年份: r ,rybh:r:h='编号',RYXM:r:h='姓名' ;
 , 缴费月数 : 8 :H = '缴费月数' , 缴费基数 : 8 :H = '缴费基数';
 , 单位缴费 : 12 :H = '当年单位缴费' , 社平5 : 7 :H = '社平5%';
 , 单位利息 : 12 :H = '当年单位利息' , 比例 :h ='个人缴费比例%' , 个人缴费 : 12 :H = '当年个人缴费';
 , 当年利息 : 12 :H = '当年个人利息' , 当年合计 : 10 :H = '当年合计';
 , 历年单位缴 : 12 :H = '历年单位缴费' , 历年利息1 : 12 :H =;
 '历年单位利息' , 历年个人缴 : 12 :H = '历年个人缴费' , 历年利息2;
 : 12 :H = '历年个人利息' , 总累计 : 10 :H = '总    计' titl  '养老保险个人帐户登记卡输入本人缴费起始年数据 按 [ Ctrl+T ] 删除'  
*************(3)计算首年帐户(当年数据)******先初始化上年字段变量作为下年关联数据*********** 
单位缴=单位缴费
社平=社平5
单位息=单位利息
历年单位=历年单位缴
历息1=历年利息1
个人缴=个人缴费
当年息=当年利息
历年个人=历年个人缴
历息2=历年利息2
总计1=总累计&&严格区分ylgrzh.dbf中的“总计”2014-09-09
repl 个人本息 with 个人缴费+当年利息+历年个人缴+历年利息2 , 历年本息 with 历年个人缴+历年利息2 all
pack
go top
首年=val(年份)
loca for 年份='1998'
if 首年=1998
   个人总本息=个人本息
   个人历年本息=历年本息-历年利息2
endif
close all
use 基础数据
copy to temp for val(年份)>首年
select 2
use temp
inde on 年份 to temp
go top
sele 1
use ylgrzh
set relation to 年份 into 2
go top
IF !EOF()
   SKIP
ENDIF   
DO while !eof()
年=val(年份)
if 年<=1997
  SN = b.上年社平￥
else
  SN = 0 
endif
LL = b.当年利率％/100
if 年<=2006
    LL = b.当年利率‰/1000
endif
LNLL = b.历年利率％
DW = b.单位缴纳％
GR = b.个人缴纳％
PJGZ = b.上年社平￥
XS = b.系数
if 年<2012
********2012年后个人缴费保留到分************************
 replace 缴费基数 with 月缴费基数*缴费月数 , 月单位缴费 with round(月缴费基数*DW/100,1) , 单位缴费;
 with 月单位缴费*缴费月数 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费;
 with round(月缴费基数*GR/100,1) , 个人缴费 with round(月个人缴费*缴费月数,1) , 比例 with GR
ELSE
 replace 缴费基数 with 月缴费基数*缴费月数 , 月单位缴费 with round(月缴费基数*DW/100,1) , 单位缴费;
 with 月单位缴费*缴费月数 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费;
 with round(月缴费基数*GR/100,2) , 个人缴费 with round(月个人缴费*缴费月数,2) , 比例 with GR
endif 
if 年=2001
    do case
    case 缴费月数=1
      replace 单位缴费 with round(平均2*4/100,1) , 个人缴费 with round(平均2*7/100,1) , 缴费基数 with 平均2
    case 缴费月数=2
      replace 单位缴费 with round(平均2*4/100,1)*2 , 个人缴费 with round(平均2*7/100,1)*2 , 缴费基数 with;
 平均2*2
    case 缴费月数=3
      replace 单位缴费 with round(平均2*4/100,1)*3 , 个人缴费 with round(平均2*7/100,1)*3 , 缴费基数 with;
 平均2*3
    case 缴费月数>3
      replace 单位缴费 with 月单位缴费*(缴费月数-3)+round(平均2*4/100,1)*3 , 个人缴费 with;
 月个人缴费*(缴费月数-3)+round(平均2*7/100,1)*3 , 缴费基数 with 月缴费基数*(缴费月数-3)+平均2*3
    endcase
endif
  if 年=2002
    do case
    case 缴费月数=1
     replace 单位缴费 with round(月缴费基数*3/100,1) , 个人缴费 with round(月缴费基数*8/100,1)
    case 缴费月数=2
     replace 单位缴费 with round(月缴费基数*3/100,1)*2 , 个人缴费 with round(月缴费基数*8/100,1)*2
    case 缴费月数=3
     replace 单位缴费 with round(月缴费基数*3/100,1)*3 , 个人缴费 with round(月缴费基数*8/100,1)*3
    case 缴费月数>3
     replace 单位缴费 with 月单位缴费*(缴费月数-3)+round(月缴费基数*3/100,1)*3 , 个人缴费 with;
 月个人缴费*(缴费月数-3)+round(月缴费基数*8/100,1)*3
    endcase
  endif
  IF 年>=2007&&政策变化：单位缴费不计入个人帐户
 if 年<2012
********2012年后个人缴费保留到分************************ 
 replace 当年利息 with round(个人缴费*LL*XS*1/2,1) , 当年合计 with 个人缴费+当年利息 ,总累计 with  round(总计1*(1+LNLL/100),1)+当年合计 , 历年累计 with 总计1,历年本息 with round(总计1*(1+LNLL/100),1)
 ******************个人帐户累计储存额＝上年累计储存额×〔1＋历年利率）＋个人帐户当年记帐金额×（1＋当年利率×系数÷2〕*************************************************结算历年本息=已记帐本息**20160622******************
 repl  当年缴纳 with 个人缴费 , 累计利息 with 当年利息 , 上年社平 with PJGZ 
  ELSE
   replace 当年利息 with round(个人缴费*LL*XS*1/2,2) , 当年合计 with 个人缴费+当年利息 ,总累计 with  round(总计1*(1+LNLL/100),2)+当年合计 , 历年累计 with 总计1,历年本息 with round(总计1*(1+LNLL/100),2)
 ******************个人帐户累计储存额＝上年累计储存额×〔1＋历年利率）＋个人帐户当年记帐金额×（1＋当年利率×系数÷2〕*************************************************结算历年本息=已记帐本息**20160622******************
 repl  当年缴纳 with 个人缴费 , 累计利息 with 当年利息 , 上年社平 with PJGZ 
 endif
  ELSE
   replace 单位利息 with round((单位缴费+社平5)*LL*缴费月数/2,1);
 , 当年利息 with round(个人缴费*LL*缴费月数/2,1) , 当年合计;
 with 单位缴费+社平5+单位利息+个人缴费+当年利息 , 历年单位缴 with 单位缴+社平+单位息+历年单位+历息1;
 , 历年利息1 with round(历年单位缴*(LNLL/100),1) , 历年个人缴 with;
 个人缴+当年息+历年个人+历息2 , 历年利息2 with round(历年个人缴*(LNLL/100);
,1) , 总累计 with 当年合计+历年单位缴+历年利息1+历年个人缴+历年利息2;
 , 当年缴纳 with 单位缴费+个人缴费 , 累计利息 with 单位利息+当年利息+历年利息1+历年利息2;
 , 历年累计 with 总计 , 上年社平 with PJGZ , 基础养老金 with round(上年社平*0.2;
,1) , 账户养老金 with round(总累计/120,2) , 养老待遇Ⅰ with round(基础养老金+账户养老金;
,1) , 个人本息 with 个人缴费+当年利息+历年个人缴+历年利息2 , 历年本息 with 历年个人缴+历年利息2
 ENDIF
单位缴=单位缴费
社平=社平5
单位息=单位利息
历年单位=历年单位缴
历息1=历年利息1
个人缴=个人缴费
当年息=当年利息
历年个人=历年个人缴
历息2=历年利息2
总计1=总累计&&严格区分ylgrzh.dbf中的“总计”2014-09-09
skip
ENDDO
if 首年=1998  
   repl 个人本息 with 个人总本息 , 历年本息 with 个人历年本息 for val(年份)=1998
endif
close all
use ylgrzh
go top
*********校正***********
loca for val(年份)=2002
if !eof()
gz_21=月缴费基数
skip -1
if val(年份)=2001
   gz_20=月缴费基数
   REPLACE 月缴费基数 with round((gz_20*9+gz_21*3)/12,0)
endif
endif
************************
go top
******(4)审核个人帐户处理结果******
browse part 12 field 年份: r ,rybh:r:h='编号',RYXM:r:h='姓名' ;
 , 缴费月数 : 8 :H = '缴费月数' , 缴费基数 : 8 :H = '缴费基数';
 , 单位缴费 : 12 :H = '当年单位缴费' , 社平5 : 7 :H = '社平5%';
    , 单位利息 : 12 :H = '当年单位利息' , 个人缴费 : 12 :H = '当年个人缴费';
 , 当年利息 : 12 :H = '当年个人利息' , 当年合计 : 10 :H = '当年合计';
 , 历年单位缴 : 12 :H = '历年单位缴费' , 历年利息1 : 12 :H =;
 '历年单位利息' , 历年个人缴 : 12 :H = '历年个人缴费' , 历年利息2;
 : 12 :H = '历年个人利息' , 总累计 : 10 :H = '总    计';
 titl '养老保险个人帐户登记卡审核'
******(5)计算和审核退休人员养老金预测数据***************************生成电子表*************养老金报告清单**********************
BROWSE FIELDS 独生子女,工种年限:h='累计特殊工种年限(年)',工龄系数:h='折算工龄系数(小于1)' FOR VAL(年份)=VAL(nd) titl '“独生子女”、“特殊工种”退休待遇计算参数'
***********FOR VAL(年份)=VAL(nd)*****************
*do while .t.
*  @ 10 , 33 say '.'+xm+'封定工资(元):' get 岗技 picture '9999'
*  read
 * if 岗技>0
  *  exit
  *endif
*enddo
*if val(GZRQ)<19980401
*  do while .t.
*    @ 12 , 33 say '.'+xm+'个人帐户九七年总计(元):' get 九七 picture;
 '9999.9'
*    read
*    if 九七>0
*      exit
*    endif
*  enddo
*endif
*@ 14,33 clear to 14,70
*@ 14 , 33 color w+/b say '.'+allt(ryxm)+'档案记载工龄工资(元):' get b.GLGZ picture;
 '999.9'
*read
*@ 14,33 clear to 14,70
*@ 14 , 33 color w+/b say '.'+allt(ryxm)+'档案记载医药费(元):' get b.YLF picture;
 '999.9'
*read
GO bott
连续工龄 = round(gl+工种年限*工龄系数,2) 
***************系统初始化生成工龄FOR VAL(年份)=VAL(nd)*********20160714**********************
do case
case 连续工龄>=10 and 连续工龄<15
  计发比例1 = '65'
case 连续工龄>=15 and 连续工龄<20
  计发比例1 = '70'
  计发比例3 = ' 5'
case 连续工龄>=20 and 连续工龄<25
  计发比例1 = '75'
  计发比例3 = ' 5'
case 连续工龄>=25 and 连续工龄<30 and XB1='男'
  计发比例1 = '75'
  计发比例2 = '10'
case 连续工龄>=25 and 连续工龄<30 and XB1='女'
  计发比例1 = '75'
  计发比例2 = '10'
  计发比例4 = ' 5'
case 连续工龄>=30 and 连续工龄<35 and XB1='男'
  计发比例1 = '75'
  计发比例2 = '15'
case 连续工龄>=30 and 连续工龄<35 and XB1='女'
  计发比例1 = '75'
  计发比例2 = '15'
  计发比例4 = ' 5'
case 连续工龄>=35
  计发比例1 = '75'
  计发比例2 = '20'
endcase
replace 计发比例 WITH str(val(计发比例1)+val(计发比例2)+val(计发比例3)+val(计发比例4),2),计发月数 WITH 139 FOR VAL(年份)=VAL(nd)
************************************************老办法***************************************************************************
BROWSE FIELDS 计发比例,计发月数 FOR VAL(年份)=VAL(nd) titl '退休待遇计算参数' 
***************** FOR VAL(年份)=VAL(nd) ***************
JT = 煤贴+厂内
BT = 交贴+书报洗理费
DZF = 0
if 独生子女=1
 * DZF = round(岗技*5/100,1)
  DZF=ROUND(社平养老金*5/100,1)
ENDIF
****Ver2014****一、月基本养老金=基础养老金+个人账户养老金(+过渡性养老金)*******二、独生子女的职工退休，按照办理退休的上年度全省退休人员月平均基本养老金的5%，计发独生子女费***********三、单位津补贴********
**************1.基础养老金=（退休上年度全省职工月平均工资+本人指数化月平均工资）÷2×累计缴费年限×1%
**************2.个人账户养老金=个人账户储存额÷计发月数（按照退休时年龄确定计发月数：60周岁退休为139个月、55周岁为170个月、50周岁为195个月）
**************3.过渡性养老金=本人指数化月平均工资×视同缴费年限×1.3%
replace 基础养老金 with round((上年社平+指数化工资)/2*round(累计缴费年限,2)*1/100,1),账户养老金 with round(总累计/计发月数,1),过渡养老金 with round(指数化工资*指数化系数/100*round(视同缴费年,2),1) FOR VAL(年份)=VAL(nd)
if val(GZRQ)<19951001
   缴费调节金 = round(round((累计缴费年限-15)*0.4/100*上年社平,2)+round(缴费年限*实缴平均*0.6/100,2),1)
else
   缴费调节金 = round(缴费年限*实缴平均*0.6/100,1)
endif
REPLACE 养老待遇Ⅰ with 基础养老金+账户养老金+缴费调节金+DZF+JT+BT FOR VAL(年份)=VAL(nd)
**********************************法一：老办法*****************************
REPLACE 养老待遇Ⅱ with 基础养老金+账户养老金+过渡养老金+DZF+JT+BT FOR VAL(年份)=VAL(nd) 
**********************************法二：中办法*****************************
REPLACE 养老待遇Ⅲ with 基础养老金+账户养老金+DZF+JT+BT FOR VAL(年份)=VAL(nd)
**********************************法三：新办法*****************************
BROWSE FIELDS 年份,缴费月数,月缴费基数,缴费基数,单位缴费,单位利息,比例,个人缴费,当年利息,当年合计,历年单位缴,历年利息1,历年个人缴,历年利息2,总累计,养老待遇Ⅰ, 养老待遇Ⅱ,养老待遇Ⅲ FOR VAL(年份)=VAL(nd) titl '退休待遇'
copy to &bf1.&xm.历年帐户 type XL5 fiel 年份,缴费月数,月缴费基数,缴费基数,单位缴费,单位利息,比例,个人缴费,当年利息,当年合计,历年单位缴,历年利息1,历年个人缴,历年利息2,总累计,养老待遇Ⅰ, 养老待遇Ⅱ,养老待遇Ⅲ
yn = MESSAGEBOX("成功生成电子表“"+bf1+xm+"历年帐户”！需要编辑使用吗？",4+32,"提示")
IF yn = 6
FileName = GETFILE('XLS', '历年帐户: ', '确定')
IF EMPTY(alltrim(FileName)) 
    FileName='&xp'+'\历年帐户.xls'
ENDIF  
wjm=ALLTRIM(FileName)
**************标准Excel读取文件*************
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&wjm")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
release oleapp
WAIT CLEAR        
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf1"+"&xm"+"历年帐户.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表" 
ENDIF  
if !file("&bf1"+"历年帐户.DBF")
   copy to 历年帐户 type fox2x
else
   use 历年帐户 excl
   go top
   loca for BH=allt(RYbh)
   if !eof()
      delete for BH=allt(RYbh)
      pack
      appe from ylgrzh
   else
       appe from ylgrzh
   endif
endif        
close all
人可 = '00'
retu


