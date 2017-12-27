*******************************
* .\grzhcl.PRG(Build20060321)
*******************************
close all
set talk off
clear
_SCREEN.PICTURE = ''
平均2=0
use lnk
copy to temp stru
close all
use fox\grzh excl
if file("d:\grzh.dbf")
   zap
   appe from d:\grzh
endif 
go top
loca for aae001=1995
if !eof()
yn = MESSAGEBOX("是否导入“云南社保系统”中缴费基数和月数？",4+32,"提示")
IF yn = 6
   人可 = '卡片'
ENDIF   
endif 
*************(1)初始化缴费基数、月数*****************
IF 人可 = '卡片'
******自动输入******
use fox\grzh
sort on aae001 to temp1 for allt(grbh)=bh
sele 2
use temp1
go top
sele 1
use temp 
年代 = 1995
for I = 1 to val(LY)-1994
    if allt(whtj)='ryzk'
       appe from ryzk for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
    else
       appe from ryzk1 for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
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
enddo
ELSE
******人工输入******
use temp 
年代 = 1995
for I = 1 to val(LY)-1994
   if whtj='ryzk'
       appe from ryzk for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
    else
       appe from ryzk1 for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
    endif 
    replace 年份 with str(年代,4)
    年代 = 年代+1
endfor
replace 缴费月数 with 6 for 年份='1995'
replace 缴费月数 with 12 for 年份<>'1995'
replace 缴费月数 with 缴费月数-3 for 年份='2003'and 缴费月数>3
replace 月缴费基数 with 0 for val(年份)<val(LY)
go bott
replace 缴费月数 with month(date()) , 说明 with 年份+'年实缴月份数'
ENDIF
******审核“缴费月数 , 月缴费基数”******
replace 说明 with str(val(年份)-1,4)+'年月平均工资' all
close all
use temp excl
go top
clear
reca all
delete for '-1'$说明
pack
@ 36 , 90 say '输入提示' font '宋体',14
@ 38 , 70 say '(1)请输末年缴费月数,2001年直接输2000年月平均工资' font '宋体',12
@ 40 , 70 say '(2)打印“跨行业统筹转移基金说明”从1998年输入' font '宋体',12
@ 42 , 70 say '(3)打印“个人帐户登记卡”从本人缴费起始年输入' font '宋体',12
 browse for left(XM,8)=left(RYXM,8) or left(XM,5)=left(RYBH,5) field;
 RYBH :R : 8 :H = '编  号' , RYXM :R : 8 :H = '姓  名' , 年份 :R ;
,缴费月数 , 月缴费基数 , 说明:R titl '养老保险手册'
go top
loca for 年份='2002' and (allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH))
if !eof()
   平均2=月缴费基数
endif
delete for 月缴费基数=0
pack
use ylbxk
copy to ylgrzh stru 
use ylgrzh 
appe from temp
*************(2)自动或人工输入首年帐户(当年数据)*****************
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
LL = b.当年利率‰
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
 replace 缴费基数 with 月缴费基数*缴费月数 , 月单位缴费 with round(月缴费基数*DW/100,1) , 单位缴费;
 with 月单位缴费*缴费月数 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费;
 with round(月缴费基数*GR/100,1) , 个人缴费 with 月个人缴费*缴费月数 , 比例 with GR;
 ,单位利息 with round((单位缴费+社平5)*(LL/1000)*缴费月数/2,1);
 , 当年利息 with round(个人缴费*(LL/1000)*缴费月数/2,1) , 当年合计;
 with 单位缴费+社平5+单位利息+个人缴费+当年利息,总累计 with 当年合计
ENDIF
******人工输入******
close all 
use ylgrzh excl
clear
@ 38 , 70 say '(1)打印“跨行业统筹转移基金说明”最早输入1998年当年数据(不输历年数据)' font '宋体',12
@ 40 , 70 say '(2)打印“个人帐户登记卡”输入本人缴费起始年数据' font '宋体',12
go top
browse part 12 field 年份: r ,rybh:r:h='编号',RYXM:r:h='姓名' ;
 , 缴费月数 : 8 :H = '缴费月数' , 缴费基数 : 8 :H = '缴费基数';
 , 单位缴费 : 12 :H = '当年单位缴费' , 社平5 : 7 :H = '社平5%';
 , 单位利息 : 12 :H = '当年单位利息' , 比例 :h ='个人缴费比例%' , 个人缴费 : 12 :H = '当年个人缴费';
 , 当年利息 : 12 :H = '当年个人利息' , 当年合计 : 10 :H = '当年合计';
 , 历年单位缴 : 12 :H = '历年单位缴费' , 历年利息1 : 12 :H =;
 '历年单位利息' , 历年个人缴 : 12 :H = '历年个人缴费' , 历年利息2;
 : 12 :H = '历年个人利息' , 总累计 : 10 :H = '总    计' titl  '养老保险个人帐户登记卡 按 [ Ctrl+T ] 删除' 
 
*************(3)计算首年帐户(当年数据)***************** 
单位缴=单位缴费
社平=社平5
单位息=单位利息
历年单位=历年单位缴
历息1=历年利息1
个人缴=个人缴费
当年息=当年利息
历年个人=历年个人缴
历息2=历年利息2
总计=总累计
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
skip
do while !eof()
年=val(年份)
if 年<=1997
  SN = b.上年社平￥
else
  SN = 0
endif
LL = b.当年利率‰
LNLL = b.历年利率％
DW = b.单位缴纳％
GR = b.个人缴纳％
PJGZ = b.上年社平￥
 replace 缴费基数 with 月缴费基数*缴费月数 , 月单位缴费 with round(月缴费基数*DW/100,1) , 单位缴费;
 with 月单位缴费*缴费月数 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费;
 with round(月缴费基数*GR/100,1) , 个人缴费 with 月个人缴费*缴费月数 , 比例 with GR
if 年份='2001'
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
  if 年份='2002'
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
   replace 单位利息 with round((单位缴费+社平5)*(LL/1000)*缴费月数/2,1);
 , 当年利息 with round(个人缴费*(LL/1000)*缴费月数/2,1) , 当年合计;
 with 单位缴费+社平5+单位利息+个人缴费+当年利息 , 历年单位缴 with 单位缴+社平+单位息+历年单位+历息1;
 , 历年利息1 with round(历年单位缴*(LNLL/100),1) , 历年个人缴 with;
 个人缴+当年息+历年个人+历息2 , 历年利息2 with round(历年个人缴*(LNLL/100);
,1) , 总累计 with 当年合计+历年单位缴+历年利息1+历年个人缴+历年利息2;
 , 当年缴纳 with 单位缴费+个人缴费 , 累计利息 with 单位利息+当年利息+历年利息1+历年利息2;
 , 历年累计 with 总计 , 上年社平 with PJGZ , 基础养老金 with round(上年社平*0.2;
,1) , 账户养老金 with round(总累计/120,1) , 养老待遇Ⅰ with round(基础养老金+账户养老金;
,1) , 个人本息 with 个人缴费+当年利息+历年个人缴+历年利息2 , 历年本息 with 历年个人缴+历年利息2
单位缴=单位缴费
社平=社平5
单位息=单位利息
历年单位=历年单位缴
历息1=历年利息1
个人缴=个人缴费
当年息=当年利息
历年个人=历年个人缴
历息2=历年利息2
总计=总累计
skip
enddo
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
clear
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
yn = MESSAGEBOX("需要打印或转换使用“历年帐户”吗？",4+32,"提示")
IF yn = 6
copy to 历年帐户 type xl5 fiel 年份,缴费月数,月缴费基数,缴费基数,单位缴费,单位利息,比例,个人缴费,当年利息,当年合计,历年单位缴,历年利息1,历年个人缴,历年利息2,总累计
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("d:\simis\历年帐户.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"   
ENDI    
if !file("历年帐户.DBF")
   copy to 历年帐户 type fox2x
else
   use 历年帐户 excl
   go top
   loca for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
   if !eof()
      delete for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
      pack
      appe from ylgrzh
   else
       appe from ylgrzh
   endif
endif        
close all
人可 = '00'
delete file temp.dbf
delete file temp1.dbf
clear
 _SCREEN.WINDOWSTATE = 2
close all
use data\ccrq
go 8
repl 图片 with str(val(图片)+1,2)
if allt(图片)='10'
   repl 图片 with '0'
endi   
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
_SCREEN.MAXBUTTON = .F.
_SCREEN.MOVABLE = .F.
_SCREEN.CONTROLBOX = .F.
_SCREEN.CLOSABLE = .F.
retu


