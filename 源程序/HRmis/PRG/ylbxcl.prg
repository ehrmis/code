***************************
*YLBXCL.prg(Build20151103)
***************************
人可 = '00'
平均2=0
ly=nd
LY1 = str(val(LY)-1,4)
if not file('yl'+LY1+'.DBF')
    MESSAGEBOX('请先生成'+LY1+'年养老保险！！','提示')
    return
ENDIF
use yl&ly1
if recc()=0
   MESSAGEBOX('请先生成'+LY1+'年养老保险！！','提示')
   return
endif
use ylbxk
COPY TO yltemp stru  
 USE yltemp  
 appe from ryzk  
replace 年份 with LY , A with 1 all
close all
use grzh excl
lj="&xp"+"\backup\"
if file("&lj"+"grzh.dbf")
   zap
   appe from &lj.grzh
endif 
go top
loca for aae001=val(ly)
if !eof()
yn = MESSAGEBOX("是否导入“云南社保系统”中缴费基数和月数？",4+32,"提示")
IF yn = 6
   人可 = '卡片' 
ENDIF   
endif 
use 基础数据
go top
loca for 年份='&LY'
PJGZ = 上年社平￥
LNLL = 历年利率％
DW = 单位缴纳％
GR = 个人缴纳％
GRSY = 失业保险％
年=val(年份)
XS = 系数
if 年<=1997
  SN = 上年社平￥
else
  SN = 0
endif
LL = 当年利率％/100
if 年<=2006
    LL = 当年利率‰/1000
endif
close all
IF 人可 <> '卡片' 
sele 3
use zr1&ly excl
delete for pj<1
pack
inde on rybh to ZR1&ly
ENDIF
sele 2
IF 人可 = '卡片'
use grzh
sort on aae001 to temp for aae001=val(LY)
use temp
inde on allt(grbh) to temp
go top
ELSE
if not file('zr1'+LY1+'.DBF')
    MESSAGEBOX('请先进行'+LY1+'年工资总额处理！！！','提示')
    return
endif
use zr1&ly1 excl
delete for pj<1
pack
inde on rybh to ZR1&ly1
ENDIF
sele 1
use yltemp
**************************************************
IF 人可 = '卡片'
set relation to allt(grbh) into 2
go top
do while !eof() 
wait window '正在填写缴费月数、基数：'+allt(ryxm)+'（'+allt(str(recn(),5))+'／'+allt(str(recc(),5))+'）......' nowait   
    sele 2
    sum aic079 to ys for aae001=val(LY) and allt(grbh)=allt(a.grbh)
    go top
    loca for aae001=val(LY) and allt(grbh)=allt(a.grbh)   
    if !eof()
       replace a.缴费月数 with ys,a.月缴费基数 with jfjs
    endif
    continue
    if !eof()
       平均2=jfjs       
    endif
    sele 1
replace 缴费基数 with 月缴费基数*缴费月数;
  , 月单位缴费 with round(月缴费基数*DW/100,1) , 单位缴费;
 with 月单位缴费*缴费月数 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费;
 with round(月缴费基数*GR/100,1) , 个人缴费 with 月个人缴费*缴费月数 
  if 年=2001 
    do case
    case 缴费月数=1
      replace 单位缴费 with round(平均2*4/100,1) , 个人缴费 with round(平均2*7/100,2) , 缴费基数 with 平均2
    case 缴费月数=2
      replace 单位缴费 with round(平均2*4/100,1)*2 , 个人缴费 with round(平均2*7/100,2)*2 , 缴费基数 with;
 平均2*2
    case 缴费月数=3
      replace 单位缴费 with round(平均2*4/100,1)*3 , 个人缴费 with round(平均2*7/100,2)*3 , 缴费基数 with;
 平均2*3
    case 缴费月数>3
      replace 单位缴费 with 月单位缴费*(缴费月数-3)+round(平均2*4/100,1)*3 , 个人缴费 with;
 月个人缴费*(缴费月数-3)+round(平均2*7/100,2)*3 , 缴费基数 with 月缴费基数*(缴费月数-3)+平均2*3
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
  skip
enddo
repl 单位利息 with round((单位缴费+社平5)*LL*缴费月数/2,1);
 , 当年利息 with round(个人缴费*LL*缴费月数/2,2) , 当年合计 with 单位缴费+社平5+单位利息+个人缴费+当年利息 all
ELSE
wait window '正在处理'+LY+'年个人帐户......' nowait 
set relation to rybh into 3
repl 缴费月数 with c.MOU all
repl 缴费月数 with 0 for 缴费月数<0
index on RYBH to yltemp
 update on RYBH from B replace 月缴费基数 with b.jfjs;
 , 缴费基数 with 月缴费基数*缴费月数 , 月单位缴费 with b.ZR8 , 单位缴费;
 with 月单位缴费*缴费月数 , 社平5 with round(SN*5/100,1)*缴费月数 , 月个人缴费;
 with b.ZR3 , 个人缴费 with 月个人缴费*缴费月数
go top
do while  not eof()
  if val(LY)=2001  
  wait window '正在处理2001年缴费基数：'+allt(ryxm)+'（'+allt(str(recn(),5))+'／'+allt(str(recc(),5))+'）......' nowait 
    do case
    case 缴费月数=1
      replace 单位缴费 with c.ZR8 , 个人缴费 with c.ZR3 , 缴费基数 with c.jfjs
    case 缴费月数=2
      replace 单位缴费 with c.ZR8*2 , 个人缴费 with c.ZR3*2 , 缴费基数 with;
 c.jfjs*2
    case 缴费月数=3
      replace 单位缴费 with c.ZR8*3 , 个人缴费 with c.ZR3*3 , 缴费基数 with;
 c.jfjs*3
    case 缴费月数>3
      replace 单位缴费 with 月单位缴费*(缴费月数-3)+c.ZR8*3 , 个人缴费 with;
 月个人缴费*(缴费月数-3)+c.ZR3*3 , 缴费基数 with 月缴费基数*(缴费月数-3)+c.jfjs*3
    endcase
  else
    exit
  endif
  skip
enddo
go top
do while  not eof()
  if  val(LY)=2002
   wait window '正在处理2002年缴费比例：'+allt(ryxm)+'（'+allt(str(recn(),5))+'／'+allt(str(recc(),5))+'）......' nowait 
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
  else
    exit
  endif
  skip
enddo
go top
do while  not eof()
  if val(LY)=2003
    do case
    case 缴费月数=1
     replace 单位缴费 with 月单位缴费 , 个人缴费 with 月个人缴费
    case 缴费月数=2
     replace 单位缴费 with 月单位缴费*2 , 个人缴费 with 月个人缴费*2
    case 缴费月数=3
     replace 单位缴费 with 月单位缴费*3 , 个人缴费 with 月个人缴费*3
    case 缴费月数>3
     replace 缴费月数 with 缴费月数-3, 单位缴费 with 月单位缴费*缴费月数 , 个人缴费 with;
 月个人缴费*缴费月数
    endcase
    repl 缴费基数 with 月缴费基数*缴费月数
  else
    exit
  endif
  skip
enddo
repl 单位利息 with round((单位缴费+社平5)*LL*缴费月数/2,1);
 , 当年利息 with round(个人缴费*LL*缴费月数/2,1) , 当年合计 with 单位缴费+社平5+单位利息+个人缴费+当年利息 all
ENDIF
close all
select 2
use yl&ly1
index on RYBH to yl&ly1
sele 1
use yltemp excl
index on RYBH to yltemp
*******************结算、计算历年本息****体会：编程语法严谨性******************
 IF val(LY)>=2007
 replace 当年利息 with round(个人缴费*LL*XS*1/2,2) , 当年合计 with 个人缴费+当年利息 , 总累计 with  当年合计  all
 upda on rybh from B replace 总累计 with  round(b.总累计*(1+LNLL/100),2)+当年合计 , 历年累计 with b.总累计 
 repl  当年缴纳 with 个人缴费 , 累计利息 with 当年利息 , 上年社平 with PJGZ , 基础养老金 with round(上年社平*0.2;
,1) , 账户养老金 with round(总累计/120,2) , 养老待遇Ⅰ with round(基础养老金+账户养老金;
,1) , 个人本息 with 个人缴费+当年利息 all
****************新增员工**************
 replace 总累计 with  round(总计*(1+LNLL/100),2)+当年合计 , 历年累计 with 总计 for 总计>0
**************************************
inde on bmbh+rybh to yl&ly
ELSE
upda on rybh from B replace 单位利息 with round((单位缴费+社平5)*LL*缴费月数/2,2);
 , 当年利息 with round(个人缴费*LL*缴费月数/2,2) , 当年合计;
 with 单位缴费+社平5+单位利息+个人缴费+当年利息 , 历年单位缴 with b.单位缴费+b.社平5+b.单位利息+b.历年单位缴+b.历年利息1;
 , 历年利息1 with round(历年单位缴*(LNLL/100),2) , 历年个人缴 with;
 b.个人缴费+b.当年利息+b.历年个人缴+b.历年利息2 , 历年利息2 with round(历年个人缴*(LNLL/100);
,1) , 总累计 with 当年合计+历年单位缴+历年利息1+历年个人缴+历年利息2;
 , 当年缴纳 with 单位缴费+个人缴费 , 累计利息 with 单位利息+当年利息+历年利息1+历年利息2;
 , 历年累计 with b.总累计 , 上年社平 with PJGZ , 基础养老金 with round(上年社平*0.2;
,1) , 账户养老金 with round(总累计/120,2) , 养老待遇Ⅰ with round(基础养老金+账户养老金;
,1) , 个人本息 with 个人缴费+当年利息+历年个人缴+历年利息2 , 历年本息 with 历年个人缴+历年利息2
ENDIF
******审核个人帐户处理结果******
go top
browse part 12 field 年份: r ,rybh:r:h='编号',RYXM:r:h='姓名' ;
 , 缴费月数 : 8 :H = '缴费月数' , 缴费基数 : 8 :H = '缴费基数';
 , 单位缴费 : 12 :H = '当年单位缴费' , 社平5 : 7 :H = '社平5%';
    , 单位利息 : 12 :H = '当年单位利息' , 比例 :h ='个人缴费比例%' , 个人缴费 : 12 :H = '当年个人缴费';
 , 当年利息 : 12 :H = '当年个人利息' , 当年合计 : 10 :H = '当年合计';
 , 历年单位缴 : 12 :H = '历年单位缴费' , 历年利息1 : 12 :H =;
 '历年单位利息' , 历年个人缴 : 12 :H = '历年个人缴费' , 历年利息2;
 : 12 :H = '历年个人利息' , 总累计 : 10 :H = '总    计'
delete for 缴费月数=0
pack
if !file("历年帐户.DBF")
   copy to 历年帐户 type fox2x
else
  use 历年帐户 excl
  delete for val(年份)=val(LY)  
  pack  
  appe from yltemp 
endif 
sort on bmbh,rybh,年份 to temp
use temp
*********校正***********
loca for val(年份)=2002
if !eof()
gz_21=月缴费基数
skip -1
if val(年份)=2001
   gz_20=月缴费基数
   REPLACE 月缴费基数 with round((gz_20*9+gz_21*3)/12,0)
endif
continue
endif
************************
copy to 历年帐户 type fox2x
use 历年帐户
repl 比例 with GR all 
yn = MESSAGEBOX("需要打印或转换使用“历年帐户”吗？",4+32,"提示")
IF yn = 6
copy to &bf.\历年帐户 type xl5
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\历年帐户.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"   
ENDI   
close all
人可 = '00'
retu 