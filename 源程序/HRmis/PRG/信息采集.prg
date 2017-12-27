yn = MESSAGEBOX("是否新用户人员信息批量采集？",4+32,"提示")
IF yn = 7
   RETURN
ENDIF
do srny.spx
sn=STR(VAL(nd)-1,4)
NY = ND+YF
****************前月自动处理****************
if val(yf)<3
  JJND = str(val(nd)-1,4)
  JJYF = str(10+val(yf),2)
else
  JJND = ND
  JJYF = iif(val(YF)-2>9,str(val(YF)-2,2),'0'+str(val(YF)-2,1))
endif
****************上月自动处理****************
if val(YF)-1>9
  jjYF = str(val(YF)-1,2)
else
  jjYF = iif(val(YF)-1=0,'12','0'+str(val(YF)-1,1))
endif
if jjyf='12'
   syjj=str(val(nd)-1,4)+jjYF
else
   syjj = ND+jjYF
ENDIF
yf1=1
yf2=VAL(yf)
****************初始化****************
USE ryzk
BROWSE TITLE '数据采集全部或添加部分人员入库ryzk.dbf维护与测试：1、检查空代码后人工维护；2、批量更新职务代码；3、批量生成zydm'
DO jszc
DO zy
DO zwdm
*****************
PROCEDURE zwdm
**************扫描法快速维护职务代码****特殊维护必须谨慎*****20151126*******************
CLOSE all
sele 2
USE ryzk
scan
 sele 1
 use zwdm
 WAIT WINDOW NOWAIT '正在维护'+b.cjdm+'：'+allt(b.cjmc)+'、'+allt(b.bmmc)+'→'+allt(b.ryxm)+'<职务代码>... ...' 
 loca for ALLTRIM(b.zw1)=allt(name)
 repl b.zw with code     
endscan       

******************
PROCEDURE jszc
******************
**********扫描法快速维护工种岗位************
CLOSE all
sele 2
USE ryzk
scan
 sele 1
 use jszc
 WAIT WINDOW NOWAIT '正在维护'+b.cjdm+'：'+allt(b.cjmc)+'、'+allt(b.bmmc)+'→'+allt(b.ryxm)+'<技术职称>... ...' 
 IF !EMPTY(b.zcbh)
    loca for val(b.zcbh)=val(code)
    repl b.zc with allt(zc)
 ELSE    
    loca for  ALLTRIM(b.zc)=allt(zc)
    repl b.zcbh with code
    ***********技术职称自动编号*******20151105***************
 ENDIF    
endscan
******************
PROCEDURE zy
******************
**********扫描法快速维护所学专业************
CLOSE all
sele 2
USE ryzk
repl zydm with '',zymc with '',zyfl with '', kbdm with '',kbzy with '',a with 1 all
scan
 sele 1
 use zyfl
 WAIT WINDOW NOWAIT '正在维护'+b.cjdm+'：'+allt(b.cjmc)+'、'+allt(b.bmmc)+'→'+allt(b.ryxm)+'<专业分类>... ...' 
 ****第一专业*******
 loca for b.ygxz='01' and (allt(专业分类)$allt(b.zy) or allt(专业分类1)$allt(b.zy) or allt(专业分类2)$allt(b.zy))
 repl b.kbdm with allt(代码),b.kbzy with allt(专业分类)
 if empty(专业分类1) and  !empty(专业分类2)
    repl b.kbdm with allt(代码),b.kbzy with allt(专业分类2)
 endif
 ****第二专业*******
 if !empty(b.zy1)
    loca for allt(专业分类)$allt(b.zy1) or allt(专业分类1)$allt(b.zy1) or allt(专业分类2)$allt(b.zy1)
    repl b.zydm with allt(代码),b.zyfl with allt(专业分类)
    if empty(专业分类1) and  !empty(专业分类2)
       repl b.zydm with allt(代码),b.zyfl with allt(专业分类2)
    endif   
 endif 
endscan
repl kbdm with '88' for val(kbdm)=0
repl zydm with '99' for val(zydm)=0
