close all
use sbdmk
if empty(allt(dw)) or empty(allt(mc))
   MESSAGEBOX('请先进行系统维护！！！','提示')
   retu
endif
close all
DO SRNd.SPr
*********当年关联上年、前年*******************
LY1=str(val(LY)-1,4)
qn=str(val(LY)-2,4)
IF !file('zr1'+ly1+'.dbf')
MESSAGEBOX('请先进行'+LY1+'年工资总额处理！！！','提示')
  RETURN
ENDIF
use 基础数据
go top
loca for 年份 = '&LY'
go top
*警告:及时维护当前基础数据*
loca for 年份 = '&LY' 
if 上年社平￥=0 or 当年利率％=0 or 历年利率％=0 or 个人缴纳％=0 or 失业保险％=0 or 系数=0
   DO ly_gzze
endif
**************
do ylbxjs
**************
wait wind   LY+'年养老保险处理成功......'  nowa
**************
do sybxjs
**************
wait wind   LY+'年失业保险处理成功......'  nowa
CLOSE ALL
use zr&LY
DBFF1='zr'
WHNY=LY+'年工资总额'
do Qdir
close all
use zr1&LY
DBFF1='zr1'
WHNY=str(val(LY)+1,4)+'年缴费基数'
do Qdir
CLOSE ALL
if 人可='00'
use YL&LY
DBFF1='YL'
WHNY=LY+'年养老保险'
do Qdir
close all
use YL&LY
copy to LY+'年养老保险'
endif
close all
use SY&LY
DBFF1='SY'
WHNY=LY+'年失业保险'
do Qdir
close all
use SY&LY
copy to LY+'年失业保险'
close all
return