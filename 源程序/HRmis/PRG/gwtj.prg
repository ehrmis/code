***********
* gwtj.prg
***********
dhgw=''
dh14=''
dh1a=''
dhgwdy=''
YF1=month(date())
ND1=year(date())
ND = str(ND1,4)
if YF1>9
  YF = str(YF1,2)
else
  YF = '0'+str(YF1,1)
endif
NY = ND+YF
********当前系统年月*************
NYR = Right(NY,4)
YFs = iif(val(yf)>9,str(val(yf)-1,2),'0'+str(val(yf)-1,1))
IF val(yf)=1
   NY1 = str(year(date())-1,4)+'12'
   YFs = '12'
else   
   if val(yf)=10
      NY1 = str(year(date()),4)+'09'
   ELSE
      NY1 = str(year(date()),4)+YFs
    endif   
ENDIF
********当前系统年月[上月表达式]*************
if !file('gw'+ny+'.dbf')
=MESSAGEBOX('请先审核维护本月工种岗位！', 16,"提示")
 return
endif
clear
do dhk14
do case
case dh14=1
******************************统计报告*********************
use gw&NY
coun for sfgj>0 to 上浮
use 岗位统计
COPY TO temp FOR gwdm<>'lw'
***************不含劳务派遣工岗位分析*********************
USE temp
sum 人数 to rs
sum val(gwdm)*人数 to gs
sum 人数 for '操作'$gwfl1 to cz
sum 人数 for '技术'$gwfl1 to js
sum 人数 for '管理'$gwfl1 to glry
go top
brow field 序号,gwfl1:h='岗位分类',gwjb1:h='上浮前级别',gwdc1:h='档级',岗等:h='岗位等级',bzgz:h='上浮后岗位工资',人数:4 noedit titl '一;
、岗位审核(操作岗位:'+str(cz,3)+'人 管理岗位:'+str(glry,3)+'人 技术岗位:'+str(js,3)+'人  合计:'+str(rs,4)+'人 平均:'+str(round(gs/rs,1),4,1)+'岗 上浮:'+str(上浮,3)+'人 )'
use 操作分类
COPY TO temp FOR gwdm<>'lw'
***************不含劳务派遣工岗位分析*********************
USE temp
repl 序号 with str(recn(),4) all
sum 人数 to zcc for gwfl1='操作岗位'
sum 人数 to jss for gwfl1='技术岗位'
go top
brow field 序号,gwfl1:h='岗位分类',gwjb1:h='上浮前级别',gwdc1:h='档级',岗等:h='岗位等级',bzf:h='注册班长',gwgz:h='上浮后岗位工资',人数:4 noedit titl '二、持证操作人员分类('+str(zcc+jss,4)+'人):操作岗位('+str(zcc,4)+'人) 技术岗位('+str(jss,3)+'人)'
use 无证分类
repl 序号 with str(recn(),4) all
sum 人数 to wzry
go top
brow field 序号,gwfl1:h='岗位分类',gwjb1:h='上浮前级别',gwdc1:h='档级',bzf:h='注册班长',bzgz:h='上浮后岗位工资',人数:4 noedit titl '三、无证人员分类('+str(wzry,3)+'人)'

use 管理分类
repl 序号 with str(recn(),4) all
sum 人数 to gll for gwfl1='管理岗位'
sum 人数 to zyy for gwfl1='技术岗位'
go top
brow field 序号,gwfl1:h='岗位分类',gwjb1:h='上浮前级别',gwdc1:h='档级',岗等:h='行政技术职务',bzgz:h='上浮后岗位工资',人数:4 noedit titl '四、管理人员分类('+str(gll+zyy,3)+'人):管理岗位('+str(gll,3)+'人) 技术岗位('+str(zyy,3)+'人)'

use 工种统计
COPY TO temp FOR gwdm<>'lw'
***************不含劳务派遣工岗位分析*********************
USE temp 
sum a to rs
go top
browse field gz:r:h='代码' , gz1:r:h='工种' ,gw:h='岗位':r,gwjb1:h='岗级':r,岗等:r,a:h='人数':r:4 titl '五、工种岗位分布(合计:'+str(rs,4)+'人)'
if file('培训汇总.dbf')
use 培训汇总
sum a to rs
go top
browse field gwfl1:h='岗位分类':r,gz1:r:h='工种' ,gw:h='岗位':r,pxmc:h='培训名称':r,pxdj:h='培训等级':r,a:h='人数':r:4 titl '六、各工种岗位培训情况(合计:'+str(rs,4)+'人)'
endif
case dh14=2
do dhkdy2
if dh1a=1
   do dhkgwdy
   if dhgwdy=1
      do grgw
   else
      do gbgw
   endi
else
   do 岗位统计
endif
case dh14=3
retu
endcase
close all
clear
delete file gzhz.dbf
delete file gzhz.idx
return
