*******************************
*岗位管理.prg (Build 20170719)
*******************************
set escape off
dhgw=''
dh14=''
dh1a=''
dhgwdy=''
on key label F1 do grcx
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
wait window '处  理  数  据　中, 请  稍  候... ...' nowait
CLOSE all
USE ryzk
*********必须是当前ryzk与ry&ny无关*****20170320**************************
COPY TO ryzktemp
USE ryzktemp EXCLUSIVE
COUNT for year(离岗时间)=<year(date()) and year(离岗时间)>0 TO lgrs
IF lgrs>0 
DELETE for year(离岗时间)=<year(date()) and year(离岗时间)>0
BROWSE FIELDS dwmc,cjmc,bmmc,ryxm,离岗时间 for year(离岗时间)=<year(date()) and year(离岗时间)>0 TITLE '请审核是否删除离岗人员？'
*****************************应删除人员*************************************** 
yn = MESSAGEBOX("是否删除离岗人员？",4+32,"提示")
IF yn = 6
PACK
ELSE
RECALL
ENDIF
ENDIF
******及时清理预删除人员**************
close all
use 岗位管理 excl
zap
appe from ryzktemp
repl 岗等 with zcdj1 for allt(gwfl1)='操作岗位' 
repl gwgz with bzgz for allt(ygxz1)='合同制职工' 
repl 人数 with 1 , A with 1 all
IF !files("gw"+ny1+".dbf")
   COPY to gw&NY1
ENDIF  
COPY to gw&NY
close all
wait window '处  理  数  据 , 请  稍  候... ...' nowait
sele 1
use 岗位管理
repl 岗等 with '' all
go top
sele 2
use gwdmk
SUM 自定义 TO zdy
scan
 sele 1 
 repl gwdm with b->gwdm for allt(gwfl1)='操作岗位' and gwjb = b->gwdm
 repl gwdm with b->gwdm for allt(gwfl1)='技术岗位' and gwjb = b->gwdm
 repl gwdm with b->gwdm for allt(gwfl1)='管理岗位' and gwjb = b->gwdm
 IF zdy=0
 ******单位自定义工资栏有值否****20170322**************
 repl 岗等 with b->gd03 for allt(gwfl1)='技术岗位' and bzgz = b->bzgz2014
 repl 岗等 with b->gd04 for allt(gwfl1)='管理岗位' and bzgz = b->bzgz2014
 ************************************单位自定义工资标准******20170322********************************
 ELSE 
  repl 岗等 with zcdj1 for allt(gwfl1)='技术岗位'
  repl 岗等 with zw1 for allt(gwfl1)='管理岗位'
 ************************************集团工资标准******20170322********************************
 ENDIF 
ENDSCAN
********************岗等：单位自定义与集团工资制度的标准职称工资要区分开来********20170112***************
close all
********************按bzgz2014版本校正岗等*******20170322****************
use 岗位管理
repl 岗等 with '业务员' for allt(gwfl1)='管理岗位' and gwdm = '06'
repl 岗等 with '业务主办' for allt(gwfl1)='管理岗位' and gwdm = '07'
repl 岗等 with '业务主管' for allt(gwfl1)='管理岗位' and gwdm = '08'
repl 岗等 with '业务副经理' for allt(gwfl1)='管理岗位' and gwdm='09' and zw='07'
repl 岗等 with '业务经理' for allt(gwfl1)='管理岗位' and gwdm='11' and zw='06'
repl 岗等 with '高级业务经理' for allt(gwfl1)='管理岗位' and gwdm='13' 
repl 岗等 with '副总经理' for allt(gwfl1)='管理岗位' and gwdm='15' and zw='03'
repl 岗等 with '总经理' for allt(gwfl1)='管理岗位' and gwdm='17' and zw='02'
***********************Ｂ.审核维护*****************
inde on bmbh+zw+rybh to 岗位管理
for zdz=1 to fcount()
  zd=fiel(zdz)
  zdz1=(fsize(zd)+fsize(zd))*2
  if uppe(allt(zd))='RYXM' 
     exit
  endif
endfor   
go top
browse part zdz1 field dwmc:h='单位':r,cjmc:h='部门':r,bmmc:h='班组':r,ryxm:h='姓名':r,zw1:r:h='行政职务',bzf:h='注册班长(Y)',gz:h='工种代码',;
gz1:h='工种':r,gw:h='岗位':r,gwjb1:r:h='等级',gwdc1:h='岗位档差':r,gwgz:h='岗位工资':r,岗等:h='现聘岗位等级',sfgj:h='上浮岗级':r,;
bzgz:h='标准工资':r,gwdm:h='岗等代码',gl:h='工龄':r,pxdj:h='培训等级':r,zcdj1:r:h='技术职务',zc:r:h='职称',gwfl:h='岗位分类码':r,;
gwfl1:h='岗位分类':r,zymc:h='职业发展' titl '[按 F1 查个人] [按 Esc 关闭] 请认真审核和维护：①注册班长②工种代码③现聘岗位等级'
repl gwfl1 with '管理岗位' for gwfl='01'
repl gwfl1 with '技术岗位' for gwfl='02'
repl gwfl1 with '操作岗位' for gwfl='03'
CLOSE all
do 检查
do 培训
close all
sele 2
use 岗位管理
inde on rybh to 岗位管理
sele 1
use ryzk &&保存更新数据
inde on rybh to ryzk
upda on rybh from B repl 岗等 with b.岗等,bzf with b.bzf,gz with b.gz
**************************Ｃ.检测整理***********************批量更新工种岗位名称20170320*****************************
close all
use 岗位管理
inde on gwjb+gwdc to 岗位管理
tota on gwjb+gwdc to 岗级统计
copy to gw&NY
use gw&NY
repl 人数 with 1 , A with 1 all
inde on gz to gw&NY
tota on gz to temp
copy to 管理人员 for val(gwfl)<3
******************************原管理人员(2003年6月)**************
copy to 操作人员 for val(gwfl)=3
use 操作人员 excl
copy to 无证人员 for empty(pxdj) and '级工'$zcdj1
delete for empty(pxdj) and gwfl='03'
pack
inde on gwfl1+岗等+bzf to 操作人员
tota on gwfl1+岗等+bzf to 操作分类
use 无证人员 excl
inde on gwfl1+岗等+bzf to 无证人员
tota on gwfl1+岗等+bzf to 无证分类
use 管理人员
inde on gwfl1+岗等+岗等 to 管理人员
tota on gwfl1+岗等+岗等 to 管理分类
close all
yn = MESSAGEBOX("岗位变动及岗位调整趋势报表导出成功！现打开使用吗？",4+32,"提示")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\岗位变动.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\标准调整趋势.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\平均调整趋势.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
ENDIF

proc 培训

*****按每人的最高级别排序保存后统计培训情况****
use ryzk
repl a with 1 all
copy to temp for !empty(pxdj) fiel cjmc,bmmc,rybh,ryxm,gz1,gw,pxmc,pxdj,zcfl1,zcdj1,gwfl1,a
use temp
copy to temp1 fiel gwfl1,gz1,gw,pxmc,pxdj,a
use temp1
inde on gwfl1+gz1+gw+pxmc+pxdj to temp1
tota on gwfl1+gz1+gw+pxmc+pxdj to 培训统计
close all

proc 检查
close all
select 1
use gw&NY
repl a with 1 all
select 2
use gw&NY1
scan
wait window '检 索 工 种 代 码 岗 位 标 准,请 稍 候... ...' nowait
  select 1
  repl a with 0 for allt(rybh)=allt(b.rybh) and bzgz<b.bzgz &&降岗
  repl a with 2 for allt(rybh)=allt(b.rybh) and bzgz>b.bzgz &&升岗
endscan
close all
***************************
use gongzong EXCLUSIVE
yn = MESSAGEBOX("需要维护工种代码库吗？",4+32,"提示")
IF yn = 6
   APPEND BLANK
go top
brow fiel code:h='工种代码',name:h='工种',gwname:h='岗位',gwdj:h='岗位等级':r,bzgz:h='标准工资':r,pj:h='平均岗资':r,num:h='在岗人数':r titl '当前工种岗位自定义代码及其标准命名（保存后将按标准命名自动更新岗位从业人员工种岗位名称）'
************管理员专业维护代码（按单位工种岗位实际情况统一维护），只读属性，不支持修改代码****20160829******
DELETE FOR EMPTY(code)
pack
repl a with 1 all
sort on code to temp
close all
use temp excl
go top
do while !eof()
   bh=code
   skip
   if bh=code
      repl a with 0
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a=0 to aaa
if aaa>0
   inde on code to temp
   go top
   brow for a=0 titl '发现'+allt(str(aaa,4))+'个工种代码重号！请修改正确或删除多余记录' 
endif
repl a with 1 all
COPY TO data\gongzong
************新增或修改工种代码后及时排序保存****20170322*******************
ENDIF
CLOSE all
USE ryzk
REPLACE zyflm WITH  '',zymc with '' all 
use zydm EXCLUSIVE
go top
brow fiel 代码,职业,工种,岗位,职称 titl '请审核维护当前职业大典中工种岗位及职业发展数据库'
************管理员专业维护代码（按单位工种岗位实际情况统一维护），只读属性，不支持修改代码****20160829******
pack
repl a with 1 all
sort on 代码 to temp1
close all
use temp1 excl
go top
do while !eof()
   bh=代码
   skip
   if bh=代码
      repl a with 0
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a=0 to aaa
if aaa>0
   inde on 代码 to temp1
   go top
   brow for a=0 titl '发现'+allt(str(aaa,4))+'个代码重号！请修改正确或删除多余记录' 
endif
repl a with 1 all
sort on 代码 to data\zydm
CLOSE all
sele 1
use ryzk
scan
 sele 2
 use zydm
 loca for allt(工种)$allt(A.gz1) or allt(岗位)$allt(A.gw) or allt(职称)$allt(A.zc)
 WAIT WINDOW NOWAIT '正在维护'+A.cjdm+'：'+allt(A.cjmc)+'、'+allt(A.bmmc)+'→'+allt(A.ryxm)+'<职业发展>... ...' 
 repl A.zyflm with allt(代码),A.zymc with allt(职业)
 ***************规划职业生涯*********20160830******************************
 use gongzong
 loca for allt(code)=allt(A.gz)
 WAIT WINDOW NOWAIT '正在维护'+A.cjdm+'：'+allt(A.cjmc)+'、'+allt(A.bmmc)+'→'+allt(A.ryxm)+'<工种、岗位>... ...' 
 repl A.gz1 with allt(name),A.gw with allt(gwname)
 ***************1.标准工种、岗位名称维护2.对批量修改工种代码的工种、岗位名称自动更新成工种代码对应的工种岗位**************20170320************
endscan
close all
use gw&NY
REPLACE  a WITH 1 all
INDEX on gz TO gw&NY
tota on gz TO gztemp
SELECT 3
USE gwdmk
INDEX on gwdm TO gwdmk
select 1
use gongzong
REPLACE gwdm WITH '',gwdj WITH '',pj WITH 0,bzgz WITH 0,num WITH 0 all
**********用少记录数表更新多记录数表时必须先清空被更新表*********20170322**********
select 2
use gztemp EXCLUSIVE
DELETE FOR EMPTY(gz)
PACK
scan
  WAIT WINDOW NOWAIT '正在统计工种岗位人数... ...'
  select 1
  LOCATE for code=b.gz
  repl gwdm with b.gwjb, pj with b.bzgz/b.a,num with b.a 
  ******************************************以目前平均岗位工资判断定岗适合度或调整趋势分析******20170719*******
 WAIT WINDOW NOWAIT '正在维护'+gwdm+'：'+allt(name)+'→'+allt(gwname)+'<工种代码库>... ...' 
 repl gwdm with b.gwjb
 SET RELATION TO gwdm INTO 3
 repl gwdj with c.gd01,bzgz WITH c.bzgz 
 ****单位岗位等级代码维护**201719****
ENDSCAN
****************************
close all
select 1
use gw&NY
repl a with 1 all
inde on gwjb to gw&NY
tota on gwjb to gwjbtemp fiel gwjb,a
use gwjbtemp
select 2
use gwdmk
scan
  WAIT WINDOW NOWAIT '正在统计岗位级别... ...'
  select 1
  repl b.num with a for gwjb=b.gwdm
endscan
close all
use gw&NY
repl a with 1 all
copy to 调整趋势
select 1
use 调整趋势 excl
delete for bzgz=0
pack
select 2
use gongzong
scan
wait window '正在检索工种岗位与岗位标准不相符数据,请稍候... ...' nowait
  select 1
  LOCATE for allt(gz)=allt(b.code) and bzgz<b.bzgz
  *****************BUG扫描体中必须用LOCATE for 而不是for *********20170719********************************
  repl a with 0
   LOCATE for allt(gz)=allt(b.code) and bzgz=b.bzgz
  *****************BUG扫描体中必须用LOCATE for 而不是for *********20170719********************************
  repl a with 1
  LOCATE for allt(gz)=allt(b.code) and bzgz>b.bzgz 
  repl a with 2 
  LOCATE for allt(gz)=allt(b.code) and bzgz<b.pj
  *****************BUG扫描体中必须用LOCATE for 而不是for *********20170719********************************
  repl a with 3
  LOCATE for allt(gz)=allt(b.code) and bzgz=b.pj
  repl a with 4 
   LOCATE for allt(gz)=allt(b.code) and bzgz>b.pj
  *****************BUG扫描体中必须用LOCATE for 而不是for *********20170719********************************
  repl a with 5
  repl gwgz with b.pj for allt(gz)=allt(b.code)
endscan
close all
use 调整趋势 excl
coun to xy for a=0
coun to dy for a=2
coun to xy1 for a=3
coun to dy1 for a=5
IF xy+dy+xy1+dy1>0
REPLACE gz WITH '大于' FOR a=2 
REPLACE gz WITH '等于' FOR a=1 
REPLACE gz WITH '小于' FOR a=0 
go top
 brow part 35 for a=0 or a=2 fiel  cjmc:r:h='车间',bmmc:r:h='部门',ryxm:r:h='姓名',gz:r:h='趋势',gz1:r:h='工种',gw:r:h='岗位',gwjb1:h='级别':r,;
岗等:r,sfgj:h='上浮岗级':r,gwgz:r:h='平均岗位工资',bzgz:r:h='标准工资',zymc:h='职业发展' titl '职工岗位工资大于本工种岗位标准工资：'+str(dy,3)+'人；职工岗位工资小于本工种岗位标准工资：'+str(xy,3)+'人 共'+str(dy+xy,4)+'人'
copy to &bf.\标准调整趋势 FIELDS rybh,ryxm,cjmc,bmmc,岗等,gwgz,gz,gz1,gw,gwjb1,bzgz type xl5
REPLACE gz WITH '大于' FOR  a=5
REPLACE gz WITH '等于' FOR  a=4
REPLACE gz WITH '小于' FOR  a=3
 GO top
 brow part 35 for a=3 or a=5 fiel cjmc:r:h='车间',bmmc:r:h='部门',ryxm:r:h='姓名',gz:r:h='趋势',gz1:r:h='工种',gw:r:h='岗位',gwjb1:h='级别':r,;
岗等:r,sfgj:h='上浮岗级':r,gwgz:r:h='平均岗位工资',bzgz:r:h='标准工资',zymc:h='职业发展' titl '岗位调整趋势分析：职工岗位工资大于本工种岗位平均工资：'+str(dy1,3)+'人；职工岗位工资小于本工种岗位平均工资：'+str(xy1,3)+'人 共'+str(dy1+xy1,4)+'人'
copy to &bf.\平均调整趋势 FIELDS rybh,ryxm,cjmc,bmmc,岗等,gwgz,gz,gz1,gw,gwjb1,bzgz type xl5
ENDIF
**************岗位调整趋势分析****20170719**********************
close all
use gw&NY
repl a with 1 all
INDEX on bmbh+rybh TO gw&NY
brow part 35 fiel cjmc:r:h='车间',bmmc:r:h='部门',ryxm:r:h='姓名',gz1:r:h='工种',gw:r:h='岗位',gwjb1:h='级别':r,;
岗等:r,gwdc1:h='岗位档差':r,gwgz:r:h='岗位工资',zymc:h='职业发展' titl '员工职业生涯规划（说明：“级别”、“岗等”、“标准工资”是指今后职业发展目标）'
copy to temp 
jg=0
sg=0
 sele 1
   use temp
 sele 2
   use gw&NY1
   inde on rybh to gw&NY1  
scan
wait window '正 在 检 索 本 月 岗 位 变 动 情 况,请 稍 候... ...' nowait
  select 1
  repl a with 0 for allt(rybh)=allt(b.rybh) and bzgz<b.bzgz
  repl a with 2 for allt(rybh)=allt(b.rybh) and bzgz>b.bzgz
endscan
close all
use temp excl
delete for a=1
pack
if recc()>0
人可='变动'
sort on bmbh,rybh to temp1 
   close all
   sele 2
   use gw&NY1
   inde on rybh to gw&NY1
   sele 1 
   use temp1 excl
   repl 序号 with str(recn(),4) all
   count for a=0 to jg
   count for a=2 to sg
   inde on rybh to temp1
   upda on rybh from B repl dwmc with b.gz1,bmmc with b.gw,岗等 with b.gwjb1,gwgz with b.bzgz
   inde on 序号 to temp1
   go top
   brow fiel 序号,ryxm:h='姓名',erprybh:h='工号',cjmc:h='单位',dwmc:h='原工种',bmmc:h='原岗位',岗等:h='原岗级',gwgz:h='原岗位工资',gz1:h='现工种',gw:h='现岗位',gwjb1:h='现岗级',bzgz:h='现岗位工资' titl '下列人员岗位本月有变动(共'+str(recc(),4)+'人;升岗'+str(sg,4)+'人,降岗'+str(jg,4)+'人)'
   copy to &bf.\岗位变动 FIELDS 序号,ryxm,erprybh,cjmc,dwmc,bmmc,岗等,gwgz,gz1,gw,gwjb1,bzgz type xl5
ENDIF
close all

proc grcx
do while .t.
      XM = '        '
      do xmsr.spr
      locate for upper(XM)=RYXM or upper(XM)=RYBH
      if eof()
        do dhkjg.spx
    M = readkey()
    if M=12 or M=268
      return
    endif
      else
        exit
      endif
enddo