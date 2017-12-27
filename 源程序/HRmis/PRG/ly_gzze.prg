*****************************************
* .\LY_GZZE.PRG(Build20150825)
*****************************************
LY = str(val(nd)-1,4)
******('缴费年(LY)'与'自然年'转换(ND))
use 基础数据
go top
loca for 年份 = '&nd'
if eof()
  append blank
  replace 年份 with nd , 个人缴纳％ with 8 , 失业保险％ with 0.6 , 医疗保险％ with 2 ,企业年金％ with 10 , 住房公积金 with 12
endif
go top
*警告:及时维护当前基础数据*
loca for 年份 = '&nd' 
if 上年社平￥=0
browse for 年份=str(year(date()),4) titl '请及时输入社会保险基础数据(建议暂用上年“社平工资”计算，待正式公布数据后再修改当年实际官方社会保险数据）'
retu
endif
********************每年6月份按新社会平均工资重新生成当年缴费基数和个人缴费数据***************************************
GR = 个人缴纳％/100
SN = 上年社平￥
DQ = 地区社平￥
SY = 失业保险％/100
BQ = sybxbl/100
fds=ROUND(300/100*SN,0)
bds=round(60/100*SN,0)
YB = 医疗保险％/100
ybsx=医保上限值
ybxx=医保下限值
GJ = 住房公积金/100
gjjsx=公积金上限
gjjxx=公积金下限
on key label F1 do grcx
set safety off
ZDM = space(3)
close all
****初始化变量
ND1 = val(LY)
YF1 = month(date())
RQ1 = day(date())
YF = iif(YF1>9,str(YF1,2),'0'+str(YF1,1))
RQ111 = iif(RQ1>9,str(RQ1,2),'0'+str(RQ1,1))
close all
dhsr=''
年 = year(date())-val(LY)
IF not file('Zr1'+LY+'.DBF')
************1.反算工资总额**************
close all
use sbdmk
gxyf=更新月份
use 基础数据
go top
loca for 年份 = '&LY'
DWBL = 3
GRBL = 8
  use gzze1
  COPY TO zr1&LY stru
  USE zr1&LY
  if file('ry'+LY+'01.DBF')
     append from ry&LY.01 FOR '职工'$ygxz1
  else
     append from ryzk FOR '职工'$ygxz1
  endif
  replace MOU with 12 all
  close all
  select 2
  if 年>1
    六月 = str(year(date()),4)+'0'+str(gxyf,1)
    if file('gz'+六月+'.dbf')
      use gz&六月
    else
      use ryzk
    endif
  else
    use ryzk
  endif
  index on RYBH to ry
  select 1
  use zr1&LY
  inde on rybh to zr1&LY
  if val(LY)<2000
    update on RYBH from b replace zr3 with b.ylbx , PJ with round(zr3*100/GRBL,1),HJ with;
 round(PJ*MOU,1),jfjs with pj
  else
    update on RYBH from b replace zr3 with b.ylbx , PJ with round(zr3*100/GRBL,0),HJ with round(PJ*MOU,0),jfjs with pj
  endif
  close all
dh=''
dycj=''
dybm=''
do dhkdy
use zr1&LY excl
do case
case dh=2
set filter to cjdm='&dycj'
case dh=3
set filter to bmbh='&dybm'
endcase
inde on bmbh+rybh to zr1&LY
go top
************人工校正**************
loca for zr3=0
if !eof()
browse field CJMC :H = '  单位名称  ' : 12 :R , RYXM :H = '人员姓名';
 : 8 :R , MOU :H = '缴费月数' : 8 , PJ :r: 8 :H = '平均工资' , 校正 , zr8 :r: 8 :H = '单位缴费' , zr3 :r: 8 :H = '个人缴费' for zr3=0 titl '新增人员[按 Esc 退出]'
endif
pack
go top
browse field CJMC :H = '  单位名称  ' : 12 :R , RYXM :H = '人员姓名';
 : 8 :R , MOU :H = '缴费月数' : 8 , PJ :r: 8 :H = '平均工资' , 校正 , zr8 :r: 8 :H = '单位缴费' , zr3 :r: 8 :H = '个人缴费' title '请认真核对'+ly+'年平均工资和缴费月数[按 F1 查找个人　按 Esc 退出]'
pack
go top
REPL pj with 校正 ,jfjs with pj,ZR3 with round(PJ*GRBL/100,1) , ZR8 with round(PJ*DWBL/100,1) FOR 校正>0
repl zr5 with 0 , hj with pj*mou all
if val(LY)<=1997
  replace ZR5 with round(SN*5/100,1) , HEJI with ZR8+ZR5+ZR3 all
else
  replace HEJI with ZR8+ZR3 all
endif
set filter to
wait window '正在自动生成'+LY+'年工资总额、'+str(val(ly)+1,4)+'年缴费基数......' nowait
close all
**************人工校正后生成工资总额库zr&LY********************
IF i=88
use gzze
copy to zr&LY stru
close all
select 2
use zr1&LY
index on rybh to zr1&LY
select 1
use zr&LY
append from zr1&LY
index on rybh to zr&LY
update on RYBH from b replace MOU with b.MOU
go top
do while !eof()
for IJB = 1 to MOU
  CMJB = iif(IJB>9,str(IJB,2),'0'+str(IJB,1))
 upda on rybh from B replace M&cmjb with b.PJ 
****************注意与“关联法”相区别**********
endfor
skip
enddo
replace A with 1 all
ENDIF
ELSE
close all
********2.正常情况下使用已处理的工资总额按新下达政策参数计算下年缴费基数jfjs及社会保险(五险两金)个人缴费金额***********
use zr1&LY
REPLACE jfjs with pj all
REPL jfjs with 校正  FOR 校正>0
*************工资外收入*****************************
repl zr8 with jfjs all
repl zr8 with ybsx for zr8>ybsx
repl zr8 with ybxx for zr8<ybxx
REPLACE ybx WITH ROUND(zr8*YB,2) all
***********昆明市医疗保险中心执行上下限封顶保底政策统一计算下一年缴费基数**************************
repl zr8 with jfjs all
repl zr8 with gjjsx for zr8>gjjsx
repl zr8 with gjjxx for zr8<gjjxx
REPLACE gjj WITH ROUND(zr8*GJ,0) all
***********昆明市住房公积金管理中心执行上下限封顶保底政策统一计算下一年缴费基数**************************
repl jfjs with fds for jfjs>fds
repl jfjs with bds for jfjs<bds
repl zr3 with round(jfjs*GR,2),sybx with round(jfjs*SY,2) all
repl sybx with round(jfjs*BQ,2) FOR 'B区'$zzjg
*****************A区、B区分开缴纳*************20150825********************************
repl heji WITH zr3+qynj+sybx+ybx+gjj all
go top
browse field CJMC :H = '  单位名称  ' : 12 :R , RYXM :H = '人员姓名';
 : 8 :R , MOU :H = '缴费月数' , PJ :r:H = ly+'年月平均工资' , 工资外收入 , jfjs :r:H =str(val(ly)+1,4)+'年缴费基数',校正, zr3 :r:H = '养老保险',sybx :r:H = '失业保险',ybx :r:H = '医疗保险',gjj :r:H = '住房公积金' title '请认真核对'+LY+'年月平均工资和缴费月数[按 F1 查找个人　按 Esc 退出]'
ENDIF
close all
select 2
use ryzk
index on rybh to ryzk
select 1
use zr1&LY
index on rybh to zr1&LY
update on RYBH from b replace cjdm with b.cjdm,cjmc with b.cjmc,bmbh with b.bmbh,bmmc with b.bmmc
sort on bmbh,zw,rybh to temp
use temp
copy to zr1&LY
CLOSE ALL
use zr1&LY
DBFF1='zr1'
WHNY=LY+'年工资总额'
do Qdir
close all
delete file temp.dbf
retu

