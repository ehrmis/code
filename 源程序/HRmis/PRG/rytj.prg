close all
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
use data\dmk
bf=LEFT(ALLTRIM(盘符),3)
wait window '处  理  数  据　中, 请  稍  候... ...' nowait
use 人员变动 
copy to temp stru
use temp &&当前人员
appe from ryzk
close all
sele 2
use rybdtemp &&上月或上次检查人员,每运行一次更新一次
repl a with 1 all
inde on rybh to rybdtemp
sele 1
use temp
repl a with 0 all
inde on rybh to temp
upda on rybh from B repl a with b.a 
use 人员变动 excl
zap
appe from temp for a=0
repl 变动情况 with '增加',a with 1 all
close all
sele 2
use temp
repl a with 1 all
inde on rybh to temp
sele 1
use rybdtemp
repl a with 0 all
inde on rybh to rybdtemp
upda on rybh from B repl a with b.a 
use 人员变动 
appe from rybdtemp for a=0
repl 变动情况 with '减少' for a=0
repl a with 1 all                          
use rybdtemp
repl a with 1 all
go top
scan
sele 2
use temp
 WAIT WINDOW NOWAIT '正在检索人员变动情况'+A.cjdm+'：'+allt(A.cjmc)+'、'+allt(A.bmmc)+'→'+allt(A.ryxm)+'... ...' 
 repl a with 0 ,变动情况 with '单位变动'for dwdm<>A.dwdm and rybh= A->rybh
 repl a with 0 ,变动情况 with '车间变动'for cjdm<>A.cjdm and rybh= A->rybh
 repl a with 0 ,变动情况 with '部门变动'for bmbh<>A.bmbh and rybh= A->rybh
 repl a with 0 ,变动情况 with '职务变动'for zw<>A.zw and rybh= A->rybh
 repl a with 0 ,变动情况 with '在岗变动'for zgqk<>A.zgqk and rybh= A->rybh  
 repl a with 0 ,变动情况 with '岗位变动'for gz<>A.gz and rybh= A->rybh  
endscan
use 人员变动 
appe from temp for a=0
appe from temp for year(离岗计划)=year(date()) and month(离岗计划)=month(date())
repl a with 0 all
count for a=0 to jls
if jls>0 
inde on bmbh+rybh to 人员变动
for zdz=1 to fcount()
  zd=fiel(zdz)
  zdz1=(fsize(zd)+fsize(zd))*2
  if uppe(allt(zd))='RYXM' 
     exit
  endif
endfor   
go top
browse part zdz1 field dwmc:h='单位',cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',zw1:h='职务',zgqk1:h='在岗情况',变动情况,离岗计划,离岗时间 titl '[按 F1 查个人] [按 Esc 继续] 请认真审核和维护'
endif
close all
********************
use ryzk
repl a with 1 all
copy to temp1 for zgqk='01'
copy to temp2 for !empty(bmmc) and zgqk='01'
copy to temp3 for left(zzdm,2)='02' and zgqk='01'
if jls>0
use 人员变动
copy to &bf.人员变动情况 type xl5
 =messagebox("本月各单位人数统计导出到“"+&bf+"人员变动情况”电子表中！！",48,"恭喜")
endif
use temp1
inde on cjdm to temp1
tota on cjdm to temp
use temp
copy to &bf.各单位人数统计 fiel cjmc,a type xl5
 =messagebox("本月各单位人数统计导出到“"+&bf+"各单位人数统计”电子表中！！",48,"恭喜")
use temp1
inde on dwdm+cjdm to temp1
tota on dwdm+cjdm to temp
use temp
copy to &bf.各部门人数统计 fiel dwmc,cjmc,a type xl5
 =messagebox("本月各单位人数统计导出到“"+&bf+"各部门人数统计”电子表中！！",48,"恭喜")
use temp2
inde on dwdm+cjdm+bmbh to temp2
tota on dwdm+cjdm+bmbh to temp
use temp
copy to &bf.各班组人数统计 fiel dwmc,cjmc,bmmc,a type xl5
 =messagebox("本月各班组人数统计导出到“"+&bf+"各班组人数统计”电子表中！！",48,"恭喜")
use temp3
inde on xb1 to temp3
tota on xb1 to temp
use temp
copy to &bf.生产单位性别汇总 fiel xb1,a type xl5
 =messagebox("本月各班组人数统计导出到“"+&bf+"生产单位性别汇总”电子表中！！",48,"恭喜")
use temp3
inde on cjdm+bmbh+xb1 to temp3
tota on cjdm+bmbh+xb1 to temp
use temp
copy to &bf.生产单位性别统计 fiel cjmc,bmmc,xb1,a type xl5
 =messagebox("本月各班组人数统计导出到“"+&bf+"生产单位性别统计”电子表中！！",48,"恭喜")
use ryzk
copy to temp for val(zw)<11 and zgqk='01'
use temp
inde on dwdm+cjdm+bmbh+zw to temp
tota on dwdm+cjdm+bmbh+zw to temp1
use temp1
copy to &bf.职务统计 fiel dwmc,cjmc,bmmc,zw1,a type xl5
use ryzk
copy to temp fiel dwmc,cjmc,bmmc,zgqk,zgqk1,ryxm,zw1 for zgqk<>'01'
use temp
count for zgqk<>'01' to jls1
if jls1>0 
copy to &bf.在岗情况 type xl5
 =messagebox("本月职务统计导出到“"+&bf+"在岗情况”电子表中！！",48,"恭喜")
endif
close all
yn = MESSAGEBOX("数据导出成功！现打开使用吗？",4+32,"提示")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"人员分布统计表")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
if jls>0
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"人员变动情况")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
endif
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"各单位人数统计")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"各部门人数统计")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"各班组人数统计")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"生产单位性别汇总")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"生产单位性别统计")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"职务统计")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
if jls1>0 
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"在岗情况")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
endif
ENDI  
use 人员变动 excl
copy to rybdtemp stru
use rybdtemp &&更新为当前人员
IF FILE('ry'+ny1+'.dbf')
   appe from ry&ny1
else   
   appe from ryzk 
ENDIF
delete file temp.dbf
delete file temp1.dbf