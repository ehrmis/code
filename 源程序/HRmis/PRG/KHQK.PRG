***************
* khqk.prg
***************
close all
stor 0 to 全,工,专,技,初,中,高
use zgjy
inde on dwdm to dw
tota on dwdm to dw
use dw
dm=allt(dwdm)
dw=allt(dwmc)
use khqk excl
delete for empty(dwdm)
pack
loca for dwdm='&dm'
if !found()
   appe blank
   repl dwdm with '&dm',dwmc with '&dw'
endif
close all
sele 3
use khzb
sele 2
use zjzk 
loca for 项目='操作人员'
sele 1
use khqk
go top
do while !eof()
repl 初级工培训 with int((b.初级在岗数-b.初级取证数)*c.初中级工培训指标/100),中级工培训 with int((b.中级在岗数-b.中级取证数)*c.初中级工培训指标/100),高级工培训 with int((b.高级在岗数-b.高级取证数)*c.高级工培训指标/100) 
skip
enddo 
close all   
use dmk
单位=allt(dw)
use zgjy
count for '工商管理'$allt(pxly) and year(pxrq)=year(date()) to 工商
close all
sele 3
use khzb
sele 2
use zjzk excl
loca for 项目='职 工 数'
zrs=职工人数
glrs=管理人数
jsry=技术人员
jsjy=继续教育
sum 职工人数 , 技术人员 to 完成数 , 技术 for 项目='学历培训' or  项目='资格培训' or  项目='适应性培训'
sele 1
use khqk 
loca for dwdm='&单位'
repl 工商管理培训 with 工商,技术人员 with jsry,职工人数 with zrs,全员培训 with 完成数,专业管理培训 with glrs-工商-技术
if zrs>0
   repl 全员培训率 with round(完成数/zrs*100,0)
endif
if 工商管理培训>0
  repl 工商管理培训率 with 100
endif
if 专业管理培训>0
  repl 专业管理培训率 with 100
endif   
if jsjy<c.专业人员继续教育 
   if jsry>0
      repl 继续教育 with 技术,技术人员 with jsry,继续教育完成率 with round(继续教育/技术人员*100,0)
   endif
else
   repl 继续教育完成率 with 100,继续教育 with ' ---- '
endif
sele 2
sum 初级在岗数 to 完成数 for 项目='学历培训' or  项目='资格培训' or  项目='适应性培训'
sele 1
repl 初级在岗数 with 完成数
if 初级工培训>0
  repl 初级培训率 with round(初级在岗数/初级工培训*100,1)
endif
sele 2
sum 中级在岗数 to 完成数 for 项目='学历培训' or  项目='资格培训' or  项目='适应性培训'
sele 1
repl 中级在岗数 with 完成数
if 中级工培训>0
  repl 中级培训率 with round(中级在岗数/中级工培训*100,1)
endif
sele 2
sum 高级在岗数 to 完成数 for 项目='学历培训' or  项目='资格培训' or  项目='适应性培训'
sele 1
repl 高级在岗数 with 完成数
if 高级工培训>0
  repl 高级培训率 with round(高级在岗数/高级工培训*100,1)
  if 高级培训率>c.高级工培训指标
     初=高级在岗数-高级工培训
  endif 
endif
if 全员培训率>=c.全员培训指标
   全=c.全员培训考评分
endif  
if 工商管理培训率>=100
   工=c.工商管理培训考评分
endif   
if 专业管理培训率>=100
   专=c.专业管理培训考评分
endif   
if 继续教育完成率>=100
   技=c.继续教育考评分
endif
if (初级在岗数+中级在岗数+初)/(初级工培训+中级工培训)*100>=c.初中级工培训指标
   中=c.初中级工培训考评分
endif  
 if 高级培训率>=c.高级工培训指标
   高=c.高级工培训考评分
endif 
repl 报表 with c.报表考评分 ,考核实得分 with 全+工+专+技+中+高+报表 
brow noedit title allt(dwmc)+'员工培训考核情况'
close all
retu
