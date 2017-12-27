close all
USE  ryzk
COPY TO temp FIELDS gz,gz1,gw,gwjb,gwjb1,bzgz,gwgz,gwfl1,zyflm,zymc,a
USE temp
INDEX on gz1+gw TO temp
TOTAL ON gz1+gw TO temp1
USE temp1
REPLACE gwgz WITH bzgz/a all
GO top
brow fiel gz:h='工种代码',gz1:h='工种',gw:h='岗位',gwjb1:h='岗位级别',gwgz:h='平均岗位工资',a:h='从事本岗位人数' titl '本单位当前工种岗位分类汇总查询'
copy to &bf.\工种代码 type xl5
yn = MESSAGEBOX("本单位当前工种岗位汇总导出成功！现打开使用吗？",4+32,"提示")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\工种代码.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
ENDIF
close all
***************************
yn = MESSAGEBOX("需要用本单位员工当前岗位级别标准工资更新“工种代码库”吗？",4+32,"提示")
IF yn = 6
CLOSE all
SELECT 3
USE gwdmk
INDEX on gwdm TO gwdmk
sele 2
use temp1
scan
sele 1
use gongzong
 loca for allt(name)=allt(b.gz1) and allt(gwname)=allt(b.gw) 
 WAIT WINDOW NOWAIT '正在维护'+gwdm+'：'+allt(name)+'→'+allt(gwname)+'<工种代码库>... ...' 
 repl gwdm with b.gwjb
 SET RELATION TO gwdm INTO 3
 repl gwdj with c.gd01,bzgz WITH c.bzgz 
 ****单位岗位等级代码维护**201719****
ENDSCAN
ENDIF
CLOSE all
use gongzong EXCLUSIVE
go top
brow fiel code:h='工种代码',name:h='工种',gwname:h='岗位',gwdm:h='岗等代码',gwdj:h='岗位等级',bzgz:h='标准工资' titl '自定义规范命名本单位当前工种岗位名称和代码（保存后将按工种数据库工种代码自动更新岗位从业人员工种代码）'
************管理员专业维护代码（按单位工种岗位实际情况统一维护）****20160829******
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
sort on code to data\gongzong
CLOSE all
yn = MESSAGEBOX("用工种岗位数据库置换本单位员工当前工种代码吗？",4+32,"提示")
IF yn = 6
CLOSE all
sele 1
use ryzk
scan
 sele 2
 use gongzong
 loca for allt(name)=allt(A.gz1) and allt(gwname)=allt(A.gw) 
 WAIT WINDOW NOWAIT '正在维护'+A.cjdm+'：'+allt(A.cjmc)+'、'+allt(A.bmmc)+'→'+allt(A.ryxm)+'<工种代码>... ...' 
 repl A.gz with code
 ****单位工种代码维护**20160830****
ENDSCAN
SELECT  1
go top
brow fiel dwmc:h='单位',cjmc:h='车间',bmmc:h='班组',ryxm:h='姓名',gz:h='工种代码',gz1:h='工种',gw:h='岗位',bzgz:h='标准工资',gwgz:h='岗位工资'titl '置换后本单位员工当前工种岗位'
close all
ENDIF

