***************************
* .\kqSR.PRG(Build20140123)
***************************
on escape
set escape off
on key label F1 do grsr
on key label F2 do sykqts
close all
set safety off
set talk off
clear 
_SCREEN.PICTURE = ''
dh=''
dycj=''
dybm=''
ND = '  '
YF = '  '
XH1 = 'N'
LY = ''
T123 = '  '
close all
set confirm off
XH1 = 'N'
ND1 = year(date())
YF1 = month(date())
do srny2.spx
LYNY = ND+YF
****************前月自动处理****************
if val(yf)<3
  JJND = str(val(nd)-1,4)
  JJYF = str(10+val(yf),2)
else
  JJND = ND
  JJYF = iif(val(YF)-2>9,str(val(YF)-2,2),'0'+str(val(YF)-2,1))
endif
JJLY = JJND+JJYF
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
endif
****************初始化****************
T123 = '  '
do dhkdy
do case
  case dh=2
  T123 = dyCJ
  case dh=3
  T123 = dybm
endcase
close all
clear
XH1 = 'N'
if !file('LW'+LYNY+'.DBF')
  use lwpqgzk excl
  zap
  COPY TO lw&LYNY
  USE lw&LYNY
  append from ryzk for ygxz='02'
endif
close all
*********************初始化****************
select 2
use ryzk
index on RYBH to ryzk
select 1
use lw&LYNY
inde on rybh to lw&LYNY
update on RYBH from B replace CJDM with b.CJDM , BMBH;
 with b.BMBH , CJMC with b.CJMC , BMMC with b.BMMC ,erprybh with b.erprybh,gz1 with b.gz1,gw with b.gw
*********************人工输入考勤****************
close all
select 4
use lw&syJJ
index on RYBH to temp
select 3
use ryzk
index on RYBH to ryzk
select 1
use SMK1
select 2
use lw&LYNY
set relation to RYBH into 3
set relation to RYBH into 4 additive
replace BJ with 21 for c.RYFL='33'
index on BMBH+RYBH to lw&LYNY
if !empty(T123)
  go bottom
do case
   case dh=2
   set filter to cjdm=T123
   case dh=3
   set filter to bmbh=T123
endcase
  go top
  NUM1 = recno()
  go bottom
  NUM2 = recno()
  go top
else
  go top
  NUM1 = recno()
  go bottom
  NUM2 = recno()
  go top
endif
XH1 = 'Y'
do while upper(XH1)='Y'
  XH2 = 'N'
  select 1
  go top
  KKK1 = 5
  LLL1 = 3
  do while  not eof()
    @ KKK1 , 20+LLL1 say CHINES font '宋体' , 14
    KKK1 = KKK1+2
    if KKK1>20
      KKK1 = 5
      LLL1 = LLL1+35
    endif
    skip
  enddo
  @ 23 , 33 say 'PgDn=下一条记录，PgUp=上一条记录' font '宋体' , 15
  @ 25 , 33 say 'F1=查找指定人员，F2=查询上月考勤，Esc=存盘退出' font '宋体' , 15
  do while upper(XH2)='N'
    select 1
    go top
    KKK1 = 5
    LLL1 = 3
    do while  not eof()
      RYHDS = FIELD_NAME
      TM123 = HHH1
      select 2
      if upper(trim(RYHDS))='RYBH' or upper(trim(RYHDS))='RYXM' or upper(trim(RYHDS))='BMMC';
 or upper(trim(RYHDS))='CJMC'
        RYHDS=&RYHDS
        @ KKK1 , 35+LLL1 say RYHDS font '宋体' , 14
      else
        @kkk1,35+lll1 get &ryhds font '宋体' , 14 PICT &TM123
      endif
      KKK1 = KKK1+2
      if KKK1>20
        KKK1 = 5
        LLL1 = LLL1+35
      endif
      select 1
      skip
    enddo
    read
    M = readkey()
    select 2
    do case
    case M=6 or M=262
      if recno()<>NUM1
        skip -1
      endif
      loop
    case M=7 or M=263
      if recno()=NUM2
        go bottom
      else
        skip
      endif
      loop
    case M=2 or M=258
      go top
      loop
    case M=12 or M=268
      exit
    endcase
  enddo
  FF = readkey()
  if FF=12 or FF=268
    exit
  endif
enddo
********************保存数据和钢性处理数据***********************
on key label F1 do grcx
on key label F2
select 2
*replace JRJB with 0,zbgr with 0,ybgr with 0,xcts with 0 for '处'$c.ZW1
*************************特殊处理*************************************
index on BMBH+RYBH to lw&LYNY
   rs=recc()
   rs=allt(str(rs))
   sum fzsd,kk to sfzsd , skk
   count for bj>0 to bjj
   bjj=allt(str(bjj))
   count for sj>0 to sjj
   sjj=allt(str(sjj))
if dh<>1
   if dh=2
   count for left(BMBH,2)=T123  to rs1
   rs1=allt(str(rs1))
   sum fzsd,kk to sfzsd1,skk1 for left(BMBH,2)=T123 
   count for left(BMBH,2)=T123 and bj>0 to bjj1
   bjj1=allt(str(bjj1))
   count for left(BMBH,2)=T123 and sj>0 to sjj1
   sjj1=allt(str(sjj1))
   go top
   loca for left(BMBH,2)=T123
   cjm=allt(cjmc)
   brow part 20 when sykqts() fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',zbgr:h='中班',ybgr:h='夜班',bj1:h='保健',jrjb:h='加班',bj:h='病假',sj:h='事假',cj:h='产假',gs:h='工伤',tqj:h='探亲',hsj:h='婚丧',jl:h='拘留',kg:h='旷工',fzsd:h='房租水电',kk:h='扣款';
    titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs1+'人，病假='+bjj+'人；事假='+sjj+'人；其中，'+cjm+'：'+rs1+'人，病假='+bjj1+'人；事假='+sjj1+'人；房租水电＝'+allt(str(sfzsd1,10,1))+'元；扣款＝'+allt(str(skk1,10,1))+'元   按[ Esc ]键退出'
  else
   count for BMBH=T123  to rs1
   rs1=allt(str(rs1))
   sum fzsd,kk to sfzsd2,skk2 for BMBH=T123 
   count for BMBH=T123 and bj>0 to bjj1
   bjj1=allt(str(bjj1))
   count for BMBH=T123 and sj>0 to sjj1
   sjj1=allt(str(sjj1))
   go top
   loca for BMBH=T123
   bm=allt(bmmc)
  brow part 20 when sykqts() fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',zbgr:h='中班',ybgr:h='夜班',bj1:h='保健',jrjb:h='加班',bj:h='病假',sj:h='事假',cj:h='产假',gs:h='工伤',tqj:h='探亲',hsj:h='婚丧',jl:h='拘留',kg:h='旷工',fzsd:h='房租水电',kk:h='扣款';
   titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs1+'人，病假='+bjj+'人；事假='+sjj+'人；其中，'+bm+'：'+rs1+'人，病假='+bjj1+'人；事假='+sjj1+'人；房租水电＝'+allt(str(sfzsd2,10,1))+'元；扣款＝'+allt(str(skk2,10,1))+'元  按[ Esc ]键退出'
   endif
else
   go top
  brow part 20 when sykqts() fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',zbgr:h='中班',ybgr:h='夜班',bj1:h='保健',jrjb:h='加班',bj:h='病假',sj:h='事假',cj:h='产假',gs:h='工伤',tqj:h='探亲',hsj:h='婚丧',jl:h='拘留',kg:h='旷工',fzsd:h='房租水电',kk:h='扣款';
   titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs+'人，病假='+bjj+'人；事假='+sjj+'人；房租水电＝'+allt(str(sfzsd,10,1))+'元；扣款＝'+allt(str(skk,10,1))+'元  按[ Esc ]键退出'
endif
go top
loca for jrjb+bj+sj+cj+gs+tqj+hsj+jl+kg+gj>0
if !eof()
 brow part 20 when sykqts() fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',jrjb:h='加班',bj:h='病假',sj:h='事假',cj:h='产假',gs:h='工伤假',tqj:h='探亲假',gj:h='年休假',hsj:h='婚丧假',jl:h='拘留',kg:h='旷工';
   titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs+'人，病假='+bjj+'人；事假='+sjj+'人  按[ Esc ]键退出' for jrjb+bj+sj+cj+gs+tqj+hsj+jl+kg+gj>0
endif
close all
********************DHS*****************
*if val(dm111)=99
*sele 2
*use ryzk
*inde on rybh to ryzk
*select 1
*use lw&LYNY
*inde on rybh to lw&lyny
*set relation to rybh into 2
 *  repl sfgz with 800*b.sfbl/100 for xcts>=15
  * repl sfgz with xcts*round(800/15,2)*b.sfbl/100 for xcts<15
   *repl sfgz with 0 for tqj+gs>=21
   *repl jang2 with 240,sfgz with 250*c.sfbl/100 for bmbh='0114' and xcts>=15
   *repl jang2 with 240,sfgz with xcts*round(250/15,2)*c.sfbl/100 for bmbh='0114' and xcts<15
*endif
close all
delete file temp.idx
delete file temp.dbf
*************************还原图片*****************
clear
 _SCREEN.WINDOWSTATE = 2
close all
use ccrq
go 8
repl 图片 with str(val(图片)+1,2)
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
retu


******************
PROCEDURE SYKQTS
******************
wait wind alltrim(b.RYXM)+'[上月]：'+'中班:'+alltrim(str(d.zbgr))+'夜班:'+alltrim(str(d.ybgr))+'保健:'+alltrim(str(d.Bj1))+'加班:'+alltrim(str(d.JrJB))+';病假:'+alltrim(str(d.BJ))+';'+'事假:'+alltrim(str(d.SJ))+';'+'产假:'+alltrim(str(d.CJ))+';工休:'+alltrim(str(d.GS))+';探亲:'+alltrim(str(d.TQJ))+';婚丧:'+alltrim(str(d.HSJ))+';拘留:'+alltrim(str(d.JL))+';旷工:'+alltrim(str(d.KG)) nowa at 34,10
