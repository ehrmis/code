***************************
* .\SRKQ.PRG(Build20161206)
***************************
on key label F1 do kqgrcx
on key label F2 do sykqts
on key label F3 do plsr
clear
_SCREEN.PICTURE = ''
dh=''
XH1 = 'N'
LY = ''
T123 = '  '
close all
USE dmk
现场数=计发天数
use deset
GO top
LOCATE FOR seted=0 AND 'sfgz'$oop
IF !EOF()
   sfgz1='清零'   
ELSE
sfgz1='正常'
USE dmk
标准=边远补贴
天数=计发天数
日标=round(标准/天数,2)
条件1=allt(计发条件)
条件2=allt(扣发条件)
条件3=allt(清零条件)
ENDIF
set confirm off
LY = ND
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
endif
****************初始化*******20161020*********
if i=99
do case
   case VAL(dwdm1)=0 and VAL(cjdm1)>0 and val(bmbh1)=0 
   dh=2
   ******车间输入******
   T123 = cjdm1
   case VAL(dwdm1)>=0 and VAL(cjdm1)>0 and val(bmbh1)>0
   dh=3 
   *************班组输入**********   
   T123 = bmbh1
   **********精确筛选班组编号****************
   case VAL(dwdm1)>0 and VAL(cjdm1)>0 and val(bmbh1)=0 
   dh=2
   T123 = cjdm1
   *************模糊单位、仅选车间输入**********
   case VAL(dwdm1)>0 and VAL(cjdm1)=0 and val(bmbh1)>0 
   dh=3 
   T123 = bmbh1
   *************模糊单位、仅选班组输入********** 
   OTHERWISE 
   T123 = '  '
   dh=1
   *************模糊输入**********
endcase
ELSE
T123 = '  '
dh=1
ENDIF
close all
clear
XH1 = 'N'
if !file('KQ'+ny+'.DBF')
  use KQK excl
  ZAP
 IF FILE('ry'+ny+'.dbf')
  append from ry&ny for ygxz='01' and !'退'$zgqk1
  ********************非退休退养在册职工****20151217****************
 ELSE
  append from ryzk for ygxz='01' and !'退'$zgqk1
 ENDIF 
 SORT ON bmbh,zw,rybh  TO KQ&ny
 *********直接排序生成当月考勤*****20161026********************
endif
close all
*********************初始化****************
************修正初始化直接全部追加数据****20161205******************************
select 2
 IF FILE('ry'+ny+'.dbf')
    use ry&ny 
    repl a with 1 all
    index on RYBH to ry&ny
 ELSE
    use ryzk 
    repl a with 1 all
    index on RYBH to ryzk
 ENDIF
select 1
use kq&ny EXCLUSIVE
repl 年休计划 with CTOD(''),a with 0 all
************修正初始化直接全部追加数据*****检查本月应删除/增加人员及时更新***20160613*********
inde on rybh to kq&ny
update on RYBH from B replace CJDM with b.CJDM , BMBH;
 with b.BMBH , CJMC with b.CJMC , BMMC with b.BMMC ,erprybh with b.erprybh,bjgz with b.bjgz,bjdj with b.bjdj,bjjb with b.bjjb,;
 zbjt with b.zbjt,ybjt with b.ybjt,gz1 with b.gz1,zw with b.zw,作业班制 with b.作业班制,zgqk with b.zgqk,a with b.a
GO top
loca for a=0 
IF !eof()
   delete for a=0 
   brow for a=0 titl '请认真审核下列人员是否删除'
    yn = MESSAGEBOX("需要删除上述标记人员吗？",4+32,"提示")
   IF yn = 6
      PACK
****功能：先提示删除再关联当月人员状况库确定是否有新增加人员？******20160613*************
   ELSE
       RECALL all
   ENDIF         
ENDIF
CLOSE ALL
select 2
use kq&ny
repl a with 0 all
index on RYBH to kq&ny
select 1
IF FILE('ry'+ny+'.dbf')
    use ry&ny 
    repl a with 1 all
    index on RYBH to ry&ny
 ELSE
    use ryzk 
    repl a with 1 all
    index on RYBH to ryzk
 ENDIF 
update on RYBH from B replace a with b.a
GO top
loca for a=1 and ygxz='01' 
if !eof()
   brow for a=1 and ygxz='01' titl '请认真审核下列新增加入考勤库中人员'
   CLOSE all
   use kq&ny
   IF FILE('ry'+ny+'.dbf')
    APPEND FROM ry&ny for a=1 and ygxz='01'
   ELSE
    APPEND FROM ryzk for a=1 and ygxz='01'
   ENDIF  
ENDIF
CLOSE all
use kq&ny 
repl a with 1 all
 IF FILE('ry'+ny+'.dbf')
    use ry&ny 
    repl a with 1 all
 ELSE
    use ryzk 
    repl a with 1 all
 ENDIF 
**************必须还原默认标识*****20160613************************      
close all
IF files('kq'+syjj+'.dbf')
select 4
use kQ&syJJ
index on RYBH to temp
ENDIF 
select 3
IF FILE('ry'+ny+'.dbf')
    use ry&ny 
    index on RYBH to ry&ny
 ELSE
    use ryzk 
    index on RYBH to ryzk
 ENDIF 
select 1
use SMK1
select 2
use kQ&ny
set relation to RYBH into 3
IF files('kq'+syjj+'.dbf')
set relation to RYBH into 4 additive
ENDIF 
index on BMBH+zw+RYBH to kq&ny
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
DEFINE WINDOW temp FROM INT((SROWS()-40)/2)  , INT((SCOLS()-160)/2)  TO INT((SROWS()+40)/2) , INT((SCOLS()+160)/2) title 'DOS时代经典小键盘快速考勤输入模式 [ 使用小键盘快速输入，用回车键确认保存数据 ]'
activate window temp
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
  @ 23 , 33 say 'PgDn = 下一个　PgUp = 上一个  Enter = 保存' font '宋体' , 15
  @ 25 , 33 say 'F1 = 查找个人　F2 = 上月考勤　F3 = 批量输入　Esc = 退出' font '宋体' , 15
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
        @kkk1,35+lll1 get &ryhds font '宋体' , 14 
        *PICT &TM123
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
ENDDO
replace 年休计划 with c.年休计划 FOR YEAR(c.年休计划)=YEAR(date()) AND YEAR(终止时间)=YEAR(date()) AND rybh=c.rybh
*********************自动填写当月休假人员的当年年休计划****************
CLEAR WINDOWS temp
********************1.保存数据和钢性处理数据***********************
on key label F2
on key label F3
************释放自定义热键****20161128**************
select 2
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
   IF files('kq'+syjj+'.dbf')
   brow part 20 when sykqts() fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',xcts:h='现场天数',zbgr:h='中班',ybgr:h='夜班',gj:h='年休假',年休计划,起始时间,终止时间,bj1:h='保健',jrjb:h='节日加班',jjb:h='假日加班',xxpx:h='出差培训',bj:h='病假',sj:h='事假',cj:h='陪护/产假',gs:h='工伤',tqj:h='探亲',hunj:h='婚假',sangj:h='丧假',gx:h='工休',ly:h='疗养',jl:h='拘留',kg:h='旷工',fzsd:h='房租水电',kk:h='扣款';
    titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs1+'人，病假='+bjj+'人；事假='+sjj+'人；其中，'+cjm+'：'+rs1+'人，病假='+bjj1+'人；事假='+sjj1+'人；房租水电＝'+allt(str(sfzsd1,10,1))+'元；扣款＝'+allt(str(skk1,10,1))+'元   按[ Esc ]键退出'
   ELSE
   brow part 20 fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',xcts:h='现场天数',zbgr:h='中班',ybgr:h='夜班',gj:h='年休假',gj:h='年休假',年休计划,起始时间,终止时间,bj1:h='保健',jrjb:h='节日加班',jjb:h='假日加班',xxpx:h='出差培训',bj:h='病假',sj:h='事假',cj:h='陪护/产假',gs:h='工伤',tqj:h='探亲',hunj:h='婚假',sangj:h='丧假',gx:h='工休',ly:h='疗养',jl:h='拘留',kg:h='旷工',fzsd:h='房租水电',kk:h='扣款';
    titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs1+'人，病假='+bjj+'人；事假='+sjj+'人；其中，'+cjm+'：'+rs1+'人，病假='+bjj1+'人；事假='+sjj1+'人；房租水电＝'+allt(str(sfzsd1,10,1))+'元；扣款＝'+allt(str(skk1,10,1))+'元   按[ Esc ]键退出'
   ENDIF
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
   IF files('kq'+syjj+'.dbf')
  brow part 20 when sykqts() fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',xcts:h='现场天数',zbgr:h='中班',ybgr:h='夜班',gj:h='年休假',年休计划,起始时间,终止时间,bj1:h='保健',jrjb:h='节日加班',jjb:h='假日加班',xxpx:h='出差培训',bj:h='病假',sj:h='事假',cj:h='陪护/产假',gs:h='工伤',tqj:h='探亲',hunj:h='婚假',sangj:h='丧假',gx:h='工休',ly:h='疗养',jl:h='拘留',kg:h='旷工',fzsd:h='房租水电',kk:h='扣款';
   titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs1+'人，病假='+bjj+'人；事假='+sjj+'人；其中，'+bm+'：'+rs1+'人，病假='+bjj1+'人；事假='+sjj1+'人；房租水电＝'+allt(str(sfzsd2,10,1))+'元；扣款＝'+allt(str(skk2,10,1))+'元  按[ Esc ]键退出'
   ELSE
   brow part 20 fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',xcts:h='现场天数',zbgr:h='中班',ybgr:h='夜班',gj:h='年休假',年休计划,起始时间,终止时间,bj1:h='保健',jrjb:h='节日加班',jjb:h='假日加班',xxpx:h='出差培训',bj:h='病假',sj:h='事假',cj:h='陪护/产假',gs:h='工伤',tqj:h='探亲',hunj:h='婚假',sangj:h='丧假',gx:h='工休',ly:h='疗养',jl:h='拘留',kg:h='旷工',fzsd:h='房租水电',kk:h='扣款';
   titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs1+'人，病假='+bjj+'人；事假='+sjj+'人；其中，'+bm+'：'+rs1+'人，病假='+bjj1+'人；事假='+sjj1+'人；房租水电＝'+allt(str(sfzsd2,10,1))+'元；扣款＝'+allt(str(skk2,10,1))+'元  按[ Esc ]键退出'
   ENDIF
   endif
else
   go top
  IF files('kq'+syjj+'.dbf')  
  brow part 20 when sykqts() fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',xcts:h='现场天数',zbgr:h='中班',ybgr:h='夜班',gj:h='年休假',年休计划,起始时间,终止时间,bj1:h='保健',jrjb:h='节日加班',jjb:h='假日加班',xxpx:h='出差培训',bj:h='病假',sj:h='事假',cj:h='陪护/产假',gs:h='工伤',tqj:h='探亲',hunj:h='婚假',sangj:h='丧假',gx:h='工休',ly:h='疗养',jl:h='拘留',kg:h='旷工',fzsd:h='房租水电',kk:h='扣款';
   titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs+'人，病假='+bjj+'人；事假='+sjj+'人；房租水电＝'+allt(str(sfzsd,10,1))+'元；扣款＝'+allt(str(skk,10,1))+'元  按[ Esc ]键退出'
  ELSE
  brow part 20 fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',xcts:h='现场天数',zbgr:h='中班',ybgr:h='夜班',gj:h='年休假',年休计划,起始时间,终止时间,bj1:h='保健',jrjb:h='节日加班',jjb:h='假日加班',xxpx:h='出差培训',bj:h='病假',sj:h='事假',cj:h='陪护/产假',gs:h='工伤',tqj:h='探亲',hunj:h='婚假',sangj:h='丧假',gx:h='工休',ly:h='疗养',jl:h='拘留',kg:h='旷工',fzsd:h='房租水电',kk:h='扣款';
   titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs+'人，病假='+bjj+'人；事假='+sjj+'人；房租水电＝'+allt(str(sfzsd,10,1))+'元；扣款＝'+allt(str(skk,10,1))+'元  按[ Esc ]键退出'
  ENDIF
ENDIF
index on BMBH+zw+RYBH to kq&ny
REPLACE xcts WITH 0,sbts WITH 0 FOR zgqk>'02'
*************不在岗人员考勤清零*****20170726*****************
count for sbts=0 OR xcts>25 or xcts<现场数 or xcts<>sbts or sj>7 TO rs1
**********考勤异常提示****20160902*************** 
if rs1>0
 IF files('kq'+syjj+'.dbf')
   brow part 20 when sykqts() field cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',xcts:h='现场天数',bj:h='病假',sj:h='事假',cj:h='陪护/产假',gs:h='工伤',tqj:h='探亲',gx:h='工休',jl:h='拘留' for sbts=0 OR xcts>25 or xcts<现场数 or xcts<>sbts titl '请确认下列人员现场天数'
 ELSE
   brow part 20 field cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',sbts:h='出勤天数',xcts:h='现场天数',bj:h='病假',sj:h='事假',cj:h='陪护/产假',gs:h='工伤',tqj:h='探亲',gx:h='工休',jl:h='拘留' for sbts=0 OR xcts>25 or xcts<现场数 or xcts<>sbts titl '请确认下列人员现场天数'
 ENDIF
   REPLACE a WITH 8 for sbts=0 OR xcts>25 or xcts<现场数 or xcts<>sbts or sj>7
   *****************争议性异常标识a=8为出勤差错*******20161027***********************************
endi
count for !'常早班'$作业班制 and (zbgr=0 OR ybgr=0) TO rs1
****************非常早班人员可能少输入中夜班*****************************************
if rs1>0
   IF files('kq'+syjj+'.dbf')
      brow part 20 when sykqts() fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',作业班制,sbts:h='出勤天数',xcts:h='现场天数',zbgr:h='中班',ybgr:h='夜班' for !'常早班'$作业班制 and (zbgr=0 OR ybgr=0) titl '请确认下列'+allt(STR(rs1))+'人非常早班考勤'
   ELSE
      brow part 20 fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',作业班制,sbts:h='出勤天数',xcts:h='现场天数',zbgr:h='中班',ybgr:h='夜班' for !'常早班'$作业班制 and (zbgr=0 OR ybgr=0) titl '请确认下列'+allt(STR(rs1))+'人非常早班考勤'
   ENDIF
    REPLACE a WITH 8 for !'常早班'$作业班制 and (zbgr=0 OR ybgr=0)
   *****************争议性异常标识a=8为出勤差错*******20161027***********************************
endif
count for '常早班'$作业班制 and zbgr+ybgr>=13 TO rs1
****************常早班人员可能误输入中夜班*****************************************
if rs1>0
   IF files('kq'+syjj+'.dbf')   
      brow part 20 when sykqts() fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',作业班制,sbts:h='出勤天数',xcts:h='现场天数',zbgr:h='中班',ybgr:h='夜班' for '常早班'$作业班制 and zbgr+ybgr>=13 titl '请确认下列'+allt(STR(rs1))+'人常早班考勤'
   ELSE
      brow part 20 fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',作业班制,sbts:h='出勤天数',xcts:h='现场天数',zbgr:h='中班',ybgr:h='夜班' for '常早班'$作业班制 and zbgr+ybgr>=13 titl '请确认下列'+allt(STR(rs1))+'人常早班考勤'
   ENDIF
    REPLACE a WITH 8 for '常早班'$作业班制 and zbgr+ybgr>=13
   *****************争议性异常标识a=8为出勤差错*******20161027***********************************
endif
***********************2.出勤天数/中夜班考勤异常审核更正********20161024*************
index on BMBH+zw+RYBH to kq&NY 
count for bj+sj+cj+gs+tqj+hunj+sangj+jl+kg+gj+gx+ly+xxpx>0 TO rs1 
zrs=RECCOUNT()
zrs=ALLTRIM(STR(zrs))
********20160627***********
if rs1>0
rs1=ALLTRIM(STR(rs1))
SCAN
DO CASE 
   CASE bj>0
        REPLACE 缺勤情况 WITH '病假'+ALLTRIM(STR(bj))+'天'
   CASE sj>0
        REPLACE 缺勤情况 WITH '事假'+ALLTRIM(STR(sj))+'天' 
   CASE cj>0
        IF xb1='女'
           REPLACE 缺勤情况 WITH '产假'+ALLTRIM(STR(cj))+'天'
        ELSE
           REPLACE 缺勤情况 WITH '陪护假'+ALLTRIM(STR(cj))+'天'
        ENDIF            
   CASE gs>0
        REPLACE 缺勤情况 WITH '工伤假'+ALLTRIM(STR(gs))+'天'
   CASE tqj>0
        REPLACE 缺勤情况 WITH '探亲假'+ALLTRIM(STR(tqj))+'天'
   CASE gj>0
        REPLACE 缺勤情况 WITH '年休假'+ALLTRIM(STR(gj))+'天' 
   CASE hunj>0
        REPLACE 缺勤情况 WITH '婚假'+ALLTRIM(STR(hunj))+'天'
   CASE sangj>0
        REPLACE 缺勤情况 WITH '丧假'+ALLTRIM(STR(sangj))+'天'     
   CASE jl>0
        REPLACE 缺勤情况 WITH '拘留'+ALLTRIM(STR(jl))+'天'     
   CASE kg>0
        REPLACE 缺勤情况 WITH '旷工'+ALLTRIM(STR(kg))+'天' 
   CASE gx>0
        REPLACE 缺勤情况 WITH '工休'+ALLTRIM(STR(gx))+'天'
   CASE ly>0
        REPLACE 缺勤情况 WITH '疗养'+ALLTRIM(STR(ly))+'天'
   CASE xxpx>0
        REPLACE 缺勤情况 WITH '出差培训'+ALLTRIM(STR(xxpx))+'天'                   
 ENDCASE
 ENDSCAN              
 brow part 20 fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',xcts:h='出勤',缺勤情况,bj:h='病假',sj:h='事假',cj:h='产假',gs:h='工伤假',tqj:h='探亲假',gj:h='年休假',hunj:h='婚假',sangj:h='丧假',jl:h='拘留',kg:h='旷工',gx:h='工休',ly:h='疗养',xxpx:h='培训';
   titl nd+'年'+allt(str(val(yf)))+'月份缺勤人员：共'+rs1+'/'+zrs+'人 按[ Esc ]键退出' for  bj+sj+cj+gs+tqj+hunj+sangj+jl+kg+gj+gx+ly+xxpx>0
ENDIF
repl ERP编号 with erprybh,出勤天数 with sbts,现场天数 with xcts,中班 with zbgr,夜班 with ybgr,下井 with xjgr,节日加班 with jrjb,假日加班 with jjb,出差培训 WITH xxpx,病假 with bj,事假 with sj,产假 with cj,工伤 with gs,探亲假 with tqj,婚丧假 with hsj,婚假 with hunj,丧假 with sangj,疗养 with ly,工休 WITH gx,拘留 with jl,旷工 with kg,年休假 with gj,保健费 with bj1,上浮工资 with sfgz,房租水电 with fzsd,扣款 with kk,互助金 with fzj all
copy to &bf.\缺勤异常情况 fiel dwmc,cjmc,bmmc,ryxm,xb1,作业班制,出勤天数,现场天数,中班,夜班,缺勤情况,出差培训,病假,事假,产假,工伤,探亲假,婚丧假,婚假,丧假,疗养,工休,拘留,旷工,年休假 type XL5 for bj+sj+cj+gs+tqj+hunj+sangj+jl+kg+gj+gx+ly+xxpx>0 OR a=8
*************************************************3.人工输入考勤后全部置换导出ERP系统考勤汉字字段备用****************20160812*********************提供缺勤情况或出勤异常列表复核功能***************
REPLACE hsj WITH hunj+sangj,a WITH 1 all
************婚假、丧假单列******20160405**********************
go top
loca for jrjb+jjb+bj+sj+cj+gs+tqj+hsj+jl+kg+gj>0
if !eof()
 IF files('kq'+syjj+'.dbf') 
 brow part 20 when sykqts() fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',xb1,作业班制,出勤天数,现场天数,中班,夜班,缺勤情况,jrjb:h='节日加班',jjb:h='假日加班',xxpx:h='出差培训',bj:h='病假',sj:h='事假',cj:h='陪护/产假',gs:h='工伤假',tqj:h='探亲假',gj:h='年休假',hunj:h='婚假',sangj:h='丧假',gx:h='工休',ly:h='疗养',jl:h='拘留',kg:h='旷工';
   titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs+'人，病假='+bjj+'人；事假='+sjj+'人  按[ Esc ]键退出' for jrjb+bj+sj+cj+gs+tqj+hsj+jl+kg+gj>0
 ELSE
  brow part 20 fiel cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',xb1,作业班制,出勤天数,现场天数,中班,夜班,缺勤情况,jrjb:h='节日加班',jjb:h='假日加班',xxpx:h='出差培训',bj:h='病假',sj:h='事假',cj:h='陪护/产假',gs:h='工伤假',tqj:h='探亲假',gj:h='年休假',hunj:h='婚假',sangj:h='丧假',gx:h='工休',ly:h='疗养',jl:h='拘留',kg:h='旷工';
   titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs+'人，病假='+bjj+'人；事假='+sjj+'人  按[ Esc ]键退出' for jrjb+bj+sj+cj+gs+tqj+hsj+jl+kg+gj>0
 ENDIF
ENDIF
***********************当月考勤异常总审核与更正********20161024*************
close all
select 2
 IF FILE('ry'+ny+'.dbf')
    use ry&ny 
    index on RYBH to ry&ny
 ELSE
    use ryzk 
    index on RYBH to ryzk
 ENDIF 
 **************工龄取整计算连续工龄（周年计算）****“严格区分满与不满周年”*****20161121*************
scan FOR '职工'$ygxz1
周年=INT(gl)
DO case
CASE 周年<1
     REPLACE sfbl WITH 0
CASE 周年=1
     REPLACE sfbl WITH 20
CASE 周年=2
     REPLACE sfbl WITH 40
CASE 周年=3
     REPLACE sfbl WITH 60
CASE 周年=4
     REPLACE sfbl WITH 80
CASE 周年>=5
     REPLACE sfbl WITH 100
ENDCASE
endscan              
select 1
use kq&ny
REPLACE sfgz WITH 0 all
IF sfgz1='正常'
inde on rybh to kq&ny
set relation to rybh into 2
repl sfgz with 标准*b.sfbl/100 for &条件1
*****正常计发公式******************xcts>=15*******20160115*************
count for &条件2 to rs1
if rs1>0
   repl sfgz with xcts*日标*b.sfbl/100 for &条件2
*****扣发公式*******************************xcts<15*******20160115*************
   brow field cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',b.cjgzrq:h='参加工作日期',b.工作年限,b.sfbl:h='上浮比例（%）',gj:h='年休假',xcts:h='现场天数',sfgz:h='上浮工资' for &条件2 titl '下列人员现场出勤天数少于标准计发条件：'+条件1+'按规定应扣发上浮工资'
endif
count for &条件3 to rs2
if rs2>0
   repl sfgz with 0 for &条件3
*****扣发公式**********tqj+sj>15*******20160115*************
   brow field cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',b.cjgzrq:h='参加工作日期',b.工作年限,b.sfbl:h='上浮比例（%）',gj:h='年休假',xcts:h='现场天数',sfgz:h='上浮工资' for &条件3 titl '下列人员全月缺勤超过标准计发天数：'+条件3+'按文件规定不发上浮工资'
ENDIF
count for sfgz=0 to rs3
count for sfgz>0 to rs
if rs>0
   INDEX on bmbh+zw+rybh TO kq&ny
   GO top
   brow field cjmc:h='单位',bmmc:h='部门',ryxm:h='姓名',b.cjgzrq:h='参加工作日期',b.工作年限,b.sfbl:h='上浮比例（%）',gj:h='年休假',xcts:h='现场天数',sfgz:h='上浮工资' titl '本月应计发上浮工资'+allt(STR(rs))+'人（其中：扣发上浮工资'+allt(STR(rs1))+'人；不发上浮工资'+allt(STR(rs2))+'人）；工作年限不滿1周年'+allt(STR(rs3))+'人'
ENDIF
ENDIF
REPLACE bj1 WITH xcts*bjjb all
*************1.保健＝保健标准×现场天数****************
yn = MESSAGEBOX("考勤检索情况导出到"+bf+"\“缺勤异常情况”电子表中，现打开使用吗？",4+32,"提示")
IF yn = 6
**************标准Excel读取文件*************
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\缺勤异常情况.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
ENDIF
close all
on key label F1 do grcx
*********还原热键定义*********************
i=1
*********还原全局变量i=1*********************
CLEAR
*************************清屏打开图片库设置封底界面*****************
use ccrq
go 8
壁纸数=fdmax
I=val(图片)
pic=ALLTRIM(图片)
repl 图片 with ALLTRIM(str(I+1))
IF I>=壁纸数
   repl 图片 with '0'
ENDIF
**********用户清新杂志封面封底界面*********
pict_fd='fd'+pic+'.jpg'
_SCREEN.PICTURE = '&pict_fd'
close all
RETURN


******************
PROCEDURE KQgrcx
******************
do while .t.
      do srxm.spr
      sele 2
      go top
      locate for allt(RYXM)=allt(XM) or allt(RYBH)=allt(XM)
      JLH=recn()
      if eof()
        do dhkjg.spx
    M = readkey()
    if M=12 or M=268
      return
    endif
     else
      wait wind '需要查找的人员已找到，请先按[PgDn][PgUp]两个键刷新一下' nowa 
      exit    
     endif
ENDDO
******************
PROCEDURE plsr
******************
select 2
jlh=RECNO()
sb=sbts
xc=xcts
zb=zbgr
yb=ybgr
jr=jrjb
yn = MESSAGEBOX("预统一按当前记录批量输入出勤情况吗？",4+32,"提示")
IF yn = 6
   REPLACE sbts WITH sb,xcts WITH xc,zbgr WITH zb,ybgr WITH yb,jrjb WITH jr FOR sbts=0
ENDIF
go jlh
******************
PROCEDURE SYKQTS
******************
wait wind alltrim(b.RYXM)+'[上月]：'+'中班:'+alltrim(str(d.zbgr))+'夜班:'+alltrim(str(d.ybgr))+'保健:'+alltrim(str(d.Bj1))+'加班:'+alltrim(str(d.JrJB))+';病假:'+alltrim(str(d.BJ))+';'+'事假:'+alltrim(str(d.SJ))+';'+'陪护/产假:'+alltrim(str(d.CJ))+';工休:'+alltrim(str(d.GS))+';探亲:'+alltrim(str(d.TQJ))+';婚假:'+alltrim(str(d.HunJ))+';丧假:'+alltrim(str(d.SangJ))+';拘留:'+alltrim(str(d.JL))+';旷工:'+alltrim(str(d.KG)) nowa at 50,120

******************
PROCEDURE SYKQTS1
******************
wait wind alltrim(b.RYXM)+'[上月]：'+'中班:'+alltrim(str(b.zbgr))+'夜班:'+alltrim(str(b.ybgr))+';病假:'+alltrim(str(b.BJ))+';'+'事假:'+alltrim(str(b.SJ))+';婚假:'+alltrim(str(b.HunJ))+';丧假:'+alltrim(str(b.SangJ))+';旷工:'+alltrim(str(b.KG)) nowa at 50,120