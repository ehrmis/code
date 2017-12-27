***************************
* .\KQSR.PRG(Build 20041007)
***************************
on escape
set escape off
on key label F1 do grsr
close all
set safety off
set talk off
_SCREEN.PICTURE = ''
TS=0
dh=''
dycj=''
dybm=''
ND = '  '
YF = '  '
XH1 = 'N'
xm=''
LYNY = ''
dhly=''
set colo to 7+/1
clear
clear windows
set colo to 7+/2
set confirm off
**准六位年月****①本月自动处理**＊************
ND1 = year(date())
YF1 = month(date())
@ 10 , 40 say '请输入年月:        ' FONT '宋体' , 11 size 1 , 11 , 0
@ 10 , col()+2 get ND1 FONT '宋体' , 11 range 1900 , 9999 picture '9999'
@ 10 , col() say '年' FONT '宋体' , 11 get YF1 FONT '宋体' , 11 range 1 , 12 picture '99'
@ 10 , col() say '月'FONT '宋体' , 11
read 
ND = str(ND1,4)
if YF1>9
  YF = str(YF1,2)
else
  YF = '0'+str(YF1,1)
endif
clear
LYNY = ND+YF
****************②前月自动处理****************
if val(yf)<3
  JJND = str(val(nd)-1,4)
  JJYF = str(10+val(yf),2)
else
  JJND = ND
  JJYF = iif(val(YF)-2>9,str(val(YF)-2,2),'0'+str(val(YF)-2,1))
endif
JJLY = JJND+JJYF
****************③上月自动处理****************
if val(YF)-1>9
  jjYF = str(val(YF)-1,2)
else
  jjYF = iif(val(YF)-1=0,'12','0'+str(val(YF)-1,1))
endif
if month(date())=1
   syjj=str(val(nd)-1,4)+jjYF
else   
   syjj = ND+jjYF
endif
*****************④考勤日历表自动处理**********
use rlk
go top
loca for rq=allt(LYNY) and ZDGR>0
if eof() 
YN = MESSAGEBOX('不必要修改工资程序考勤日历表？！',4+32,'提示')
if yn=7
  do dhkly
  if dhly=1
     brow titl '请认真审核工资程序考勤日历表' fiel rq:h='年月',rlts:h='日历天数',gxts:h='公休天数',zdgr:h='制度工日'
  else
    do rlk  
  endi
close all  
use rlk
go top
locate for RQ=LYNY
TS = ZDGR
else
TS = 21  
endif    
else
TS = ZDGR
endif
*****************************************************
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
if !file('KQ'+LYNY+'.DBF')
  use KQK
  COPY TO &sjlj.KQ&LYNY stru
  USE KQ&LYNY
  append from &sjlj.dm1\RYZK for zgqk='01'
  close all
  select 2
  use jj&JJLY
  inde on rybh to jj
  select 1
  use kq&LYNY
  inde on rybh to kq
  update on RYBH from B replace a->jang with b->jang1+b->jang2
endif
close all
select 2
use &sjlj.dm1\RYZK
index on RYBH to ryzk
select 1
use kq&LYNY
inde on rybh to kq&LYNY
update on RYBH from B replace CJDM with b.CJDM , BMBH;
 with b.BMBH , CJMC with b.CJMC , BMMC with b.BMMC
close all
select 3
use &sjlj.dm1\RYZK
index on RYBH to &sjlj.dm1\ryzk
select 1
use SMK1
select 2
use kQ&LYNY
set relation to RYBH into 3
if TS=0
   TS = 21 
endif 
replace SBTS with TS all
replace BJ with TS for c.RYFL='33'
index on BMBH+RYBH to &sjlj.kq&LYNY
go top
NUM1 = recno()
go bottom
NUM2 = recno()
go top
if dh<>1
  go bottom
  if dh=2
     set filter to left(BMBH,2)=T123 
  else
     set filter to BMBH=T123 
  endif
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
  LLL1 = 0
  clear
  do while  not eof()
    @ KKK1 , 30+LLL1 say CHINES  font '宋体' , 14
    KKK1 = KKK1+2
    if KKK1>20
      KKK1 = 5
      LLL1 = LLL1+25
    endif
    skip
  enddo
  @ 21 , 50 say 'PgDn=下一条记录,PgUp=上一条记录' font '宋体' , 14
  @ 23 , 50 say 'F1=查找指定人员,Esc=存盘退出' font '宋体' , 14
  do while upper(XH2)='N'
    select 1
    go top
    KKK1 = 5
    LLL1 = 0
    do while  not eof()
      RYHDS = FIELD_NAME
      TM123 = HHH1
      select 2
      if upper(trim(RYHDS))='RYBH' or upper(trim(RYHDS))='RYXM' or upper(trim(RYHDS))='BMMC';
 or upper(trim(RYHDS))='CJMC'
        RYHDS=&RYHDS
        @ KKK1 , 40+LLL1 say RYHDS  font '宋体' , 14
      else
        @kkk1,40+lll1 get &ryhds  font '宋体' , 14 PICT &TM123
      endif
      KKK1 = KKK1+2
      if KKK1>20
        KKK1 = 5
        LLL1 = LLL1+25
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
  clear
  FF = readkey()
  if FF=12 or FF=268
    exit
  endif
enddo
close all
select 2
use kq&LYNY
inde on rybh to kq&LYNY
select 1
use &sjlj.dm1\RYZK
index on RYBH to ryzk
update on RYBH from B replace BJJB with b.BJJB , BK with b.BJ1*BJJB , FZSD;
 with b.FZSD
index on RYBH to ryzk
select 2
set relation to RYBH into 1
replace JRJB with 0 , zbgr with 0  , ybgr with 0 for '处'$a.ZW1
sele 1
index on BMBH+RYBH to &sjlj.dm1\RYZK
close all
use kq&LYNY
index on BMBH+RYBH to &sjlj\kq&LYNY
   count for zgqk='01'to rs
   rs=str(rs,4)
   count for bj>0 to bjj
   bjj=str(bjj,4)
   count for sj>0 to sjj
   sjj=str(sjj,4)   
if dh<>1
   if dh=2
   count for left(BMBH,2)=T123 and zgqk='01' to rs1
   rs1=str(rs1,4)
   count for left(BMBH,2)=T123 and bj>0 to bjj1
   bjj1=str(bjj1,4)
   count for left(BMBH,2)=T123 and sj>0 to sjj1
   sjj1=str(sjj1,4)
   go top
   loca for left(BMBH,2)=T123
   cjm=allt(cjmc)
   brow part 40 titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs+'人，病假='+bjj+'人；事假='+sjj+'人；其中，'+cjm+'：'+rs1+'人，病假='+bjj1+'人；事假='+sjj1+'人  按[ Esc ]键退出'
   else
   count for BMBH=T123 and zgqk='01' to rs1
   rs1=str(rs1,4)
   count for BMBH=T123 and bj>0 to bjj1
   bjj1=str(bjj1,4)
   count for BMBH=T123 and sj>0 to sjj1
   sjj1=str(sjj1,4)
   go top
   loca for BMBH=T123
   bm=allt(bmmc)
   brow part 40 titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs+'人，病假='+bjj+'人；事假='+sjj+'人；其中，'+bm+'：'+rs1+'人，病假='+bjj1+'人；事假='+sjj1+'人  按[ Esc ]键退出'
   endif
else
   go top
   brow part 40 titl '请认真审核'+nd+'年'+allt(str(val(yf)))+'月份考勤：共'+rs+'人，病假='+bjj+'人；事假='+sjj+'人  按[ Esc ]键退出'
endif
close all
clear
_SCREEN.PICTURE = '&pic_fd'
i=99
return


proc rlk

rlts1=day(date())
if rlts1=1
for rl=1 to 31
    rlts1=rlts1+1
    if day(date()+rl)=1
       exit
    endif
endfor
gxts1=1
SET DATE TO MDY
if month(date())=1
   nd=str(val(nd)-1,4)
endif
for dd1 = 1 to rlts1
    dd=iif(dd1>9,str(dd1,2),'0'+str(dd1,1))
    if dow(ctod('&jjyf./'+'&dd./'+'&nd'))=1 or dow(ctod('&jjyf./'+'&dd./'+'&nd'))=7
       gxts1=gxts1+1
    endif
endfor
if month(date())=1
   nd=str(val(nd)+1,4)
endif
SET DATE TO YMD
SET DATE TO long
use rlk excl
delete for rq=LYNY
pack
appe blank
go bott
repl rq with LYNY ,rlts with rlts1-1,gxts with gxts1-1,zdgr with rlts-gxts
brow titl '请认真审核工资程序考勤日历表' fiel rq:h='年月',rlts:h='日历天数',gxts:h='公休天数',zdgr:h='制度工日'
if day(date())=1
   =MESSAGEBOX("今天不是"+allt(str(val(yf)))+"月1号？！，请在桌面右下角将系统时间调为正常日期......", 16,"提示")
endif
quit
else
=MESSAGEBOX("今天是"+str(rlts1,2)+"号，请在桌面右下角将系统时间调为［上月1号］......", 16,"提示")
quit
endif