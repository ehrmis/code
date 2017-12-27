***************************
*jjSR1.PRG( Build20081015 )
***************************
on escape
set escape off
on key label F1 do grsr
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
T123 = '  '
close all
TS = 0
TS1 = 0
XH1 = 'N'
set confirm off
ND1 = year(date())
YF1 = month(date())
*if yf1=0
*   yf1=12
*   nd1=nd1-1
*endif   
clear
do srny.spx
NY = ND + YF
if !file('KQ'+NY+'.DBF')
  use KQK
  COPY TO KQ&NY stru
  USE KQ&NY
  append from RYZK for zgqk='01'  
endif
*********************初始化****************
close all
select 2
use ryzk
index on RYBH to ryzk
select 1
use kq&NY
inde on rybh to kq&NY
update on RYBH from B replace CJDM with b.CJDM , BMBH;
 with b.BMBH , CJMC with b.CJMC , BMMC with b.BMMC ,erprybh with b.erprybh
close all
do dhkdy
do case
   case dh=2
  T123 = dyCJ
  case dh=3
  T123 = dybm
endcase 
use kq&NY
on key label F1 do grcx
do case
case dh=1
clear
go top
brow titl '请认真输入'+nd+'年'+allt(str(val(yf)))+'月份绩效工资：共'+allt(str(recc()))+'人 按[ F1 ]键查找 按[ Esc ]键退出';
fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',jang1:h='绩效工资',jang2:h='单项奖'
count for jang2>0 to rs1
sum jang2 to xyj
if rs1>0
   pjxyj=round(xyj/rs1,1)
   pjxyj=allt(str(pjxyj,10,1))
else
   xyj=0
   pjxyj='0'
endif
count for jang1>0 to rs
sum jang1 to zhj1
zhj=allt(str(zhj1,10,2))
if rs>0
   pjzhj=round(zhj1/rs,1)
   pjzhj=allt(str(pjzhj,10,1))
   pjj=round((zhj1+xyj)/rs,1)
   pjj=allt(str(pjj,10,1))   
   rs=str(rs,4)
else
   zhj='0'
   pjzhj='0'
   pjj='0'
   rs='0'
endif
go top
brow titl '共'+rs+'人，绩效工资='+zhj+'元(人均='+pjzhj+'元/人)，单项奖='+allt(str(xyj,10,2))+'元(人均='+pjxyj+'元/人)，合计＝'+allt(str(zhj1+xyj,10,2))+'元(人均='+pjj+'元/人)';
fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',jang1:h='绩效工资',jang2:h='单项奖' 
case dh=2
count for cjdm=T123 to rs
go top
brow titl '请认真输入'+nd+'年'+allt(str(val(yf)))+'月份绩效工资：共'+allt(str(rs))+'人 按[ F1 ]键查找 按[ Esc ]键退出';
fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',jang1:h='绩效工资',jang2:h='单项奖' for cjdm=T123
   count for cjdm=T123 and jang2>0 to rs1
   sum jang2 to xyj for cjdm=T123
   if rs1>0
       pjxyj=round(xyj/rs1,0)
       pjxyj=allt(str(pjxyj,10,1))
   else
       xyj=0
       pjxyj='0'
   endif
count for cjdm=T123 and jang1>0 to rs 
   sum jang1 to zhj1 for cjdm=T123
   zhj=allt(str(zhj1,10,2))
   if rs>0
   pjzhj=round(zhj1/rs,1)
   pjzhj=allt(str(pjzhj,10,1))
   pjj=round((zhj1+xyj)/rs,1)
   pjj=allt(str(pjj,10,1))   
   rs=str(rs,4)
else
   zhj='0'
   pjzhj='0'
   pjj='0'
   rs='0'
endif
   go top
   loca for cjdm=T123   
   brow titl '共'+rs+'人，绩效工资='+zhj+'元(人均='+pjzhj+'元/人)，单项奖='+allt(str(xyj,10,2))+'元(人均='+pjxyj+'元/人)，合计＝'+allt(str(zhj1+xyj,10,2))+'元(人均='+pjj+'元/人)';
fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',jang1:h='绩效工资',jang2:h='单项奖' for cjdm=T123
set filter to
case dh=3
count for bmbh=T123 to rs
go top
brow titl '请认真输入'+nd+'年'+allt(str(val(yf)))+'月份绩效工资：共'+allt(str(rs))+'人 按[ F1 ]键查找 按[ Esc ]键退出';
fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',jang1:h='绩效工资',jang2:h='单项奖' for bmbh=T123
   count for bmbh=T123 and jang2>0 to rs1
   sum jang2 to xyj for bmbh=T123
   if rs1>0
      pjxyj=round(xyj/rs1,0)
      pjxyj=allt(str(pjxyj,10,1))
   else
      xyj=0
      pjxyj='0'     
   endif
count for bmbh=T123 and jang1>0 to rs  
   sum jang1 to zhj1 for bmbh=T123
   zhj=allt(str(zhj1,10,2))
if rs>0
   pjzhj=round(zhj1/rs,1)
   pjzhj=allt(str(pjzhj,10,1))
   pjj=round((zhj1+xyj)/rs,1)
   pjj=allt(str(pjj,10,1))   
   rs=str(rs,4)
else
   zhj='0'
   pjzhj='0'
   pjj='0'
   rs='0'
endif
   go top
   loca for bmbh=T123   
brow titl '共'+rs+'人，绩效工资='+zhj+'元(人均='+pjzhj+'元/人)，单项奖='+allt(str(xyj,10,2))+'元(人均='+pjxyj+'元/人)，合计＝'+allt(str(zhj1+xyj,10,2))+'元(人均='+pjj+'元/人)';
 fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',jang1:h='绩效工资',jang2:h='单项奖' for bmbh=T123
set filter to
endcase
*******自动筛选ERP考勤数据*********
repl ERP编号 with erprybh,中班 with zbgr,夜班 with ybgr,节日加班 with jrjb,病假 with bj,事假 with sj,产假 with cj,工伤 with gs,探亲假 with tqj,;
婚丧假 with hsj,拘留 with jl,旷工 with kg,年休假 with gj,保健费 with bj1*bjjb,jang with jang1+jang2,上浮工资 with sfgz,绩效工资 with jang,房租水电 with fzsd,互助金 with fzj all
ERPKQ='ryxm,ERP编号,中班,夜班'
J = 8
go top
do while j<24
  tt=fiel(j)
IF type(field(j))='N'
sum &tt to stt
if stt>0
     ERPKQ=allt(ERPKQ)+','+'&tt'
endif
ENDIF
  j=j+1
enddo
************导出ERP考勤************
index on BMBH+RYBH to kq&NY
 yn = MESSAGEBOX("需要导出"+ND+"年"+allt(str(val(YF)))+"月份绩效工资数据吗？",4+32,"提示")
IF yn = 6 
 old_dirs='d:'
Declare integer Qdir_m In Qdir.dll string @ck,;
integer yy,;
string cc
*注册Qdir_m()，@ck=返回的文件夹路径，yy=文件夹类别,cc=提示
*==========================================================
ck=spac(64)      &&存放文件夹路径
yy=1
cc='请选择好文件备份目录,单击确定按钮。'
xx=Qdir_m(@ck,yy,cc)   
*调用Qdir_m()，选择的结果在ck中返回
dirs=allt(subs(ck,2,64))
if len(dirs)=0
   dirs=old_dirs
endif
if right(dirs,1)<>'\'
   dirs=dirs+'\'
endif
wjm='KQ'
copy to &dirs.&wjm.&nd.&yf TYPE fox2x
wjm='ERP考勤奖金'
copy to &dirs.&wjm fiel &ERPKQ for zbgr+ybgr+jrjb+bj+sj+cj+gs+tqj+hsj+jl+kg+bj1+gj+sfgz+JANG+FZSD+KK+FZJ>0 and ygxz='01' and zgqk='01' type xl5
 =messagebox("本月ERP绩效工资数据已成功导出！！",48,"恭喜")  
ENDIF
***********************************
clear
 _SCREEN.WINDOWSTATE = 2
close all
use ccrq
go 8
repl 图片 with str(val(图片)+1,2)
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
ND=str(year(date()),4)
YF=iif(month(date())>9,str(month(date()),2),'0'+str(month(date()),1))
close all
retu
