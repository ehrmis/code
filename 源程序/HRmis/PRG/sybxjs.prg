****************************
* sybxjs.prg (Build20160720)
****************************
close all
set date to YMD
set date to long
set safety off
qn=str(val(LY)-2,4)
LY1=str(val(LY)-1,4)
use sbdmk
编号=dwbh
单位=mc
单位率=单位失业率/100
补扣否=allt(失业补扣否)
use 基础数据
loca for 年份='&qn'
GRSYL=失业保险％/100
BQSYL=sybxbl/100
********A区、B区分开缴纳********************************
use sy_bx
copy to sy&LY stru
use sy&LY
appe from zr1&LY for ygxz='01'
close all
if file(" xzjs"+LY+".dbf") 
   select 6
   use  xzjs&LY 
   inde on rybh to xzjs&LY
endif   
     sele 5
     use zr1&LY1
     inde on rybh to zr1&LY1
     sele 4
     use ryzk
     inde on rybh to ryzk
    * sele 3
    * use zr1&QN
    * inde on rybh to zr1&QN
     sele 2
     use zr1&LY
     inde on rybh to zr1&LY
     sele 1
     use sy&LY 
     inde on rybh to sy&LY
     upda on rybh from E repl sybx with e.sybx,jfjs with e.jfjs
     upda on rybh from D repl sfz with d.sfz,xb1 with d.xb1,csrq with d.csrq,cjgzrq with d.cjgzrq,cjdm with d.cjdm,cjmc with d.cjmc,bmmc with d.bmmc,bmbh with d.bmbh,zz with d.zz,tsry with d.tsry
     upda on rybh from B repl 缴费月数 with b.mou
     if file(" xzjs"+LY+".dbf") 
        upda on rybh from F repl 缴费月数 with f.缴费月数,jfjs with f.月缴费基数
     endif
     **************校正*************
    * IF uppe(补扣否)='Y'
        repl snjfjs with jfjs all
    * ELSE
    *    upda on rybh from C repl snjfjs with c.jfjs    
    * ENDIF   
     repl snsybx with round(snjfjs*grsyl,1) , dwsybx with round(jfjs*单位率,1) , sndwsybx with round(snjfjs*单位率,1) , A with 1 ,;
gzze with round(jfjs*缴费月数,1) , gryj with round(sybx*缴费月数,1) , dwyj with round(dwsybx*缴费月数,1) , grsj with gryj all 
     go top
    do whil !eof()
    if val(LY)=year(date()) and month(date())<=3
        repl gzze with round(snjfjs*缴费月数,1) , gryj with round(snsybx*缴费月数,1) , dwyj with round(sndwsybx*缴费月数,1) , grsj with gryj ,;
jfjs with 0,sybx with 0,dwsybx with 0
    else
    do case
    case 缴费月数=10 
        repl gzze with round(snjfjs+jfjs*(缴费月数-1),1) , gryj with round(snsybx+sybx*(缴费月数-1),1) , dwyj with round(sndwsybx+dwsybx*(缴费月数-1),1) , grsj with gryj
    case 缴费月数=11 
        repl gzze with round(snjfjs*2+jfjs*(缴费月数-2),1) , gryj with round(snsybx*2+sybx*(缴费月数-2),1) , dwyj with round(sndwsybx*2+dwsybx*(缴费月数-2),1) , grsj with gryj
    case 缴费月数=12 
        repl gzze with round(snjfjs*3+jfjs*(缴费月数-3),1) , gryj with round(snsybx*3+sybx*(缴费月数-3),1) , dwyj with round(sndwsybx*3+dwsybx*(缴费月数-3),1) , grsj with gryj
    case 缴费月数<10 
        repl gzze with round(jfjs*缴费月数,1) , gryj with round(sybx*缴费月数,1) , dwyj with round(dwsybx*缴费月数,1) , grsj with gryj ,;
snjfjs with 0,snsybx with 0,sndwsybx with 0
    endcase
    endif
    skip
    enddo
SORT ON BMBH,RYBH to temp
close all
use SY&LY excl
zap
appe from temp
close all   
retu
