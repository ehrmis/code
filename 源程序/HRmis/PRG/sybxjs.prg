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
���=dwbh
��λ=mc
��λ��=��λʧҵ��/100
���۷�=allt(ʧҵ���۷�)
use ��������
loca for ���='&qn'
GRSYL=ʧҵ���գ�/100
BQSYL=sybxbl/100
********A����B���ֿ�����********************************
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
     upda on rybh from B repl �ɷ����� with b.mou
     if file(" xzjs"+LY+".dbf") 
        upda on rybh from F repl �ɷ����� with f.�ɷ�����,jfjs with f.�½ɷѻ���
     endif
     **************У��*************
    * IF uppe(���۷�)='Y'
        repl snjfjs with jfjs all
    * ELSE
    *    upda on rybh from C repl snjfjs with c.jfjs    
    * ENDIF   
     repl snsybx with round(snjfjs*grsyl,1) , dwsybx with round(jfjs*��λ��,1) , sndwsybx with round(snjfjs*��λ��,1) , A with 1 ,;
gzze with round(jfjs*�ɷ�����,1) , gryj with round(sybx*�ɷ�����,1) , dwyj with round(dwsybx*�ɷ�����,1) , grsj with gryj all 
     go top
    do whil !eof()
    if val(LY)=year(date()) and month(date())<=3
        repl gzze with round(snjfjs*�ɷ�����,1) , gryj with round(snsybx*�ɷ�����,1) , dwyj with round(sndwsybx*�ɷ�����,1) , grsj with gryj ,;
jfjs with 0,sybx with 0,dwsybx with 0
    else
    do case
    case �ɷ�����=10 
        repl gzze with round(snjfjs+jfjs*(�ɷ�����-1),1) , gryj with round(snsybx+sybx*(�ɷ�����-1),1) , dwyj with round(sndwsybx+dwsybx*(�ɷ�����-1),1) , grsj with gryj
    case �ɷ�����=11 
        repl gzze with round(snjfjs*2+jfjs*(�ɷ�����-2),1) , gryj with round(snsybx*2+sybx*(�ɷ�����-2),1) , dwyj with round(sndwsybx*2+dwsybx*(�ɷ�����-2),1) , grsj with gryj
    case �ɷ�����=12 
        repl gzze with round(snjfjs*3+jfjs*(�ɷ�����-3),1) , gryj with round(snsybx*3+sybx*(�ɷ�����-3),1) , dwyj with round(sndwsybx*3+dwsybx*(�ɷ�����-3),1) , grsj with gryj
    case �ɷ�����<10 
        repl gzze with round(jfjs*�ɷ�����,1) , gryj with round(sybx*�ɷ�����,1) , dwyj with round(dwsybx*�ɷ�����,1) , grsj with gryj ,;
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
