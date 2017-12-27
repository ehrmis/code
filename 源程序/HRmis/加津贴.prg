close all
set safety off
set talk off
use d:\simis\backup\zr9905
copy to d:\simis\zr992005
use d:\simis\zr992005
close all
for yf1=1 to 7
   yf='0'+str(yf1,1)
sele 2
use kq2005&yf
inde on rybh to kq2005&yf
sele 1
use d:\simis\zr992005
inde on rybh to d:\simis\zr992005
upda on rybh from B repl j&yf with j&yf+b.xcts*16,sfje with sfje+b.xcts*16
endfor
close all
use d:\simis\zr992005
repl gzhj with m01+m02+m03+m04+m05+m06+m07+m08+m09+m10+m11+m12,jjhj with j01+j02+j03+j04+j05+j06+j07+j08+j09+j10+j11+j12,; 
hj with gzhj+jjhj,PJ with round(HJ/MOU,0) for mou>0
go top
brow