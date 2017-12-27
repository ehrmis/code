yf=''
nd=''
yf1=1
yftemp=month(date())
close all
clear 
do srny
yf0=val(yf)
yf1=val(yf)
*****收集基础数据*****
use kqk
copy to temp
  use temp excl
  zap
  if vartype(yftemp)='C'   
     yftemp=val(yftemp)
  endif
  yf1=yf0
    for ii=yf1 to yftemp+1      
    yf=iif(yf1>9,str(yf1,2),'0'+str(yf1,1))
    ny=nd+yf
    if file('KQ'+NY+'.dbf')
       appe from kq&ny for xcts>0
    else
       exit
    endif   
    yf1=yf1+1
    endfor
    inde on rybh to temp
    tota on rybh to xcts
    do form xcts
  