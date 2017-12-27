set safety off
set colo to 7+/2
close all
use d:\tyh\dm1\ryzk
go top
dw=dwdm
close all
pf=space(1)
do while .T.
@ 10,25 say '请输入拷贝盘符(E,F,G,H,I):' font "宋体",16  get pf  font "宋体",16 
read
mm1=1
mm2=month(date())
@ 12,25 say '请输入拷贝月份:' font "宋体",16  get mm1  font "宋体",16 
@ 12,col() say '月至 ' font "宋体",14  get mm2  font "宋体",14 
@ 12,col() say '月'
read
pf=uppe(allt(pf))+':'
if pf='E:' or pf='F:' or pf='G:' or pf='H:' or pf='I:'
exit
endi
endd
nd=str(year(date()),4)
for ii=mm1 to mm2 
yf=iif(ii>9,str(ii,2),'0'+str(ii,1))
clear
@ 10,20 say '正在拷贝'+nd+'年'+str(val(yf),2)+'月份数据到'+pf+'盘... ...' font "宋体",16 
dwny=dw+right(nd,2)+yf
nd1=nd
yf1=iif(ii-1>9,str(ii-1,2),'0'+str(ii-1,1))
if yf1='00'
   yf1='12'
   nd1=str(val(nd)-1,4)
endif
if file('d:\tyh\gz'+nd+yf+'.dbf')   
use d:\tyh\gz&nd.&yf
copy to &pf\backup\gz&nd.&yf
endi
if file('d:\tyh\kq'+nd+yf+'.dbf')
use d:\tyh\kq&nd.&yf
copy to &pf\backup\kq&nd.&yf
endi
if file('d:\tyh\kg'+dwny+'.dbf')
use d:\tyh\kg&dwny
copy to &pf\backup\kg&dwny
endi
if file('d:\tyh\ty'+nd+yf+'.dbf')
use d:\tyh\ty&nd.&yf
copy to &pf\backup\ty&nd.&yf
endi
if file('d:\tyh\ry'+nd+yf+'.dbf')
use d:\tyh\ry&nd.&yf
copy to &pf\backup\ry&nd.&yf
endi
endfor
close all
quit