******************
*DKGZ.prg
******************
on error retu
nd=str(year(date()),4)
yf=iif(month(date())>9,str(month(date()),2),'0'+str(month(date()),1))
ny=nd+yf
ny1=right(ny,4)
wait window '正在处理数据，请稍候... ...' nowait
use dmk
if right(allt(退休管理),1)<>'\'
   sslj=allt(退休管理)+'\'
else
   sslj=allt(退休管理)
endif   
use &sslj.dkgz&ny1
copy to dkgz for kdf+kdyf+kdhf+kssf+kqt>0
use dkgz 
repl kkhj with kdf+kdyf+kdhf+kssf+kqt all
inde on dwdm to dkgz
tota on dwdm to dkhj
do form dkgz
