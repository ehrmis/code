*************************
*源程序代码清单
*************************
*CADmis.prg &&(主控程序)
*************************
_screen.caption='昆钢员工培训管理信息系统'
do setting
=sys(2002)
do form fm
=inkey(8,"M")
=sys(2002,1)
clear window
deac wind 'PROJECT MANAGER'
PUBLIC RK,I,CJ,BM,ND,YF,MD,XM,BH,DM,MC,sslj
RK = 4
I=0
CJ=''
BM=''
ND=year(date())
YF=iif(month(date())>9,str(month(date()),2),'0'+str(month(date()),1))
XM=''
BH=''
DM=''
MC=''
MD=0
sslj=''
clear
set color to /2+
close all
do MAIN.mpr
do form MM
READ EVENTS
close all
clear all
quit