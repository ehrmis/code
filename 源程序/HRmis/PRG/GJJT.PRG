*****************************
* GJJT.prg( Build 20050822 )
*****************************
close all
use dmk
jzy=¶ÔÕÊ»ù×¼ÔÂ
¶¨Î»=Ñ¹¸ÐÖ½¶¨Î»
close all
ny=nd+yf
sy=iif(month(date())>9,str(month(date()),2),'0'+str(month(date()),1))
sn=str(year(date())-1,4)+"Äê"+allt(jzy)+"ÔÂ"
if file('kg'+dm111+right(nd,2)+yf+'.dbf')
   jn=nd+"Äê"+allt(str(val(yf)))+"ÔÂ"
   if file('kg'+ny+'.dbf')
      gjjny=ny
   else
      gjjny=dm111+right(nd,2)+yf
   endif 
else
   jn=nd+"Äê"+allt(str(val(yf)-1))+"ÔÂ"
    if file('kg'+nd+sy+'.dbf')
       gjjny=nd+sy
    else
       gjjny=dm111+right(nd,2)+sy
    endif   
endif
S0="                                           ×¡·¿¹«»ý½ð¸öÈË½á´æºË¶Ôµ¥"
S2="©°©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ð©¤©¤©¤©¤©Ð©¤©¤©¤©¤©¤©¤©¤©Ð©¤©¤©¤©¤©¤©¤©¤©Ð©¤©¤©¤©¤©¤©¤©¤©Ð©¤©¤©¤©¤©¤©¤©¤©Ð©¤©¤©¤©¤©¤©¤©¤©¤©¤©´"
S3="©¦      ¸öÈËÕÊºÅ      ©¦ÈËÔ±ÐÕÃû©¦ ÆÚ³õÓà¶î(Ôª) ©¦ ÆÚ¼ä½ÉÄÉ(Ôª) ©¦ ±¾ÆÚÀûÏ¢(Ôª) ©¦ ÀÛ¼Æ±¾Ï¢(Ôª) ©¦    ±¸     ×¢     ©¦"
S4="©À©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©È"
S5="©¸©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¼"
***************************************
dh=''
dycj=''
dybm=''
do fjdy
***ÃæÏòÃüÁî±¨±í³ÌÐò·Ö¼¶´òÓ¡½â¾ö·½°¸****
do case
 case dh=1
 use kg&gjjny
 copy to temp for gjjz<>0
 case dh=2
 use kg&gjjny
 copy to temp for cjdm=dycj.and.gjjz<>0
 case dh=3
 use kg&gjjny
 copy to temp for bmbh=dybm.and.gjjz<>0
endcase
set device to print
set print on
set print font "ËÎÌå",8
*********************¡ø****************
clear
set colo to 0/7
***Õë¶Ôfjdy.prgÑÕÉ«ÉèÖÃÎÊÌâ***********
 sele 2
 use ryzk
 inde on rybh to ryzk
 sele 1
 use temp
 set relation to rybh into 2
 go top
 xh=0
 do while !eof()
  do while !eof()
  do case
  case dh=1
       WAIT WINDOW NOWAIT 'ÕýÔÚ´òÓ¡×¡·¿¹«»ý½ð¸öÈË½á´æºË¶Ôµ¥<¹²'+allt(str(recc()))+'ÈË>... ...' 
  case dh=2
       WAIT WINDOW NOWAIT 'ÕýÔÚ´òÓ¡'+cjdm+'£º'+allt(cjmc)+'<¹²'+allt(str(recc()))+'ÈË>... ...' 
  case dh=3
       WAIT WINDOW NOWAIT 'ÕýÔÚ´òÓ¡'+cjdm+'£º'+allt(cjmc)+'¡¢'+allt(bmmc)+'<¹²'+allt(str(recc()))+'ÈË>... ...'       
  endcase
  if mod(xh,(22-¶¨Î»)/2)=0
     @ 0,0 say S0
  else
     @prow()+2,0 say S0
  endif  
  S1=" ³µ¼ä£º"+cjmc+"   ²¿ÃÅ£º"+bmmc+"                              ºË¶ÔÆÚ¼ä£º"+sn+"ÖÁ"+jn
  @prow()+1,0 say S1
  @prow()+1,0 say S2
  @prow()+0.8,0 say S3
  @prow()+0.8,0 say S4
  @prow()+0.8,0 say "©¦"
  @prow(),pcol() say gjjzh
  @prow(),pcol() say "©¦"
  @prow(),pcol() say RYxm
  @prow(),pcol() say "©¦"
  @prow(),pcol() say b.ydhj1 PICT "@Z 999,999,999.99"
  @prow(),pcol() say "©¦"
  @prow(),pcol() say snjn+bnjn PICT "@Z 999,999,999.99"
  @prow(),pcol() say "©¦"
  @prow(),pcol() say bnlx PICT "@Z 999,999,999.99"
  @prow(),pcol() say "©¦"
  @prow(),pcol() say gjjz+bnlx PICT "@Z 999,999,999.99"
  @prow(),pcol() say "©¦"
  @prow(),pcol() say "±¾´ÎÀûÏ¢½áÖÁ"+allt(jzy)+"ÔÂ·Ý ©¦"
  @prow()+0.8,0 say S5
  skip
  xh=xh+1
  if mod(xh,(22-¶¨Î»)/2)=0
    exit
  endif
  enddo
 enddo
close all
delete file temp.dbf
set device to screen
?''
set print off
set print to
clear
retu
ÿ