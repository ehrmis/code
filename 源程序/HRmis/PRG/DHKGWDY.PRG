***************************
* .\DHKGWDY.PRG
***************************
*set shadows on
*define window WIN_DHKGWD from 5 , 24 to 13 , 54 color 7+/1 float shadow;
* panel noclose
*activate window WIN_DHKGWD
private m.CURRAREA , m.TALKSTAT , m.COMPSTAT
if set('TALK')='ON'
  set talk off
  m.TALKSTAT = 'ON'
else
  m.TALKSTAT = 'OFF'
endif
m.COMPSTAT = set('COMPATIBLE')
set compatible off
m.CURRAREA = select()
if  not wexist('_ruz0oak99')
  define window _RUZ0OAK99 from int((srows()-6)/2) , int((scols()-18)/2);
 to int((srows()-6)/2)+7 , int((scols()-40)/2)+55 color 7+/2 nofloat;
 noclose shadow double
endif
if wvisible('_ruz0oak99')
  activate window same _RUZ0OAK99
else
  activate window noshow _RUZ0OAK99
endif
store 1 to DHGWDY
@ 1 , 13 say '请  选  择  岗  位'
@ 3 , 6 get DHGWDY size 2 , 12 , 8 function '*h 操  作;管  理'
read
deactivate window WIN_DHKGWD
release window _RUZ0OAK99
select (m.CURRAREA)
if m.TALKSTAT='ON'
  set talk on
endif
if m.COMPSTAT='ON'
  set compatible on
endif
return 
� 
