***************************
* .\DHK14.PRG
***************************
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
store 1 to DH14
@ 2 , 6 get DH14 size 2 , 9 , 2 function '*h œ‘ æ;¥Ú”°;∑µªÿ'
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