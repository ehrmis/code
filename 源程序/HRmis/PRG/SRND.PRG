***************************
* .\SRND.SPR
***************************
set date to YMD
set date to long
set stri to 0
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
  define window _RUZ0OAK99 from int((srows()-6)/2) , int((scols()-40)/2);
 to int((srows()-6)/2)+5 , int((scols()-40)/2)+39 colo 7+/2 float;
 close shadow double
endif
if wvisible('_ruz0oak99')
  activate window same _RUZ0OAK99
else
  activate window noshow _RUZ0OAK99
endif
nd1=year(date())
@ 1 , 12 say '请输入年份:        ' size 1 , 11 , 0
@ 1 , 23 get ND1 range 1900 , 9999 picture '9999'
@ 1 , col() say '年'
@ 3 , 14 get PD default 1 size 1 , 10 , 1 picture '@*HT \!  \<确  定'
if  not wvisible('_ruz0oak99')
  activate window _RUZ0OAK99
endif
read cycle
FF = readkey()
read
if FF=12 or FF=268
  return
endif
ND = str(ND1,4)
release window _RUZ0OAK99
select (m.CURRAREA)
if m.TALKSTAT='ON'
  set talk on
endif
if m.COMPSTAT='ON'
  set compatible on
endif
