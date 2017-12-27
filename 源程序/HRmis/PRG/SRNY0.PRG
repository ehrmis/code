***************************
* .\SRNY1.PRG
***************************
set date to YMD
set date to long
set stri to 0
 PRIVATE M.CURRAREA , M.TALKSTAT , M.COMPSTAT
 IF SET('TALK') = 'ON'
    SET TALK OFF
    M.TALKSTAT = 'ON'
 ELSE 
    M.TALKSTAT = 'OFF'
 ENDIF 
 M.COMPSTAT = SET('COMPATIBLE')
 SET COMPATIBLE OFF
 M.CURRAREA = SELECT()
 IF  .NOT. WEXIST('_ruz0oak99')
    DEFINE WINDOW _RUZ0OAK99 FROM INT((SROWS() - 6) / 2) , INT((SCOLS() - 18) / 2) TO  ;
         INT((SROWS() - 6) / 2) + 7 , INT((SCOLS() - 40) / 2) + 55 colo 7+/2 FLOAT  ;
         CLOSE SHADOW DOUBLE 
 ENDIF 
 IF WVISIBLE('_ruz0oak99')
    ACTIVATE WINDOW SAME _RUZ0OAK99
 ELSE 
    ACTIVATE WINDOW NOSHOW _RUZ0OAK99
 ENDIF 
ND1 = year(date())
YF1 = 1
ND2 = year(date())
YFtemp = month(date())
@ 1 , 6 get ND1  FONT '宋体' , 11 range 1900 , 9999 picture '9999'
@ 1 , col() say '年' FONT '宋体' , 11 get YF1 FONT '宋体' , 11 range 1 , 12 picture '99'
@ 1 , col() say '月至' FONT '宋体' , 11
@ 1 , 25 get ND2 FONT '宋体' , 11 range 1900 , 9999 picture '9999'
@ 1 , col() say '年' FONT '宋体' , 11 get YFtemp FONT '宋体' , 11 range 1 , 12 picture '99'
@ 1 , col() say '月' FONT '宋体' , 11 
@ 4 , 18 get PD default 1 size 1.5 , 8 , 2 picture '@*HT \! \<确 定' FONT '宋体' , 11
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
if YF1>9
  YF = str(YF1,2)
else
  YF = '0'+str(YF1,1)
endif
ND2 = str(ND2,4)
if YFtemp>9
  YFtemp = str(YFtemp,2)
else
  YFtemp = '0'+str(YFtemp,1)
endif
NY1 = ND2+YFtemp
release window _RUZ0OAK99
select (m.CURRAREA)
if m.TALKSTAT='ON'
  set talk on
endif
if m.COMPSTAT='ON'
  set compatible on
endif
deactivate window _RUZ0OAK99
clear window _RUZ0OAK99
