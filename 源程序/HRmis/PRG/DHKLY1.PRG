* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
*  文件名: DHKLY1.PRG <-- 本文件由 UnFoxAll 创建
* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
 SET TALK OFF
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
    DEFINE WINDOW _RUZ0OAK99 FROM INT((SROWS() - 6) / 2) , INT((SCOLS() - 22) / 2) TO  ;
         INT((SROWS() - 6) / 2) + 7 , INT((SCOLS() - 40) / 2) + 48 NOFLOAT  ;
         NOCLOSE SHADOW DOUBLE COLOR W+/G
 ENDIF 
 IF WVISIBLE('_ruz0oak99')
    ACTIVATE WINDOW SAME _RUZ0OAK99
 ELSE 
    ACTIVATE WINDOW NOSHOW _RUZ0OAK99
 ENDIF 
 @ 1 , 8 SAY '选择工资输入方式' font '宋体',13
 STORE 1 TO DHLY1
 @ 3 , 6 GET DHLY1 SIZE 2 , 8 , 7 FUNCTION '*h 手工;自动' font '宋体',10
 READ 
 CLEAR WINDOW
 RETURN 
*
