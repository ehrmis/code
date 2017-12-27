****************累积数据更新：Update Build 20170907****************
IF files("ry20171018.dbf")
yn = MESSAGEBOX("是否需要数据更新：Update Build 20170907？",4+32,"提示")
IF yn = 6
CLOSE ALL
USE ＇
REPLACE userjb WITH 9 FOR 'Admin'$user
USE ry20171018 EXCLUSIVE
zap
APPEND FROM ryzk
BROWSE TITLE '数据累积更新：Update Build 20171018 人员状况维护：户籍'
**********************************************Build 更新日期、字段*************************************************************************************************20161008***********************
COPY TO data\ryzk
CLOSE ALL
USE ry20171018 EXCLUSIVE
zap
APPEND FROM ryzk1
COPY TO data\ryzk1

FOR ny1=201701 TO 201712
    ny=STR(ny1,6)
IF files('ry'+ny+'.dbf')
   USE ry20171018 EXCLUSIVE
   ZAP
   APPEND FROM ry&ny
   COPY TO ry&ny
ENDIF
******************************    
*!*	IF files('kq'+ny+'.dbf')
*!*	   USE kqk EXCLUSIVE
*!*	   ZAP
*!*	   APPEND FROM kq&ny
*!*	   COPY TO kq&ny
*!*	ENDIF
*!*	IF files('lw'+ny+'.dbf')
*!*	   USE lwpqgzk EXCLUSIVE
*!*	   ZAP
*!*	   APPEND FROM lw&ny
*!*	   COPY TO lw&ny
*!*	ENDIF
*!*	IF files('bj'+ny+'.dbf')
*!*	   USE dcbj EXCLUSIVE
*!*	   ZAP
*!*	   APPEND FROM bj&ny
*!*	   COPY TO bj&ny
*!*	ENDIF
ENDFOR

*!*	FOR n1=2014 TO 2017
*!*	    nn=STR(n1,4)
*!*	IF files('zr'+nn+'.dbf')
*!*	   USE gzze EXCLUSIVE
*!*	   ZAP
*!*	   APPEND FROM zr&nn
*!*	   COPY TO zr&nn
*!*	ENDIF    
*!*	IF files('zr1'+nn+'.dbf')
*!*	   USE gzze1 EXCLUSIVE
*!*	   ZAP
*!*	   APPEND FROM zr1&nn
*!*	   COPY TO zr1&nn
*!*	ENDIF
*!*	IF files('yl'+nn+'.dbf')
*!*	   USE ylbxk EXCLUSIVE
*!*	   ZAP
*!*	   APPEND FROM yl&nn
*!*	   COPY TO yl&nn
*!*	ENDIF
*!*	ENDFOR
ENDIF
ENDIF

