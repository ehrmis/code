****************�ۻ����ݸ��£�Update Build 20170907****************
IF files("ry20171018.dbf")
yn = MESSAGEBOX("�Ƿ���Ҫ���ݸ��£�Update Build 20170907��",4+32,"��ʾ")
IF yn = 6
CLOSE ALL
USE ��
REPLACE userjb WITH 9 FOR 'Admin'$user
USE ry20171018 EXCLUSIVE
zap
APPEND FROM ryzk
BROWSE TITLE '�����ۻ����£�Update Build 20171018 ��Ա״��ά��������'
**********************************************Build �������ڡ��ֶ�*************************************************************************************************20161008***********************
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

