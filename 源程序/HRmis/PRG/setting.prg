**********************************************************************
*SETTING.PRG &&(ͨ��ϵͳ���л������� Build 20151203)
**********************************************************************
set safety off
set talk off
set palette off
set status bar on
set notify on
*ϵͳ��Ϣ��ʾ�� Visual FoxPro �����ڵײ���ͼ��״̬���������ǻ����ַ���״̬����
SET CLOCK status
set escape on
set keycomp to window
* to DOS �Ի����е�Ĭ�ϰ�ť���н��㣬������۲��仯������ʾִ�У���֮������ set keycomp to windows�����е��ӱ�ĵ�Ԫ����ճ�����ܡ�****20151124*********** 
set carry on
*�����еĹ������д����¼�¼ʱ������ǰ��¼�������ֶε����ݸ��Ƶ��� INSERT�� APPEND �� BROWSE ������¼�¼�� 
set confirm on
set exact on
*���������ʽ�Ľ϶̵�һ�����ұ߼��Ͽո����(0)�ֽڣ���ʹ����ϳ����ʽ�ĳ�����ƥ�䡣���ǣ��ڱȽ��е��κα��ʽβ���Ŀո�����ֽڶ�������
set near on
set ansi off
*�� SET ANSI ����Ϊ OFF ʱ���������ַ����������ַ��ıȽϣ�ֱ���϶��ַ�����β��
set lock off
******20151114****************
*(Ĭ�� off) ʹ����������Թ���ʽ���ʱ�������شӱ��л�����µ���Ϣ����ʹ�� SET LOCK OFF 
set exclusive on
**********20151114****************
*Ҫ�Զ�ռʹ�÷�ʽ�򿪱�����������¼���ļ����Զ�ռʹ�÷�ʽ�򿪵ı�ȷ�������û����ܸ����ļ�������ĳЩ������Ǳ��Զ�ռʹ�÷�ʽ�򿪣�������ִ�С���Щ������ INSERT�� INSERT BLANK�� MODIFY STRUCTURE�� PACK�� REINDEX �� ZAP
set multilocks on
set deleted off
*****��ʾɾ����ʶ************
set optimize on
*(Ĭ��) ���� Rushmore �Ż�
set refresh to 1,1
**********20151114****************
*ÿ��ָ����������ˢ�±����ڴ滺����������С��ֵ
set seconds on
set century on
set currency left
set currency to 'nt$'
set sysformats off
**********20151203*************************************
*off���������������أ�on����������ʱ�䡢����������Ч
set hours to 24
set date to YMD
set date to long
set stri to 0
set decimals to 2
set fdow to 1
set fweek to 1
set mark to '.'
set separator to ','
set point to '.'
PUBLIC I,ND,LY,YF,NY,RK,�汾,�˿�,XM,bh,MD,WHNY,WHNY1,WHNY2,WHTJ,DBFF,DBFF1,bf,xp
I=1
*****������ȫ�ֱ���**********
RK = 8
****�û�Ȩ�޷ּ���ֵȫ�ֱ���*RK******
�˿�=''
****���̵��ñ�ʶȫ�ֱ���*�˿�******
XM=''
bh=''
MD=0
WHNY=''
WHTJ=''
*********ϵͳ��������*********
 _SCREEN.WINDOWSTATE = 2 
 **********��󻯴���************
set default to sys(5)+curdir()
SET PATH TO SYS(5) + CURDIR() ; DATA\
close all
USE ��Ա״��
*****************��ʼ����װĿ¼�����ݿ���Ŀ¼*****************
FileName=DBF()
FPath=justpath(FileName)
*************·��********** 
bf=justdrive(FPath)
*************��׼�������̷�***eg:D:*******  
xp=ALLTRIM(FPath)
**********��׼ϵͳ��װĿ¼**��Ƭ�����Ŀ¼***eg:D:\HRmis**20150928********* 
use ccrq
go 8
I=val(ͼƬ)
do form fm
=inkey(8,"M")
=sys(2002,1)
USE dmk
_screen.caption='������Դ������Ϣϵͳ��Ver 5.17.2063 Build 171220��'
*******************�������汾��***20151110******�ֹ��趨**************************
�汾=ALLTRIM(�汾����)
*******************ϵͳ������ʼ����ֵ������ϵͳ�����û�������汾���Ͷ����нṹ��������********20151110********************************
use ��
GO top
LOCATE FOR ALLTRIM(user) == 'Admin'
IF !FOUND()
APPEND BLANK
GO bott
REPLACE user WITH 'Admin',password WITH 'root',userjb WITH 9 &&�����趨�����û����������������ʱ���ݿ�ͬ�����ۻ����¡�
ELSE
REPLACE password WITH 'root',userjb WITH 9 &&�����趨�����û����������������ʱ���ݿ�ͬ�����ۻ����¡�
ENDIF
close all
_SCREEN.MAXBUTTON = .T.
_SCREEN.MOVABLE = .T.
_SCREEN.CONTROLBOX = .T.
_SCREEN.CLOSABLE = .F.

