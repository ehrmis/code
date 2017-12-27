**********************************************************************
*SETTING.PRG &&(通用系统运行环境设置 Build 20151203)
**********************************************************************
set safety off
set talk off
set palette off
set status bar on
set notify on
*系统信息显示在 Visual FoxPro 主窗口底部的图形状态栏，而不是基于字符的状态栏中
SET CLOCK status
set escape on
set keycomp to window
* to DOS 对话框中的默认按钮具有焦点，它的外观不变化，无提示执行；反之，就用 set keycomp to windows，具有电子表的单元格复制粘贴功能。****20151124*********** 
set carry on
*在所有的工作区中创建新记录时，将当前记录的所有字段的内容复制到由 INSERT、 APPEND 和 BROWSE 命令创建新记录中 
set confirm on
set exact on
*在两个表达式的较短的一个的右边加上空格或零(0)字节，以使它与较长表达式的长度相匹配。但是，在比较中的任何表达式尾部的空格或零字节都被忽略
set near on
set ansi off
*当 SET ANSI 设置为 OFF 时，对两个字符串进行逐字符的比较，直到较短字符串的尾部
set lock off
******20151114****************
*(默认 off) 使后述命令可以共享方式访问表，如果不必从表中获得最新的信息，可使用 SET LOCK OFF 
set exclusive on
**********20151114****************
*要以独占使用方式打开表，不必锁定记录或文件。以独占使用方式打开的表，确保其他用户不能更改文件。对于某些命令，除非表以独占使用方式打开，否则不能执行。这些命令是 INSERT、 INSERT BLANK、 MODIFY STRUCTURE、 PACK、 REINDEX 和 ZAP
set multilocks on
set deleted off
*****显示删除标识************
set optimize on
*(默认) 允许 Rushmore 优化
set refresh to 1,1
**********20151114****************
*每隔指定的秒数都刷新本地内存缓冲区，允许小数值
set seconds on
set century on
set currency left
set currency to 'nt$'
set sysformats off
**********20151203*************************************
*off与下面语句配置相关，on则下面三句时间、日期设置无效
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
PUBLIC I,ND,LY,YF,NY,RK,版本,人可,XM,bh,MD,WHNY,WHNY1,WHNY2,WHTJ,DBFF,DBFF1,bf,xp
I=1
*****计数器全局变量**********
RK = 8
****用户权限分级赋值全局变量*RK******
人可=''
****过程调用标识全局变量*人可******
XM=''
bh=''
MD=0
WHNY=''
WHTJ=''
*********系统变量设置*********
 _SCREEN.WINDOWSTATE = 2 
 **********最大化窗口************
set default to sys(5)+curdir()
SET PATH TO SYS(5) + CURDIR() ; DATA\
close all
USE 人员状况
*****************初始化安装目录、数据库子目录*****************
FileName=DBF()
FPath=justpath(FileName)
*************路径********** 
bf=justdrive(FPath)
*************标准驱动器盘符***eg:D:*******  
xp=ALLTRIM(FPath)
**********标准系统安装目录**像片存放子目录***eg:D:\HRmis**20150928********* 
use ccrq
go 8
I=val(图片)
do form fm
=inkey(8,"M")
=sys(2002,1)
USE dmk
_screen.caption='人力资源管理信息系统（Ver 5.17.2063 Build 171220）'
*******************保存编译版本号***20151110******手工设定**************************
版本=ALLTRIM(版本类型)
*******************系统启动初始化赋值：根据系统设置用户类别分配版本类型而进行结构化错误处理********20151110********************************
use ＇
GO top
LOCATE FOR ALLTRIM(user) == 'Admin'
IF !FOUND()
APPEND BLANK
GO bott
REPLACE user WITH 'Admin',password WITH 'root',userjb WITH 9 &&刚性设定超级用户，用于主程序更新时数据库同步“累积更新”
ELSE
REPLACE password WITH 'root',userjb WITH 9 &&刚性设定超级用户，用于主程序更新时数据库同步“累积更新”
ENDIF
close all
_SCREEN.MAXBUTTON = .T.
_SCREEN.MOVABLE = .T.
_SCREEN.CONTROLBOX = .T.
_SCREEN.CLOSABLE = .F.

