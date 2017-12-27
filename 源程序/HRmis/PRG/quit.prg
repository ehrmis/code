yn = MESSAGEBOX("  您真的要退出系统吗？",4+32,"提示")
IF yn = 6
close all
clear all							
CLEAR  MEMORY &&  清除所有内存变量
CLEAR EVENTS  &&  结束事件的读取
quit &&强行退出系统					
ENDIF


