***************************
* .\LY_HFSJ.PRG
***************************
set safety off
 CLOSE ALL
 YN0 = MESSAGEBOX('需要恢复自定义备份数据吗？！...请插入磁盘或选择自定义文件夹！',291,'提示')
 DO CASE 
 CASE YN0 = 6
 old_dirs='d:'
Declare integer Qdir_m In Qdir.dll string @ck,;
integer yy,;
string cc
*注册Qdir_m()，@ck=返回的文件夹路径，yy=文件夹类别,cc=提示
*==========================================================
ck=spac(64)      &&存放文件夹路径
yy=1
cc='请选择好要恢复文件的备份目录,单击确定按钮。'
xx=Qdir_m(@ck,yy,cc)   
*调用Qdir_m()，选择的结果在ck中返回
dirs=allt(subs(ck,2,64))
if len(dirs)=0
   dirs=old_dirs
endif
if right(dirs,1)<>'\'
   dirs=dirs+'\'
endif
if file(dirs+'YL'+ly+'.dbf')
yn = MESSAGEBOX('恢复'+alltrim(LY)+'年养老保险文件('+dirs+'YL'+alltrim(LY)+'.DBF)？',4+32,"提示")
IF yn = 6
  wait window '正在恢复养老保险库... ...' nowait
    use &dirs.YL&ly
    copy to YL&LY
=MESSAGEBOX("数据恢复成功!",48,"恭喜")
endif
endif
if file(dirs+'zr'+ly+'.dbf') 
yn = MESSAGEBOX('恢复'+alltrim(LY)+'年工资总额文件('+dirs+'ZR'+alltrim(LY)+'.DBF)？',4+32,"提示")
    IF yn = 6
   wait window '正在恢复工资总额库... ...' nowait
    if file(dirs+'zr'+ly+'.dbf')
       use &dirs.ZR&ly
       copy to ZR&LY
    endif
    endif
endif
if file(dirs+'zr1'+ly+'.dbf')
yn = MESSAGEBOX('恢复'+str(val(alltrim(LY)+1),4)+'年缴费基数文件('+dirs+'ZR'+alltrim(LY)+'.DBF)？',4+32,"提示")
IF yn = 6
    if file(dirs+'zr1'+ly+'.dbf')
       use &dirs.ZR1&ly
       copy to ZR1&LY
    endif
    =MESSAGEBOX("数据恢复成功!",48,"恭喜")
endif
endif
 CASE YN0 = 7
   if file('backup\yl'+ly+'.dbf')
    yn = MESSAGEBOX('恢复备份库中'+alltrim(LY)+'年社会保险文件(backup\YL'+alltrim(LY)+'.DBF)？',4+32,"提示")
IF yn = 6
  wait window '正在恢复养老保险库... ...' nowa
  use backup\yl&ly
  copy to sy&LY
   =MESSAGEBOX("数据恢复成功!",48,"恭喜")
  wait window '正在恢复失业保险库... ...' nowa
  use backup\sy&ly
  copy to sy&LY
   =MESSAGEBOX("数据恢复成功!",48,"恭喜")
endif 
   endif 
   if file('backup\zr'+ly+'.dbf')
    yn = MESSAGEBOX('恢复备份库中'+alltrim(LY)+'年工资总额文件(backup\ZR'+alltrim(LY)+'.DBF)？',4+32,"提示")
IF yn = 6
  wait window '正在恢复工资总额库... ...' nowa
  use backup\zr&ly
  copy to zr&LY
   =MESSAGEBOX("数据恢复成功!",48,"恭喜")
endif
   endif
   xn=str(val(LY)+1,4)
   if file('backup\zr1'+ly+'.dbf')
    yn = MESSAGEBOX('恢复备份库中'+alltrim(xn)+'年缴费基数文件(backup\ZR1'+alltrim(LY)+'.DBF)？',4+32,"提示")
IF yn = 6
  wait window '正在恢复缴费基数库... ...' nowa
  use backup\zr1&ly
  copy to zr1&LY
   =MESSAGEBOX("数据恢复成功!",48,"恭喜")
endif 
   endif    
 CASE YN0 = 2
    RETURN 
 ENDCASE