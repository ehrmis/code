IF I=1
YN = MESSAGEBOX('需要备份'+WHNY+'？...请选择文件备份目录',4+32,'提示')
IF YN = 6
 on error do copyerr 
 old_dirs='backup\'
Declare integer Qdir_m In Qdir.dll string @ck,;
integer yy,;
string cc
*注册Qdir_m()，@ck=返回的文件夹路径，yy=文件夹类别,cc=提示
*==========================================================
ck=spac(64)      &&存放文件夹路径
yy=1
cc='选择'+WHNY+'备份目录（“取消”则默认备份到BACKUP子目录中）'
xx=Qdir_m(@ck,yy,cc)   
*调用Qdir_m()，选择的结果在ck中返回
dirs=allt(subs(ck,2,64))
if len(dirs)=0
   dirs=old_dirs
ENDIF
if right(dirs,1)<>'\'
   dirs=dirs+'\'
endif
    use dbf()
    copy to &dirs.&DBFF1.&ND
    if 人可='失败'
       =MESSAGEBOX("数据备份失败！",48,"恭喜")
    else
      =MESSAGEBOX("数据备份成功！",48,"恭喜")
    endif
close all
on error do CCCL with error() , message() , message(1) , sys(16) , lineno()
ENDIF
ELSE
 on error do copyerr 
 old_dirs='backup\'
Declare integer Qdir_m In Qdir.dll string @ck,;
integer yy,;
string cc
*注册Qdir_m()，@ck=返回的文件夹路径，yy=文件夹类别,cc=提示
*==========================================================
ck=spac(64)      &&存放文件夹路径
yy=1
cc='选择'+WHNY+'备份目录（“取消”则默认备份到BACKUP子目录中）'
xx=Qdir_m(@ck,yy,cc)   
*调用Qdir_m()，选择的结果在ck中返回
dirs=allt(subs(ck,2,64))
if len(dirs)=0
   dirs=old_dirs
ENDIF
if right(dirs,1)<>'\'
   dirs=dirs+'\'
endif
    use dbf()
    copy to &dirs.&DBFF1.&ND
    if 人可='失败'
       =MESSAGEBOX("数据备份失败！",48,"恭喜")
    else
      =MESSAGEBOX("数据备份成功！",48,"恭喜")
    endif
close all
on error do CCCL with error() , message() , message(1) , sys(16) , lineno()
ENDIF