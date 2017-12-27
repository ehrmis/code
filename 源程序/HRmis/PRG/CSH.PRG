*CSH.prg
set safety off
wait wind "正在拷贝原始数据... ..." nowa
use dmk
if !empty(路径)
if right(allt(路径),1)<>'\'
   sslj=allt(路径)+'\'
else
   sslj=allt(路径)
endif   
!copy &sslj.ryzk.dbf
!copy &sslj.kq*.dbf
!copy &sslj.yl*.dbf
else
messagebox('请先进行主题代码维护','提示')
endif