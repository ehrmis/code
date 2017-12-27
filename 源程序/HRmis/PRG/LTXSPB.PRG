******************
*LTXSPB.prg
******************
ny=''
yn=messagebox('需要更新当前数据？！...请插入磁盘！',3+32+256,'提示')
do case
   case yn=6
do srny
ny1=right(ny,4)
use dmk
dw=allt(dw)
JB=dw+ny1
wait window '正在处理数据，请稍候... ...' nowait
use a:\tx&JB
copy to tx for lbdm<>'02'
copy to lx for lbdm='02'
use tx
inde on dwdm to txspb
tota on dwdm to txspb
use lx
inde on dwdm to lxspb
tota on dwdm to lxspb
do form txspb
   case yn=7
do form txspb
   case yn=2
   retu
endcase