***************************
* .\dhkfj.PRG(2016.10.2)
***************************
close all
define window MYWIN1 FROM INT((SROWS()-15)/2)  , INT((SCOLS()-65)/2)  TO INT((SROWS()+15)/2) , INT((SCOLS()+65)/2) shadow title '筛选单位[慎用鼠标操作]:';
 footer '[慎用鼠标操作]'
activate window MYWIN1
@ 1 , 8 SAY '请选择范围[用“←↑↓→”操作]:' font '宋体',13
store 1 to DH
@ 5 , 5 get DH size 2 , 12 , 2 function '*h 全  部;分单位;分部门' font '宋体',12
read
do case
   case dh=1
close all
clear window MYWIN1
retu
   case dh=2
  use cjk
  go top
 define popup MYPOP1 from 1 ,0  to 26 , 75 prompt field CJmc title '单位名称';
 shadow
  on selection popup MYPOP1 DEACT POPUP MYPOP1
  activate popup MYPOP1
  @ 0 ,10 say CJmc
  DYCJ = allt(cjDM)
close all
clear window MYpop1
clear window MYWIN1
retu
    case dh=3
  select 1
  use cjk
  go top
* set colo to 7+/2 
 define popup MYPOP1 from 1 , 0 to 26 , 75 prompt field CJmc title '部门(班组)';
 shadow
  on selection popup MYPOP1 DEACT POPUP MYPOP1
  activate popup MYPOP1
  @ 0 , 10 say CJmc
   dm1=cjdm
  select 2
  use BMDM
  set filter to left(bmbh,2)='&dm1'
  go top
  define popup MYPOP1 from 1 , 0 to 26 , 75 prompt field BMMC title '部门(班组)' shadow
  on selection popup MYPOP1 DEACT POPUP MYPOP1
  activate popup MYPOP1
  @ 0 , 20 say BMMC
  DYBM = allt(b.BMbh)
endcase
set filter to
clear window MYpop1
clear window MYWIN1
close all
retu 