***************************
* .\YB.PRG
***************************
close all
set safety off
do 岗位管理
use rk
zap 
append from \tt1\ry02
replace A with 0 for (val(substr(SFZ,7,2))=0 or val(substr(SFZ,9,2))=0 or;
 val(substr(SFZ,9,2))>12 or val(substr(SFZ,11,2))=0 or val(substr(SFZ,11;
,2))>31) and len(alltrim(SFZ))=15
replace A with 0 for val(substr(SFZ,9,2))=2 and val(substr(SFZ,11,2))>28;
 and len(alltrim(SFZ))=15
replace A with 0 for len(alltrim(SFZ))<15
replace A with 0 for empty(left(SFZ,1))
replace ZZ with '空' for empty(ZZ)
replace JG with '空' for empty(JG)
count for A=0 or len(alltrim(SFZ))>15 to SF
if SF>0
  browse for A=0 or len(alltrim(SFZ))>15 field RYXM , CSRQ , SFZ title '请认真核对或修正下列身份证号'
endif 
replace SFZ with '111111111111111' for empty(SFZ)
replace A with 1 all
index on MZ to rk
total on MZ to mz
index on WHCD to rk
total on WHCD to wh
index on ZW to rk
total on ZW to zw
index on ZCDJ to rk
total on ZCDJ to zc
use mz
browse field MZ , MZ1 , A
use rk
go top
locate for empty(MZ) or empty(MZ1)
if  not eof()
  browse field RYXM , MZ , MZ1
endif 
use wh
browse field WHCD , WHCD1 , A
use rk
go top
locate for empty(WHCD) or empty(WHCD1)
if  not eof()
  browse field RYXM , WHCD , WHCD1
endif 
use zw
browse field ZW , ZW1 , A
use rk
go top
locate for empty(ZW) or empty(ZW1)
if  not eof()
  browse field RYXM , ZW , ZW1
endif 
use zc
browse field ZCDJ , ZC , A
use rk
go top
locate for empty(ZCDJ) or empty(ZC)
if  not eof()
  browse field RYXM , ZCDJ , ZC
endif 
clear 
close all
use rk
replace BAC013 with ZC for  not '技师'$ZC and  not empty(ZC)
close all
use 医保代码
copy to ybwh for alltrim(代码)='WHCD'
select 2
use ybwh
go top
do while  not eof()
  select 1
  use rk
  scan for alltrim(WHCD1)=alltrim(b.内容)
    replace WHCD with b.编号
  endscan 
  select 2
  skip 
enddo 
close all
use rk
replace AAC001 with '00000'+alltrim(RYBH) , AAC002 with left(SFZ,18) , AAC003;
 with RYXM , AAC006 with ctod(substr(CSRQ,5,2)+'/'+substr(CSRQ,7,2)+'/'+left(CSRQ;
,4)) , AAC007 with ctod(substr(CJGZRQ,5,2)+'/'+substr(CJGZRQ,7,2)+'/'+left(CJGZRQ;
,4)) , AAB001 with DWDM , AAC004 with XB , AAC005 with MZ , AAE006 with ZZ;
 all
replace WHCD with '06' for '技'$WHCD1
replace AAC004 with '2' for XB='0'
replace BAC012 with GD for  not '岗'$GD and TSRY=0 and upper(CJF)<>'Y'
replace BAC012 with GW for TSRY=0 and upper(CJF)='Y'
replace AAC020 with '08' for ZW='06'
replace AAC020 with '07' for ZW='05'
replace AAC020 with '06' for ZW='04'
replace AAC020 with '05' for ZW='03'
replace AAC020 with '09' for '科员'$GD
replace AAC020 with '10' for GD='办事员'
replace AAC014 with '2' for ZCDJ='02'
replace AAC014 with '4' for ZCDJ='03'
replace AAC014 with '5' for ZCDJ='04'
replace AAC014 with '6' for ZCDJ='05'
replace AAC012 with '01' for val(RYFL)>2 and val(RYFL)<8
replace AAC012 with '02' for empty(AAC012)
replace AAC008 with '1' for empty(LBDM)
replace AAC008 with '2' for LBDM='02'
replace AAC008 with '5' for LBDM='01'
replace BAC001 with '01' for '已'$HF
replace BAC001 with '01' for  not '已'$HF
close all
select 2
use zr12001
index on RYBH to zr12001
select 1
use rk
index on RYBH to rk
update on RYBH from 2 replace BAC009 with b.PJ
browse title '请认真审核'
close all
use yb
zap 
append from rk
replace AAE001 with 2002 , CSD with left(JG,6) , AAC009 with '城镇户口' ;
, AAC013 with '9' , XX with '9' all
replace AAE007 with '651206' , AAC010 with '云南省禄丰县仁兴镇' all
browse title '请重新认真审核'
close all
clear 
quit 

