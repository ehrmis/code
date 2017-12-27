***************************
*LDHTGL.PRG(Build20080415)
***************************
close all
set date to YMD
set date to long
CJBH = ''
DH = ''
DH1 = ''
DH1A = ''
DYCJ = ''
DYBM = ''
do FJDY
do DHK1
do DHK1A
if DH1A=2 and DH1=2
   do htdc
endif
close all
if DH1=2 and DH1A=1
  nn=MESSAGEBOX("若打印不成功，请先预览后打印！...",0+64,"信息") 
  set device to printer
  set printer on  
else
  set print font "宋体",8
*********************▲****************
  set device to printer
  set printer to 劳动合同.TXT
  set printer on 
endif
use ldht excl
zap
append from ryzk
RS = reccount()
do case
case DH=2
  coun for cjdm='&dycj' to rs
  set filter to CJDM='&DYCJ'
case DH=3
  coun for bmbh='&dybm' to rs
  set filter to BMBH='&DYBM'
endcase
PG = 1
N = 1
go top
do while  not eof()
  set print font "宋体",16
  @ 3 , 0 say space(18)+'昆钢集团劳动合同订立情况调查表'
  set print font "宋体",8
  @ prow()+3 , 3 say '填报单位:'+alltrim(MC111)+space(33)+'第'+str(PG;
,2)+'页'+space(28)+'日期:'+str(year(date()),4)+'年'+str(month(date()),2)+'月'+str(day(date());
,2)+'日'
  T1 = '┏━━┯━━━━┯━┯━━━━━━━┯━━┯━━━━━━┯━━━━━━━┯━━━━┯━━━━━━━━━┯━━━┓'
  T2 = '┃    │        │别│              │年限│            │   合同时间   │        │ 期满超欠履行时间 │      ┃'
  T3 = '┠──┼────┼─┼───────┼──┼──────┼───────┼────┼─────────┼───┨'
  T4 = '┗━━┷━━━━┷━┷━━━━━━━┷━━┷━━━━━━┷━━━━━━━┷━━━━┷━━━━━━━━━┷━━━┛'
  @ prow()+1 , 0 say T1
  @ prow()+0.8 , 0 say '┃序号│ 姓  名 │性│   出生日期   │工作│ 岗位(工种) │   签订劳动   │合同期限│'+str(year(date()),4)+'年'+str(month(date());
,2)+'月末至合同│备  注┃'
  @ prow()+0.8 , 0 say T2
  @ prow()+0.8 , 0 say T3
  if PG>1
    N = (PG-1)*30+1
    skip
  endif
  do while  not eof()
    @ prow()+0.8 , 0 say '┃'+str(N,4)+'│'+RYXM+'│'+XB1+'│'
    do case
    case len(allt(dtoc(csrq)))=12
         CS=allt(dtoc(csrq))+space(2)
    case len(allt(dtoc(csrq)))=13
         CS=allt(dtoc(csrq))+space(1)
    case len(allt(dtoc(csrq)))=14
         CS=allt(dtoc(csrq))
    endcase             
    @ prow() , pcol() say CS+'│'+str(gl,2)+'年│'
    if  not empty(GZ1)
      @ prow() , pcol() say left(GZ1,12)+'│'
    else
      @ prow() , pcol() say left(GW,12)+'│'
    endif
IF year(XHTrQ)>0
***********续签订合同人员**************
       合同年限 = val(xQHTQ)+(year(xHTRQ)-year(date()))+round((month(xHTRQ)-month(date()))/12,2)
    do case
    case len(allt(dtoc(xhtrq)))=12
         CS=allt(dtoc(xhtrq))+space(2)
    case len(allt(dtoc(xhtrq)))=13
         CS=allt(dtoc(xhtrq))+space(1)
    case len(allt(dtoc(xhtrq)))=14
         CS=allt(dtoc(xhtrq))
    endcase    
      @ prow() , pcol() say CS+'│'+XQHTQ1+'│'            
     if !'无固定期'$XQHTQ1 and !empty(xqhtq1)
      do case
      case val(left(str(合同年限,5,2),2))>0 and val(right(str(合同年限,5;
,2),2))/100*12>0
        replace HTNX with left(str(合同年限,5,2),2)+'年零'+str(val(right(str(合同年限;
,5,2),2))/100*12,2)+'个月'
        @ prow() , pcol() say space(3)+HTNX+space(3)
      case val(left(str(合同年限,5,2),2))>0 and val(right(str(合同年限,5;
,2),2))/100*12=0
        replace HTNX with left(str(合同年限,5,2),2)+'年'
        @ prow() , pcol() say space(3)+HTNX+space(3)
      case val(left(str(合同年限,5,2),2))=0 and val(right(str(合同年限,5;
,2),2))/100*12>0
        replace HTNX with str(val(right(str(合同年限,5,2),2))/100*12,2)+'个月'
        @ prow() , pcol() say space(3)+HTNX+space(3)
         case val(left(str(合同年限,5,2),2))=0 and val(right(str(合同年限,5;
,2),2))/100*12=0
        replace HTNX with '0个月'
         @ prow() , pcol() say space(3)+HTNX+space(3)
      endcase      
else
@ prow() , pcol() say space(18)
endif
@ prow() , pcol() say '│      ┃'
ELSE    
***********非续签订合同人员**************
       if !'无固定期'$sQHTQ1 and !empty(sqhtq1)
       合同年限 = val(SQHTQ)+(year(SHTRQ)-year(date()))+round((month(SHTRQ)-month(date()))/12,2)
        do case
    case len(allt(dtoc(shtrq)))=12
         CS=allt(dtoc(shtrq))+space(2)
    case len(allt(dtoc(shtrq)))=13
         CS=allt(dtoc(shtrq))+space(1)
    case len(allt(dtoc(shtrq)))=14
         CS=allt(dtoc(shtrq))
    endcase 
       @ prow() , pcol() say CS+'│'+SQHTQ1+'│'
      else
        合同年限 = val(SQHTQ)+(year(SHTRQ)-year(date()))+round((month(SHTRQ)-month(date()))/12,2)
        do case
    case len(allt(dtoc(shtrq)))=12
         CS=allt(dtoc(shtrq))+space(2)
    case len(allt(dtoc(shtrq)))=13
         CS=allt(dtoc(shtrq))+space(1)
    case len(allt(dtoc(shtrq)))=14
         CS=allt(dtoc(shtrq))
    endcase    
        @ prow() , pcol() say CS+'│'+SQHTQ1+'│'
      endif
       if !'无固定期'$sQHTQ1 and !empty(sqhtq1)
      do case
      case val(left(str(合同年限,5,2),2))>0 and val(right(str(合同年限,5;
,2),2))/100*12>0
        replace HTNX with left(str(合同年限,5,2),2)+'年零'+str(val(right(str(合同年限;
,5,2),2))/100*12,2)+'个月'
        @ prow() , pcol() say space(3)+HTNX+space(3)
      case val(left(str(合同年限,5,2),2))>0 and val(right(str(合同年限,5;
,2),2))/100*12=0
        replace HTNX with left(str(合同年限,5,2),2)+'年'
        @ prow() , pcol() say space(3)+HTNX+space(3)
      case val(left(str(合同年限,5,2),2))=0 and val(right(str(合同年限,5;
,2),2))/100*12>0
        replace HTNX with str(val(right(str(合同年限,5,2),2))/100*12,2)+'个月'
        @ prow() , pcol() say space(3)+HTNX+space(3)
        case val(left(str(合同年限,5,2),2))=0 and val(right(str(合同年限,5;
,2),2))/100*12=0
        replace HTNX with '0个月'
        @ prow() , pcol() say space(3)+HTNX+space(3)
      endcase
         @ prow() , pcol() say '│ 首签 ┃'
      else
         @ prow() , pcol() say space(18)+'│      ┃'
      endif   
ENDIF
    if mod(N,30)=0
      @ prow()+0.8 , 0 say T4
      exit
    else
      @ prow()+0.8 , 0 say T3
    endif     
    skip
    N = N+1
   enddo
  PG = PG+1
enddo
@ prow()+0.8 , 0 say T4
set filter to
close all
@ prow()+1 , 10 say '法定代表人(或委托代理人):'+ZG111+space(42)+'制表人:'+ZB111
@ prow()+1 , 0 say ' '
set device to screen
set print off
set print to
close all
use ldht
repl htdqsj with allt(htnx) for !'零'$htnx
replace SQHTQ1 with XQHTQ1 for  not empty(XQHTQ1)
index on CJDM+SQHTQ1+HTNX to ldht
copy to d:\劳动合同 type xl5 for !empty(htdqsj)
=messagebox("今年合同到期员工情况导出到“D：\劳动合同”和“D：\劳动合同调查”电子表中！！",48,"恭喜")
total on CJDM+SQHTQ1+HTNX to 合同调查
use 合同调查
copy to d:\劳动合同调查 type xl5
index on CJDM+SQHTQ1 to ldht
total on CJDM+SQHTQ1 to 合同1 
index on SQHTQ1 to ldht
total on SQHTQ1 to 合同2 
use 合同调查
copy to htdc
close all
if DH1=1 and DH1A=1
  modi file 劳动合同.TXT
endif
if DH1A=2 and DH1=1
  use 合同调查
  sum to RS1 A
  index on SQHTQ1 to 合同调查
  go top
  do case
  case DH=3
    browse field BMMC :R :H = '部门名称' , SQHTQ1 :R :H = '签订合同期限';
 , HTNX :R :H = '至今尚未履行合同时间' , A :R : 4 :H = '人数';
 title '合同调查表('+alltrim(CJMC)+':'+str(RS1,4)+'人)'
  case DH=2
    browse field CJMC :R :H = '车间名称' , SQHTQ1 :R :H = '签订合同期限';
 , HTNX :R :H = '至今尚未履行合同时间' , A :R : 4 :H = '人数';
 title '合同调查表('+alltrim(CJMC)+':'+str(RS1,4)+'人)'
  case DH=1
    browse field CJMC :R :H = '车间名称' , SQHTQ1 :R :H = '签订合同期限';
 , HTNX :R :H = '至今尚未履行合同时间' , A :R : 4 :H = '人数';
 title '合同调查表('+str(RS1,4)+'人)'
  use 合同1
    browse field CJMC :R :H = '车间名称' , SQHTQ1 :R :H = '签订合同期限';
 , A :R : 4 :H = '人数';
 title '合同调查表('+str(RS1,4)+'人)'
  use 合同2
    browse field SQHTQ1 :R :H = '签订合同期限';
 , A :R : 4 :H = '人数';
 title '合同调查表('+str(RS1,4)+'人)'
 endcase
endif
close all
set device to screen
if DH1A=2 and DH1=2
   set printer to lpt1
   do htdc   
endif
set printer off
set printer to
close all
clear
return

proc htdc
set talk off
set safety off
tt1='┏━━┯━━━━━━┯━━━━━━┯━━━━━━━━━━┯━━━┓'
tt2='┃序号│  车间名称  │签订合同期限│至今尚未履行合同时间│ 人数 ┃'
tt3='┠──┼──────┼──────┼──────────┼───┨'
tt4='┗━━┷━━━━━━┷━━━━━━┷━━━━━━━━━━┷━━━┛'
close all
use htdc
sum a to rs1
go top
M = 1
xh1 = 1
set device to printer
set printer on
set print font "宋体",12
*********************▲****************
 do while !eof()
    @ 3 , 0 say space(26)+alltrim(MC111)+'劳动合同调查表'
    @ prow()+1 , 20 say '第'+str(M;
,2)+'页'+space(18)+'日期:'+str(year(date()),4)+'年'+str(month(date());
,2)+'月'+str(day(date()),2)+'日'
      @  prow()+1 , 10 say Tt1
      @  prow()+0.8 , 10 say Tt2
      @  prow()+0.8 , 10 say Tt3
  do while !eof()
  @ prow()+0.8 , 10 say '┃'
  @ prow() , pcol() say str(xh1,4)
  @ prow() , pcol() say '│'
  @ prow() , pcol() say cjmc
  @ prow() , pcol() say '│  '
  @ prow() , pcol() say sqhtq1
  @ prow() , pcol() say '  │'
  @ prow() , pcol() say space(4)+htnx+space(4)
  @ prow() , pcol() say '│'
  @ prow() , pcol() say a picture '@Z 999999'
  @ prow() , pcol() say '┃'
  skip
  xh1=xh1+1
  if mod(xh1,25)=0
     @ prow()+1 , 10 say tt4
     m=m+1
     exit
  else
     @ prow()+0.8 , 10 say tt3
  endif
  endd
enddo
  @ prow()+0.8 , 10 say '┃'
  @ prow() , pcol() say '合计'
  @ prow() , pcol() say '│'
  @ prow() , pcol() say space(12)
  @ prow() , pcol() say '│'
  @ prow() , pcol() say space(12)
  @ prow() , pcol() say '│'
  @ prow() , pcol() say space(20)
  @ prow() , pcol() say '│'
  @ prow() , pcol() say str(rs1,6)
  @ prow() , pcol() say '┃'
@ prow()+0.8 , 10 say tt4
?''
set device to screen
set print off
set print to
close all
clear
retu
