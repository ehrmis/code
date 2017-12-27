***************************
* .\YYJLB.PRG(Build20060322)
***************************
clear
_SCREEN.PICTURE = ''
PD = ''
YX = '1'
RQ = ''
XM = ''
DHYY = 1
DZZ = space(1)
连续工龄 = space(6)
工种年限 = space(5)
工龄系数 = space(4)
store space(2) to STJFN , STJFN1
store 0 to 支付金额1 , 支付金额2 , 支付金额3 , 缴费年限 , 全部缴费年 , GRBX;
 , GRBX1 , TXYS , YLJE , ZFJE , JFN , 视同缴费年
store 1 to YX1
use ryzk
go top
do while .t.
  do srxm.spr
  locate for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
  if eof()
    use ryzk1
    go top
    locate for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
    if eof()
       do dhkjg.spr
    else
       exit
    endi
    M = readkey()
    if M=12 or M=268
       return
    endif
  else
    exit
  endif
enddo
do srny.spr
人可 = '支付'
use ryzk
locate for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
if eof()
   use ryzk1
   locate for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
   whtj='ryzk1'
else
   whtj='ryzk'
endif
GZRQ = alltrim(CJGZRQ)
bh=allt(grbh)
***********
do grzhcl
***********
close all
select 2
use ryzk
locate for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
if eof()
   use ryzk1
   locate for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH) 
endif
GZRQ = alltrim(CJGZRQ)
bh=allt(grbh)
select 1
use ylgrzh
do while val(GZRQ)<19951001
clear
_SCREEN.PICTURE = ''
 @ 22 , 43 say '提示:参见个人帐户登记卡' font "宋体",12
      @ 11 , 30 say '请输入'+alltrim(RYXM)+'视同缴费年:' font "宋体",12 get STJFN font "宋体",12
      read
      @ 11 , 65 say '年' font "宋体",12 get STJFN1 font "宋体",12
      read
      JFN = round(val(STJFN)+val(STJFN1)/12,2)
      @ 11 , 71 say '个月=' font "宋体",12
      PD = 'Y'
      @ 11 , col() say str(JFN,5,2)+'年' font "宋体",12
      @ 13 , 46 say '正确吗?(Y/N)' font "宋体",12 get PD font "宋体",12
      read
  replace 视同缴费年 with JFN for allt(XM)=allt(RYXM) or allt(XM)=allt(RYBH)
  if JFN>0 and upper(PD)='Y'
    exit
  endif
enddo
GO BOTT
clear
do dhkyy
do case
case DHYY=1
  支付金额1 = 历年个人缴+历年利息2+个人缴费+当年利息
case DHYY=2
  支付金额2 = 总累计
case DHYY=3
  支付金额3 = 历年个人缴+历年利息2+个人缴费+当年利息
case DHYY=4
  do YYJLB2
  retu
case DHYY=5
  do YYJLB1
  retu
endcase
TT1 = '  单位名称:'+'&MC111'+'             单位编号:        ;
   金额单位:    元(保留角)'
TT2 = '┏━━━━━┯━━━━━┯━━━━━━┯━━━━━━━━━━━┯━━━━┯━━━━━━┓'
TT3 = '┠─────┼─────┼──────┼───────┬───┴──┬─┴──────┨'
TT4 = '┃终止养老保│          │ 当年月记帐 │              │ 当年月记帐;
 │                ┃'
TT5 = '┠─────┴─────┼──────┴───┬───┴──────┼────────┨'
TT6 = '┃     终止养老保险     │      累计本息      │    个人缴费本息 ;
   │  一次性支付    ┃'
TT7 = '┃       关系原因       │      储存总额      │      储存金额   ;
   │    金  额      ┃'
TT8 = '┠───────────┼──────────┼──────────┼────────┨'
TT9 = '┃     在职职工出国     │      _________     │                 ;
   │                ┃'
TT10 = '┠───────────┼──────────┼──────────┼────────┨'
TT11 = '┠───────────┼──────────┼──────────┼────────┨'
TT12 = '┃     不符合退休条件   │                    │     __________ ;
    │                ┃'
TT13 = '┠───────────┼──────────┴──────────┴────────┨'
TT14 = '┃一次性支付金额(大写)  │                                      ;
                      ┃'
缴费年限 = (val(left(RQ,4))-1996)+round((val(substr(RQ,6,2))+3)/12,2)
全部缴费年 = round(缴费年限+视同缴费年,2)
wait window '正  在  打  印,请  稍  候... ...'  nowait
set device to printer
set printer on
set print font "宋体",16 
  @ 3 , 0 say '个人账户储存额一次性支付审批表(云养险六表)'
  set print font "宋体",12
  @ prow()+1 , 0 say TT1
  @ prow()+0.8 , 0 say TT2
  @ prow()+0.8 , 0 say '┃ 姓    名 │'+space(2)+RYXM+'│ 身份证号码 │'+space(2)+b.SFZ+space(2)+'│缴费年限│'+left(str(全部缴费年;
,5,2),2)+'年零'+str(val(right(str(全部缴费年,5,2),2))/100*12,2)+'个月┃'
  @ prow()+0.8 , 0 say TT3
  @ prow()+0.8 , 0 say TT4
  @ prow()+0.8 , 0 say '┃险关系日期│'+left(RQ,4)+'年'+substr(RQ,6,2)+'月│额个人％金额│'+space(4)+str(月个人缴费;
,6,2)+space(4)+'│额单位％金额│'+space(5)+str(月单位缴费,6,2)+space(5)+'┃'
  @ prow()+0.8 , 0 say TT5
  @ prow()+0.8 , 0 say TT6
  @ prow()+0.8 , 0 say TT7
  @ prow()+0.8 , 0 say TT8
  @ prow()+0.8 , 0 say TT9
  @ prow()+0.8 , 0 say '┃       出境定居       │                    │'
  @ prow() , pcol() say 支付金额3 picture '@z 99999999999999999.99'
  @ prow() , pcol() say '│'
  @ prow() , pcol() say 支付金额3 picture '@z 9999999999999.99'
  @ prow() , pcol() say '┃'
  @ prow()+0.8 , 0 say TT10
  @ prow()+0.8 , 0 say '┃       在职死亡       │     ──────   │'
  @ prow() , pcol() say 支付金额1 picture '@z 99999999999999999.99'
  @ prow() , pcol() say '│'
  @ prow() , pcol() say 支付金额1 picture '@z 9999999999999.99'
  @ prow() , pcol() say '┃'
  @ prow()+0.8 , 0 say TT11
  @ prow()+0.8 , 0 say TT12
  @ prow()+0.8 , 0 say '┃      达到退休年龄    │'
  @ prow() , pcol() say 支付金额2 picture '@z 99999999999999999.99'
  @ prow() , pcol() say '│                    │'
  @ prow() , pcol() say 支付金额2 picture '@z 9999999999999.99'
  @ prow() , pcol() say '┃'
  @ prow()+0.8 , 0 say TT13
  @ prow()+0.8 , 0 say TT14
  @ prow()+0.8 , 0 say '┠───────────┴────────┬─────────────────────┨'
  @ prow()+0.8 , 0 say '┃ 单位意见:                              │社会保险局审批;
 (盖章)                     ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │同意一次性支付____________;
 元             ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                经办人:             ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃  经办人:          审核人:              │      ;
              年   月   日          ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┠────────────────────┼─────────────────────┨'
  @ prow()+0.8 , 0 say '┃ 本人或家属申请:                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃   因:                                  │     收款人;
 (盖章)                        ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                   申请人:              │      ;
              年   月   日          ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┗━━━━━━━━━━━━━━━━━━━━┷━━━━━━━━━━━━━━━━━━━━━┛'
  @ prow()+1 , 0 say ' 注:1.单位、本人或家属办理申请时必须带有关证明，如:身份证、养老保险手册、出国出境定'
  @ prow()+1 , 0 say '    居证明、死亡证明等。'
  @ prow()+1 , 0 say '    2.此表一式二份,一份退业务经办人留存,一份作财务支付凭单。'  
close all
人可 = '00'
delete file yyjlb.dbf
delete file yyjlb.idx
set device to screen
?''
set print off
set print to
clear
 _SCREEN.WINDOWSTATE = 2
close all
use ccrq
go 8
repl 图片 with str(val(图片)+1,2)
if allt(图片)='10'
   repl 图片 with '0'
endi   
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
retu


proc YYJLB1
GO BOTT
GRBX=个人本息
do srgrbx.spx
  TXND = left(RQ,4)
  TXYF = str(val(substr(RQ,6,2)),2)
  TXYS = (val(left(RQ1,4))-val(left(RQ,4)))*12+(val(substr(RQ1,6,2))-val(substr(RQ;
,6,2)))
  YLJE = round(GRBX/120*TXYS,2)
  ZFJE = GRBX-YLJE
TT1 = '  单位名称:'+'&MC111'+'             单位编号:                ;
   金额单位:    元(保留角)'
TT2 = '┏━━━━━┯━━━━━┯━━━━━━┯━━━━━━━━━━━┯━━━━┯━━━━━━┓'
TT3 = '┠─────┼─────┼──────┼───────┬───┴──┬─┴──────┨'
TT4 = '┃终止养老保│(转移日期)│  累计缴费  │              │本年个人核定│;
                ┃'
TT5 = '┠─────┴─────┼──────┴───┬───┴──────┼────────┨'
TT6 = '┃ 终止(转移)养老保险   │      累计本息      │    个人缴费本息 ;
   │  一次性支付    ┃'
TT7 = '┃       关系原因       │      储存总额      │      储存金额   ;
   │    金  额      ┃'
TT8 = '┠───────────┼──────────┼──────────┼────────┨'
TT9 = '┃在职职工因调往外单位转│                    │                 ;
   │                ┃'
TT10 = '┠───────────┼──────────┼──────────┼────────┨'
TT11 = '┠───────────┼──────────┼──────────┼────────┨'
TT13 = '┠───────────┼──────────┴──────────┴────────┨'
TT14 = '┃一次性支付金额(大写)  │                                      ;
                      ┃'
缴费年限 = (val(left(RQ,4))-1996)+round((val(substr(RQ,6,2))+3)/12,2)
wait window '正  在  打  印,请  稍  候... ...'  nowait
set device to printer
set printer on
set print font "宋体",16 
  @ 3 , 0 say '个人账户储存额一次性支付审批表(云养险六表)'
  set print font "宋体",12
  @ prow()+1, 10 say '               (  退休死亡  )'
  @ prow()+1 , 0 say TT1
  @ prow()+0.8 , 0 say TT2
  @ prow()+0.8 , 0 say '┃ 姓    名 │'+space(2)+RYXM+'│ 身份证号码 │'+space(2)+b.SFZ+space(2)+'│缴费年限│'+left(str(缴费年限;
,5,2),2)+'年零'+str(val(right(str(缴费年限,5,2),2))/100*12,2)+'个月┃'
  @ prow()+0.8 , 0 say TT3
  @ prow()+0.8 , 0 say TT4
  @ prow()+0.8 , 0 say '┃险关系日期│'+left(RQ,4)+'年'+substr(RQ,6,2)+'月│;
   月  数   │'+space(4)+str(round(缴费年限*12,1),6,2)+'个月│月缴费工资额│'+space(5)+str(月缴费基数;
,6,2)+space(5)+'┃'
  @ prow()+0.8 , 0 say TT5
  @ prow()+0.8 , 0 say TT6
  @ prow()+0.8 , 0 say TT7
  @ prow()+0.8 , 0 say TT8
  @ prow()+0.8 , 0 say TT9
  @ prow()+0.8 , 0 say '┃    移养老保险关系    │'
  @ prow() , pcol() say 支付金额2 picture '@z 99999999999999999.99'
  @ prow() , pcol() say '│'
  @ prow() , pcol() say GRBX picture '@z 99999999999999999.99'
  @ prow() , pcol() say '│  ──────  ┃'
  @ prow()+0.8 , 0 say TT10
  @ prow()+0.8 , 0 say '┃       退休死亡       │     ──────   │'
  @ prow() , pcol() say GRBX picture '@z 99999999999999999.99'
  @ prow() , pcol() say '│'
  @ prow() , pcol() say ZFJE picture '@z 9999999999999.99'
  @ prow() , pcol() say '┃'
  @ prow()+0.8 , 0 say TT11
  @ prow()+0.8 , 0 say '┃  参 加 工 作 年 月   │'+left(b.CJGZRQ,4)+'年'+substr(b.CJGZRQ;
,5,2)+'月          │建个人帐户前连续工龄│'+left(str(视同缴费年,5,2),2)+'年零'+str(val(right(str(视同缴费年;
,5,2),2))/100*12,2)+'个月    ┃'
  @ prow()+0.8 , 0 say TT13
  @ prow()+0.8 , 0 say TT14
  @ prow()+0.8 , 0 say '┠───────────┴────────┬─────────────────────┨'
  @ prow()+0.8 , 0 say '┃ 单位意见:                              │社会保险局审批;
 (盖章)                     ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │同意一次性支付____________;
 元             ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                经办人:             ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃  经办人:          审核人:              │      ;
              年   月   日          ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┠────────────────────┼─────────────────────┨'
  @ prow()+0.8 , 0 say '┃ 本人或家属申请:                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃   因:                                  │     收款人;
 (盖章)                        ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┃                   申请人:              │      ;
              年   月   日          ┃'
  @ prow()+0.8 , 0 say '┃                                        │      ;
                                    ┃'
  @ prow()+0.8 , 0 say '┗━━━━━━━━━━━━━━━━━━━━┷━━━━━━━━━━━━━━━━━━━━━┛'
    @ prow()+1 , 0 say ' 注:1.单位、本人或家属办理申请时必须带有关证明，如:身份证、养老保险手册、出国出境定'
    @ prow()+1, 0 say '    居证明、死亡证明等。'
    @ prow()+1 , 0 say '    2.此表一式二份,一份退业务经办人留存,一份作财务支付凭单。'
    @ prow()+1 , 0 say '    3.本人于'+TXND+'年'+TXYF+'月份退休,至今共领取养老金:'+str(GRBX;
,6,2)+'120'+'×'+str(TXYS,3)+'个月'+str(YLJE,6,2)+'元;'
    @ prow()+1 , 0 say '    现一次性支付养老金:'+str(GRBX,6,2)+'元'+str(YLJE;
,6,2)+'元'+str(ZFJE,6,2)+'元'
close all
人可 = '00'
delete file yyjlb.dbf
delete file yyjlb.idx
set device to screen
?''
set print off
set print to
clear
 _SCREEN.WINDOWSTATE = 2
close all
use ccrq
go 8
repl 图片 with str(val(图片)+1,2)
if allt(图片)='10'
   repl 图片 with '0'
endi   
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
retu




proc YYJLB2
GO BOTT
if 缴费月数<12
  replace 单位利息 with 0 , 当年利息 with 0 , 当年合计 with 单位缴费+社平5+单位利息+个人缴费+当年利息;
 , 历年利息1 with 0 , 历年利息2 with 0 , 总累计 with 当年合计+历年单位缴+历年利息1+历年个人缴+历年利息2;
 , 当年缴纳 with 单位缴费+个人缴费 , 累计利息 with 单位利息+当年利息+历年利息1+历年利息2
endif
TT1 = '                                                                  编号:'
TT2 = '┏━━━━━━━━━━┯━━━━━━━━━━━━━━━━━┯━━━━━━━━━━━━━┓'
TT3 = '┃                    │          转  出  单  位          │      转  入  单  位      ┃'
TT4 = '┠──────────┼─────────────────┼─────────────┨'
TT5 = '┃   单  位  全  称   │                                  │                          ┃'
TT6 = '┠──────────┼──────────┬──────┼─────────────┨'
TT7 = '┃                    │                    │  职工帐号  │                          ┃'
TT8 = '┠──────────┼──────────┴──────┴─────────────┨'
TT9 = '┠──────────┼───────┬──────────────┬────────┨'
TT10 ='┠──────────┼───────┼─────┬────────┼────────┨'
TT11= '┃   累  计  本  息   │ 累计个人缴费 │ 累计缴费 │ 本年职工个人核 │  转出单位缴费  ┃'
TT12= '┃   储  存  总  额   │ 本息储存金额 │ 月    数 │  定月缴费工资  │    截止时间    ┃'
TT13= '┠──────────┼───────┼─────┼────────┼────────┨'
TT14= '┗━━━━━━━━━━┷━━━━━━━┷━━━━━┷━━━━━━━━┷━━━━━━━━┛'
缴费年限 = (val(left(RQ,4))-1996)+round((val(substr(RQ,6,2))+3)/12,2)
工作=left(b.CJGZRQ,4)+'年'+substr(b.CJGZRQ,5,2)
wait window '正  在  打  印,请  稍  候... ...'  nowait
set device to printer
set printer on
set print font "宋体",16
@ 3 , 0 say '云南省职工养老保险关系基金转移单(云养险六表)'
set print font "宋体",12
  @ prow()+1 , 0 say TT1
  @ prow()+0.8 , 0 say TT2
  @ prow()+0.8 , 0 say TT3
  @ prow()+0.8 , 0 say TT4
  @ prow()+0.8 , 0 say TT5
  @ prow()+0.8 , 0 say TT6
  @ prow()+0.8 , 0 say TT7
  @ prow()+0.8 , 0 say '┃   姓          名   │'+space(6)+RYXM+space(6)+'│ (身份证号) │'+space(4)+b.SFZ+space(4)+'┃'
  @ prow()+0.8 , 0 say TT8
  @ prow()+0.8 , 0 say '┃个人帐户储存额(大写)│                                          ￥'
  @ prow() , pcol() say 总累计 picture '@z 999999999999999.99'
  @ prow() , pcol() say '┃'
  @ prow()+0.8 , 0 say TT9
  @ prow()+0.8 , 0 say '┃    参加工作时间    │'+allt(工作)+'月    │   建立个人帐户前连续工龄   │'+left(str(视同缴费年,5,2),2)+'年零'+str(val(right(str(视同缴费年;
,5,2),2))/100*12,2)+'个月    ┃'
  @ prow()+0.8 , 0 say TT10
  @ prow()+0.8 , 0 say TT11
  @ prow()+0.8 , 0 say TT12
  @ prow()+0.8 , 0 say TT13
  @ prow()+0.8 , 0 say '┃'+str(总累计,20,2)+'│'+str(个人本息,14,2)+'│'+str(round(缴费年限*12,1),6,0)+'个月│'+space(5)+str(月缴费基数,6,2)+space(5)+'│  '+left(RQ,4)+'年'+str(val(substr(RQ,6,2)),2)+'月    ┃'
  @ prow()+0.8 , 0 say TT14
  set print font "宋体",11
  @ prow()+1 , 0 say ' 注:1、本表为草稿表,请认真审核后抄入正式表中。'
  @ prow()+1 , 0 say '    2、若转入转出单位之间属于跨省统筹范围转移时,请打印“转移说明”。' 
  @ prow()+1 , 0 say '    3、“累计本息储存总额”为转出当年的“总计”数。'
  @ prow()+1 , 0 say '    4、“累计个人本息储存金额”为“转出当年和历年个人缴费本息4个数字之和”。'
  @ prow()+1 , 0 say '  　5、转出当年缴费不足12个月时暂不结息，“累计个人本息储存金额”为“转出当年和历年个人缴费本金2个数字之和”。'   
人可 = '00'
delete file yyjlb.dbf
delete file yyjlb.idx
set device to screen
?''
set print off
set print to
clear
 _SCREEN.WINDOWSTATE = 2
close all
use ccrq
go 8
repl 图片 with str(val(图片)+1,2)
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
MESSAGEBOX("接下来按操作步骤处理数据并打印“转移说明”！",48,"提示")
use ylgrzh 
loca for val(年份)=1997
if !eof()
   个人97本息=个人缴费+当年利息+历年个人缴+历年利息2 
else
   个人97本息=0
endif
do grzhcl
use ylgrzh
GO BOTT
if 缴费月数<12
  replace 单位利息 with 0 , 当年利息 with 0 , 当年合计 with 单位缴费+社平5+单位利息+个人缴费+当年利息;
 , 历年利息1 with 0 , 历年利息2 with 0 , 总累计 with 当年合计+历年单位缴+历年利息1+历年个人缴+历年利息2;
 , 当年缴纳 with 单位缴费+个人缴费 , 累计利息 with 单位利息+当年利息+历年利息1+历年利息2
endif
set device to printer
  set printer to d:\转移说明.txt &&预览
  set printer on  
@ 3 , 0 say '               转移养老保险基金的说明 '
@ prow()+1 , 0 say '     '+allt(ryxm)+'同志于'+工作+'月参加工作，根据云劳（1998）72号文第十三条:'
@ prow()+1 , 0 say '“从1998年4月1日起，从业人员跨省统筹范围转移时，转出地(部门)的社会'
@ prow()+1 , 0 say '保险机构应向转入地(部门）的社会保险机构转移其养老保险关系、个人帐户'
@ prow()+1 , 0 say '档案和个人帐户中1998年4月1日之前的个人缴费部份累计本息加上1998年4月'
@ prow()+1 , 0 say '1日后的全部储存额。”的规定，其中：'
@ prow()+2 , 0 say '     一、1998年4月1日之前的个人缴费部份累计本息为'+str(个人97本息,6,1)+'元。'
@ prow()+2 , 0 say '     (大写:                   )。'
@ prow()+2 , 0 say '     二、1998年4月1日起记入个人帐户的全部储存额为'+str(总累计,8,1)+'元。'
@ prow()+2 , 0 say '     (大写:                   )。'
@ prow()+2 , 0 say '     以上两项合计金额为'+str(个人97本息+总累计,10,1)+'元(大写:                      )。'
@ prow()+6 , 0 say '    转出单位：（盖章）              审核单位：（盖章）'
@ prow()+5 , 0 say '    经办人：'+ZB111+'               审核人:'
@ prow()+6 , 0 say '    转出日期：    年   月   日      审核日期：     年   月   日'
@ prow()+5 , 0 say '  填写说明：(1)、1998年4月1日之前的个人缴费部份累计本息为“1997年当年和历年个人缴费本息4个数字之和”'
@ prow()+1 , 0 say '  (2)、1998年4月1日起记入个人帐户的全部储存额为1998年4月统筹后至转出当年的“总计”数。'
@ prow()+1 , 0 say '  (3)、转出当年缴费不足12个月时暂不结息，“总计”为“转出当年和历年个人缴费本金2个数字之和”。'   
@ prow()+2 , 0 say '  打印说明：请关闭此窗口后立即打印或退出后用WORD字处理软件编辑打印“D:\转移说明.txt”'
set device to screen
set printer off
set printer to
clear
_SCREEN.PICTURE = ''
modi file d:\转移说明.txt
set device to printer
set printer on
set print font "宋体",14
@ 3 , 0 say '               转移养老保险基金的说明 '
GO BOTT
@ prow()+1 , 0 say '     '+allt(ryxm)+'同志于'+工作+'月参加工作，根据云劳（1998）72号文第十三条:'
@ prow()+1 , 0 say '“从1998年4月1日起，从业人员跨省统筹范围转移时，转出地(部门)的社会'
@ prow()+1 , 0 say '保险机构应向转入地(部门）的社会保险机构转移其养老保险关系、个人帐户'
@ prow()+1 , 0 say '档案和个人帐户中1998年4月1日之前的个人缴费部份累计本息加上1998年4月'
@ prow()+1 , 0 say '1日后的全部储存额。”的规定，其中：'
@ prow()+2 , 0 say '     一、1998年4月1日之前的个人缴费部份累计本息为'+str(个人97本息,6,1)+'元。'
@ prow()+2 , 0 say '     (大写:                   )。'
@ prow()+2 , 0 say '     二、1998年4月1日起记入个人帐户的全部储存额为'+str(总累计,8,1)+'元。'
@ prow()+2 , 0 say '     (大写:                   )。'
@ prow()+2 , 0 say '     以上两项合计金额为'+str(个人97本息+总累计,10,1)+'元(大写:                      )。'
@ prow()+6 , 0 say '    转出单位：（盖章）              审核单位：（盖章）'
@ prow()+5 , 0 say '    经办人：'+ZB111+'               审核人:'
@ prow()+6 , 0 say '    转出日期：    年   月   日      审核日期：     年   月   日'
set print font "宋体",12
@ prow()+5 , 0 say '  填写说明：(1)、1998年4月1日之前的个人缴费部份累计本息为“1997年当年和历年个人缴费本息4个数字之和”'
@ prow()+1 , 0 say '  (2)、1998年4月1日起记入个人帐户的全部储存额为1998年4月统筹后至转出当年的“总计”数。'
@ prow()+1 , 0 say '  (3)、转出当年缴费不足12个月时暂不结息，“总计”为“转出当年和历年个人缴费本金2个数字之和”。'   
@ prow()+2 , 0 say '  打印说明：请关闭此窗口后立即打印或退出后用WORD字处理软件编辑打印“D:\转移说明.txt”'
set device to screen
?''
set print off
set print to
clear
 _SCREEN.WINDOWSTATE = 2
close all
use ccrq
go 8
repl 图片 with str(val(图片)+1,2)
if allt(图片)='10'
   repl 图片 with '0'
endi   
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'

