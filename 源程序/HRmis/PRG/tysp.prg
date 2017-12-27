***************************
* .\TYSP.PRG
***************************
set talk off
close all
_SCREEN.PICTURE = ''
XM = '        '
连续工龄 = space(5)
电话 = space(12)
store 1 to 连续工龄 , JLH , YX1
YF1=month(date())
ND1=year(date())
do srrq.spx
NY = ND+YF
clear
use ryzk
do while .t.      
      do xmsr.spr
      locate for upper(XM)=RYXM or upper(XM)=RYBH
      if eof()
        do dhkjg.spx
    M = readkey()
    if M=12 or M=268
      return
    endif
      else
      xm=allt(ryxm)
      clear
      @ 15 , 38 say '请确认'+xm+'出生年月:' font '宋体',20 get CSRQ  font '宋体',20 picture '99999999'
      read
      @ 18 , 38 say '请确认'+xm+'参加工作日期:' font '宋体',20 get CJGZRQ  font '宋体',20 picture '9999999999' 
      read
      replace NL with val(ND)-val(left(CSRQ,4))+round((val(YF)-val(substr(CSRQ,5,2)))/12,2)
      JLH = recno()
        exit
      endif
enddo    
  go JLH
T1 = '   单位：&MC111'+space(25)+'公司审批编号：' 
T2='┏━━━━┯━━━┯━━┯━━━━━━┯━━━━┯━━━━━┯━━━━┯━━━━━┓'   
T3='┠────┼───┼──┼──────┼────┼─────┼────┼─────┨'    
T4='┠────┼───┴──┴──────┼────┼─────┼────┴─────┨'    
T5='┠────┼─────────────┴────┴─────┴──────────┨'
T6='┠────┼───────────────────────────────────┨'
T7='┗━━━━┷━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛'
  SR1 = ND
  SR2 = str(val(YF),2)
  SR4 = str(val(substr(CSRQ,5,2)),2)
  do case
  case XB='0' and empty(ZCJT)
    SR3 = val(left(CSRQ,4))+50
  case XB='0' and  not empty(ZCJT)
    SR3 = val(left(CSRQ,4))+55
  case XB='1'
    SR3 = val(left(CSRQ,4))+60
  endcase
  SR3 = str(SR3,4)
  clear
  @ 15 , 38 say '退养时间:' font '宋体',20 get SR1 font '宋体',20 picture '9999'
  @ row() , col()+2 say '年' font '宋体',20 get SR2 font '宋体',20 picture '99'
  @ row() , col()+2 say '月至法定退休' font '宋体',20 get SR3 font '宋体',20 picture '9999'
  @ row() , col()+2 say '年' font '宋体',20 get SR4 font '宋体',20 picture '99'
  @ row() , col()+2 say '月' font '宋体',20
  read  
  @ 18 , 38 say '请确认'+xm+'家庭地址:' font '宋体',20 get ZZ font '宋体',20
  read
   连续工龄 = val(ND)-val(left(CJGZRQ,4))+round((val(YF)-val(substr(CJGZRQ;
,5,2)))/12,2)
***************
  工龄 = val(ND)-val(left(CJGZRQ,4))+1
  医疗费 = 2002-val(left(CJGZRQ,4))+1+20
  if 连续工龄>=25
     工龄工资 = round(工龄*2.4,1)
  else
     工龄工资 = round(工龄*2,1)
  endif
  clear
  @ 15 , 38 say '请确认'+xm+'当前连续工龄:' font '宋体',20 get 连续工龄 font '宋体',20 picture;
 '99.99' 
  @ row() , col() say '年' font '宋体',20
  read
  do case
  case 连续工龄<25 and XB='0'
    计发比例 = ' 65'
  case 连续工龄<30 and XB='1'
    计发比例 = ' 65'
  case 连续工龄>=25 and 连续工龄<30 and XB='0'
    计发比例 = ' 70'
  case 连续工龄>=30 and 连续工龄<35 and XB='1'
    计发比例 = ' 70'
  case 连续工龄>=30 and XB='0'
    计发比例 = ' 75'
  case 连续工龄>=35 and XB='1'
    计发比例 = ' 75'
  endcase
  if (NL>=55 and XB='1') or (NL>=45 and XB='0')
    if 计发比例<' 70'
      计发比例 = ' 70'
    endif
  endif
  if tybl='100'
     计发比例 = '100  '
  endif
  @ 18 , 38 say '请确认'+xm+'退养工资计发比例(%):' font '宋体',20 get 计发比例 font '宋体',20
  read
  JN待遇 = round(val(left(计发比例,3))/100*JNGZ,1)
  GW待遇 = round(val(left(计发比例,3))/100*GWGZ,1)
  GL待遇 = round(val(left(计发比例,3))/100*工龄工资,1)
  JT待遇 = round(val(left(计发比例,3))/100*JTBT,1)
  XL待遇 = round(val(left(计发比例,3))/100*XLF,1)
  SB待遇 = round(val(left(计发比例,3))/100*SBF,1)
  小计 = JN待遇+GW待遇+GL待遇+JT待遇+XL待遇+SB待遇+医疗费+HT+MT+SDBT+SHF
  代扣 = FZSD+YLBX+YBX+SYBX+gjj
  实发 = round(小计-代扣,1)
  clear
  @ 20 , 20 say '退养工资:应发 'font '宋体',20 
  @ 20 , col() say alltrim(str(小计,8,1)) font '宋体',20 
  @ 20 , col() say '元 — 代扣 ' font '宋体',20 
  @ 20 , col() say alltrim(str(代扣,8,1)) font '宋体',20 
  @ 20 , col() say '元　= 实领 ' font '宋体',20 
  @ 20 , col() say alltrim(str(实发,8,1)) font '宋体',20 
  @ 20 , col() say '元 ' font '宋体',20 
  = inkey(30,'M')  
***************
set device to printer
set printer on
  go JLH  
    @ 3 , 12 say '昆明钢铁集团有限责任公司职工退养审批表' font '宋体',20
    set print font "宋体",12
    @ prow()+3 , 0 say T1 
    @ prow()+2 , 0 say T2
    @ prow()+0.8 , 0 say '┃姓    名│'+left(RYXM,6)+'│性别│     '+XB1+'     │出生年月│'+left(CSRQ;
,4)+'年'+str(val(substr(CSRQ,5,2)),2)+'月'
    @ prow() , pcol() say '│工作时间│'+left(CJGZRQ,4)+'年'+str(val(substr(CJGZRQ;
,5,2)),2)+'月┃'
    @ prow()+0.8 , 0 say T3
    @ prow()+0.8 , 0 say '┃文化程度│ '+WHCD2+' │工种│'+Gz1
    @ prow() , pcol() say '│退养工龄│'+left(str(连续工龄,5,2),2)+'年'+str(val(right(str(连续工龄;
,5,2),2))/100*12,2)+'个月│享受比例│  '+left(计发比例,3)+'%    ┃'
    @ prow()+0.8 , 0 say T4
    @ prow()+0.8 , 0 say '┃        │                          │        │ 应发金额 │'+str(小计,10,2)+'元        ┃'
    @ prow()+0.8 , 0 say '┃家庭住址│'+ZZ+space(6)+'│退养工资├─────┼──────────┨'
    @ prow()+0.8 , 0 say '┃        │                          │        │ 实发金额 │'+str(实发,10,2)+'元        ┃'
    @ prow()+0.8 , 0 say T5
    @ prow()+0.8 , 0 say '┃        │                                                                      ┃'
    @ prow()+0.8 , 0 say '┃        │                                                                      ┃'
    @ prow()+0.8 , 0 say '┃退养原因│                                                                      ┃'
    @ prow()+0.8 , 0 say '┃        │                                                                      ┃'
    @ prow()+0.8 , 0 say '┃        │                                          年     月     日            ┃'
    @ prow()+0.8 , 0 say T6
    @ prow()+0.8 , 0 say '┃退养时间│'+space(12)+SR1+' 年'+SR2+'月至法定退休年月'+'                                 ┃'
    @ prow()+0.8 , 0 say T6
@ prow()+0.8 , 0 say '┃   二   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   级   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   单   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   位   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   意   │                                          年     月     日            ┃'
@ prow()+0.8 , 0 say '┃   见   │                                                                      ┃'
@ prow()+0.8 , 0 say T6
@ prow()+0.8 , 0 say '┃   集   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   团   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   公   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   司   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   审   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   批   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   意   │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   见   │                                          年     月     日            ┃'
@ prow()+0.8 , 0 say T6
@ prow()+0.8 , 0 say '┃        │                                                                      ┃'
@ prow()+0.8 , 0 say '┃        │                                                                      ┃'
@ prow()+0.8 , 0 say '┃   备   │1、“公司审批编号”由集团公司劳动人事处填写。                         ┃'
@ prow()+0.8 , 0 say '┃        │2、退养由集团公司劳动人事处审批。                                     ┃'
@ prow()+0.8 , 0 say '┃   注   │3、本表一式二份，单位存档一份，审批机关存档一份。                     ┃'
@ prow()+0.8 , 0 say '┃        │                                                                      ┃'
@ prow()+0.8 , 0 say '┃        │                                                                      ┃'
@ prow()+0.8 , 0 say T7
@ prow()+1 ,0 say '　　  职工本人签名：            单位经办人：'+'&ZB111'+'    公司审批经办人：'      
set device to screen
?''
set print off
set print to
clear
 _SCREEN.WINDOWSTATE = 2
close all
use data\ccrq
go 8
repl 图片 with str(val(图片)+1,2)
if allt(图片)='10'
   repl 图片 with '0'
endi   
pict_fd='fd'+allt(图片)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
_SCREEN.MAXBUTTON = .F.
_SCREEN.MOVABLE = .F.
_SCREEN.CONTROLBOX = .F.
_SCREEN.CLOSABLE = .F.
close all
return
