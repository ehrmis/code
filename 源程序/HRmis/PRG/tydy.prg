***************************
* .\TYDY.PRG
***************************
close all
_SCREEN.PICTURE = ''
XM = '        '
连续工龄 = space(5)
计发比例 = space(4)
store 0 to JN待遇 , GW待遇 , GL待遇 , JT待遇 , SB待遇 , XL待遇 , 小计 , 代扣 , 实发
store 1 to 连续工龄 , JLH , YX1
YF1=month(date())
ND1=year(date())
do srrq.spx
NY = ND+YF
use tydy excl
zap
append from ryzk
repl shf with 35 for shf=0
close all
select 3
use zr1&ly1
select 2
use ryzk
index on RYBH to ry
select 1
use tydy
set relation to RYBH into 2
repl shf with 0 for b.tybl='100'
select 2
go top
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
  select 3
  go top
  locate for RYXM=left(XM,8) or RYBH=left(upper(XM),5)
  上年平均=pj
  select 1 
  go JLH
  T1 = '  单位名称:&MC111'+'                                金额单位:元(保留角)' 
  T2 = '┏━━━┯━━━━┯━┯━┯━━┯━━━━━┯━━━┯━━━━━┯━━┯━━━━━┓'
  T3 = '┃      │        │别│  │年月│          │作日期│          │年月│;
          ┃'
  T4 = '┠───┴────┴─┴┬┴──┴─────┼───┴┬────┴──┴─────┨'
  T5 = '┠─┬─────────┴────────┰┴┬───┴─────────────┨'
  T6 = '┃序│   应      发      金      额       ┃序│   代      扣   ;
   金      额     ┃'
  T7 = '┃  ├──────┬───┬──┬────┨  ├─────┬───┬──┬────┨'
  T8 = '┃号│  项      目│原金额│比例│ 金  额 ┃号│ 项    目 │原金额│比例│;
 金  额 ┃'
  T9 = '┠─┼──────┼───┼──┼────╂─┼─────┼───┼──┼────┨'
  T10 = '┠─┼──────┼───┼──┼────╂─┼─────┼───┼──┼────┨'
  T11 = '┠─┼──────┴───┴──┼────╂─┼─────┴───┴──┼────┨'
  连续工龄 = val(ND)-val(left(b.CJGZRQ,4))+round((val(YF)-val(substr(b.CJGZRQ,5,2)))/12,2)
  工龄 = val(ND)-val(left(b.CJGZRQ,4))+1
  医疗费 = 2002-val(left(b.CJGZRQ,4))+1+20
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
  case 连续工龄<25 and b.XB='0'
    计发比例 = ' 65'
  case 连续工龄<30 and b.XB='1'
    计发比例 = ' 65'
  case 连续工龄>=25 and 连续工龄<30 and b.XB='0'
    计发比例 = ' 70'
  case 连续工龄>=30 and 连续工龄<35 and b.XB='1'
    计发比例 = ' 70'
  case 连续工龄>=30 and b.XB='0'
    计发比例 = ' 75'
  case 连续工龄>=35 and b.XB='1'
    计发比例 = ' 75'
  endcase
  if (b.NL>=55 and b.XB='1') or (b.NL>=45 and b.XB='0')
    if 计发比例<' 70'
      计发比例 = ' 70'
    endif
  endif
  if b.tybl='100'
     计发比例 = '100'
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
  代扣 = b.FZSD+b.YLBX+b.YBX+b.SYBX+b.gjj
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
  set device to printer
  set printer on
  go JLH  
    @ 3 , 5 say '昆明钢铁集团有限责任公司职工退养待遇预测表' font '宋体',20
    set print font "宋体",12
    @ prow()+2 , 0 say T1
    @ prow()+0.8 , 0 say T2
    @ prow()+0.8 , 0 say '┃姓  名│'+b.RYXM+'│性│'+b.XB1+'│出生│'+left(b.CSRQ;
,4)+'年'+str(val(substr(b.CSRQ,5,2)),2)+'月'
    @ prow() , pcol() say '│参加工│'+left(b.CJGZRQ,4)+'年'+str(val(substr(b.CJGZRQ;
,5,2)),2)+'月'
    @ prow() , pcol() say '│退养│'+ND+'年'+str(val(YF),2)+'月┃'
    @ prow()+0.8 , 0 say T3
    @ prow()+0.8 , 0 say T4
    @ prow()+0.8 , 0 say '┃原从事工种(职务)名称  │'+Gw+'  │连续工龄│'+left(str(连续工龄;
,5,2),2)+'年零'+str(val(right(str(连续工龄,5,2),2))/100*12,2)+'个月 ;
             ┃'
    @ prow()+0.8 , 0 say T5
    @ prow()+0.8 , 0 say T6
    @ prow()+0.8 , 0 say T7
    @ prow()+0.8 , 0 say T8
    @ prow()+0.8 , 0 say T9
    @ prow()+0.8 , 0 say '┃1 │技能工资    │'
    @ prow() , pcol() say JNGZ picture '9999.9'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say left(计发比例,3)+'%'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say JN待遇 picture '99999.99'
    @ prow() , pcol() say '┃12│房租水电费│      │    │'
    @ prow() , pcol() say b.FZSD picture '99999.99'
    @ prow() , pcol() say '┃'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '┃2 │岗位工资    │'
    @ prow() , pcol() say GWGZ picture '9999.9'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say left(计发比例,3)+'%'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say GW待遇 picture '99999.99'
    @ prow() , pcol() say '┃13│住房公积金│      │    │'
    @ prow() , pcol() say b.gjj picture '99999.99'
    @ prow() , pcol() say '┃'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '┃3 │工龄工资    │'
    @ prow() , pcol() say 工龄工资 picture '999.99'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say left(计发比例,3)+'%'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say GL待遇 picture '99999.99'
    @ prow() , pcol() say '┃14│养老保险金│      │    │'
    @ prow() , pcol() say b.YLBX picture '99999.99'
    @ prow() , pcol() say '┃'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '┃4 │水电补贴    │'
    @ prow() , pcol() say SDBT picture '999.99'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say '100%'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say SDBT picture '99999.99'
    @ prow() , pcol() say '┃15│医疗保险金│      │    │'
    @ prow() , pcol() say b.YBX picture '99999.99'
    @ prow() , pcol() say '┃'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '┃5 │交通补贴    │'
    @ prow() , pcol() say JTBT picture '999.99'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say left(计发比例,3)+'%'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say JT待遇 picture '99999.99'
    @ prow() , pcol() say '┃16│失业保险金│      │    │'
    @ prow() , pcol() say SYBX picture '99999.99'
    @ prow() , pcol() say '┃'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '┃6 │书报费      │'
    @ prow() , pcol() say SBF picture '999.99'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say left(计发比例,3)+'%'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say SB待遇 picture '99999.99'
    @ prow() , pcol() say '┃17│          │      │    │        ┃'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '┃7 │洗理费      │'
    @ prow() , pcol() say XLF picture '999.99'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say left(计发比例,3)+'%'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say XL待遇 picture '99999.99'
    @ prow() , pcol() say '┃18│          │      │    │        ┃'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '┃8 │医药费      │'
    @ prow() , pcol() say 医疗费 picture '999.99'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say '100%'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say 医疗费 picture '99999.99'
    @ prow() , pcol() say '┃19│          │      │    │        ┃'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '┃9 │煤贴        │'
    @ prow() , pcol() say MT picture '999.99'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say '100%'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say MT picture '99999.99'
    @ prow() , pcol() say '┃20│          │      │    │        ┃'
    @ prow()+0.8 , 0 say T10
    @ prow()+0.8 , 0 say '┃10│回贴、生活费│'
    @ prow() , pcol() say HT+SHF picture '999.99'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say '100%'
    @ prow() , pcol() say '│'
    @ prow() , pcol() say HT+SHF picture '99999.99'
    @ prow() , pcol() say '┃21│          │      │    │        ┃'
    @ prow()+0.8 , 0 say T11
    @ prow()+0.8 , 0 say '┃11│小计(1+2+3+4+5+6+7+8+9+10)│'
    @ prow() , pcol() say 小计 picture '99999.99'
    @ prow() , pcol() say '┃22│     月 实 领 金 额     │'
    @ prow() , pcol() say 实发 picture '99999.99'
    @ prow() , pcol() say '┃'
    @ prow()+0.8 , 0 say '┠─┼─────────────┴────┸─┴────────────┴────┨'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃备│                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃注│                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┃  │                                          ;
                                  ┃'
    @ prow()+0.8 , 0 say '┗━┷━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛'   
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
