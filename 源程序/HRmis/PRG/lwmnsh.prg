*******************************
* .\lwMNSH.PRG (Build 20150917)
*******************************
close all
set talk off
set safety off
set stat off
set escape off
on key label F1 do TS_CX
clear
CN = nd
CN1 = iif(val(yf)<>1,nd,str(val(nd)-1,4))
CM = yf
CM1 = iif(val(yf)-1>9,str(val(yf)-1,2),'0'+str(val(yf)-1;
,1))
CM1 = iif(val(yf)<>1,CM1,'12')
if  not file('lw'+CN+CM+'.DBF')
 =MESSAGEBOX(CN+'年'+CM+'月工资没有计算！', 16,"提示")
 return
else
use lw&cn1&cm1
endif
copy to abc field RYBH , RYXM , gwGZ , ;
 带班津贴, 技能津贴,zybf,zbgr,ybgr
use lw&cn&cm
copy to abc1 field RYBH , RYXM , gwGZ , ;
 带班津贴, 技能津贴,zybf,zbgr,ybgr
select 3
use ryzk
index on RYBH to ry
select 2
use abc
index on RYBH to abc
go top
select 1
use abc1
set relation to RYBH into 3 
set relation to RYBH into 2 additive
copy to abc2 for b.RYXM<>RYXM or  b.gwgZ<>gwGZ ; 
 or b.带班津贴<>带班津贴 ;
 or b.技能津贴<>技能津贴 or b.zybf<>zybf or b.zbgr<>zbgr or b.ybgr<>ybgr
copy to 错误 for (b.gwgz<>gwgz or b.带班津贴<>带班津贴 or b.技能津贴<>技能津贴)
YN = MESSAGEBOX('需要全部人员一起审核吗？！...否则将每人分别审核！',4+32,'提示')
go top
do while  not eof()
  for J0 = 1 to fcount()
    ZD = field(J0)
    ZD1 = field(J0,2)
    if &ZD<>b.&ZD1
       if yn=7
        JLH = recno()
        clear
   do case
   case uppe(allt(zd))='gwgz'
        zdm='岗位工资'       
   case uppe(allt(zd))='带班津贴'
        zdm='带班津贴'       
   case uppe(allt(zd))='技能津贴'
        zdm='技能津贴'       
   case uppe(allt(zd))='zYBf'
        zdm='中夜班津贴'      
   case uppe(allt(zd))='zBgr'
        zdm='中班工日'    
   case uppe(allt(zd))='YBgr'
        zdm='夜班工日'                      
   endcase                            
        brow field rybh:h='人员编号',ryxm:h='人员姓名',&ZD.:12:h='现&ZDm':10,b.&ZD1.:12:h='原&ZDm':10;
         colo scheme 5 for recn()=jlh when TS_TS1() titl '下列职工工资同上月相比较有变化 按[ Esc ]终止；按[Ctrl+End]继续审核...'
        endif
    else
      if J0>=4
        repl &ZD with 0
      endi
    endif
  endfor
  if readkey()=12 or readkey()=268
     clear windows all
     exit
  endif   
  skip
enddo
close all
select 4
use ryzk
index on RYBH to ry
select 3
use abc index abc
select 2
use abc1
index on RYBH to abc1
select 1
use abc2
set relation to RYBH into 4 
set relation to RYBH into 3 additive
set relation to RYBH into 2 additive
go top
do while  not eof()
  for J0 = 1 to fcount()
    ZD = field(J0)
    ZD1 = field(J0,2)
    repl &zd with b.&zd1
  endfor
  skip
enddo
go top
clear
browse field RYBH :H = '编 号' , RYXM :H = ' 姓  名 ' , ;
 gwGZ :P = '@z 9999.99' :H = '岗位工资';
 , 带班津贴;
 :P = '@z 999.9';
 ,技能津贴;
 :P = '@z 999.9';
 , zYBf :H = '中夜班津贴' :P = '@z 9999.99', zBgr :H = '中班工日' :P = '@z 9999.99', ybgr :H = '夜班工日' ;
 :P = '@z 9999.99'  noedit color scheme;
 5 when TS_TS() titl '请认真审核下列职工工资 按[↑↓]审核；按[ F1 ]查询个人；按[Ctrl+End]继续....'
sele 1
use 错误
go top
set relation to RYBH into 4
set relation to RYBH into 3 additive
set relation to RYBH into 2 additive
if recc()>0
browse field RYBH :H = '编 号' , RYXM :H = ' 姓  名 ' , ;
 gwGZ :P = '@z 9999.99' :H = '岗位工资';
 , 带班津贴;
 :P = '@z 999.9';
 ,技能津贴 ;
 :P = '@z 999.9';
 , zYBf :H = '中夜班津贴' :P = '@z 9999.99' , zBgr :H = '中班工日' :P = '@z 9999.99', ybgr :H = '夜班工日' ;
 :P = '@z 9999.99' noedit color scheme;
 5 when TS_TS() titl '下列职工工资很可能出错'
endif
close all
clear
delete file abc.dbf
delete file abc1.dbf
delete file abc.idx
delete file abc1.idx
delete file ry.idx
delete file kq.idx
delete file gz.idx
delete file gz1.idx
delete file 错误.dbf
delete file 错误.idx

PROCEDURE TS_TS
wait wind alltrim(RYXM)+'上月工资情况:'+'岗位工资:'+alltrim(str(c.gwGZ,6,1))+';带班津贴:'+alltrim(str(c.带班津贴,5,1))+';技能津贴:'+alltrim(str(c.技能津贴,5,1))+';中夜班津贴:'+alltrim(str(c.zYBf,10,6)) nowa at 25,15


PROCEDURE TS_TS1
wait wind alltrim(RYXM)+'上月考勤情况:'+'中班工日:'+alltrim(str(b.zbgr))+';夜班工日:'+alltrim(str(b.ybgr)) nowa at 15,15


PROCEDURE TS_CX
DO WHILE .T.
    DO srxm.spr
    LOCATE FOR ALLTRIM(RYXM) = ALLTRIM(upper(XM)) .OR. ALLTRIM(RYBH) = ALLTRIM(upper(XM))
    JLH=recn()
    IF EOF()
       DO dhkjg.spx
       M = READKEY()
       IF M = 12 .OR. M = 268
          RETURN 
       ENDIF 
    ELSE 
       go JLH
       EXIT 
    ENDIF 
 ENDDO 
do case
   case uppe(allt(zd))='gwgz'
        zdm='岗位工资'       
   case uppe(allt(zd))='带班津贴'
        zdm='带班津贴'       
   case uppe(allt(zd))='技能津贴'
        zdm='技能津贴'       
   case uppe(allt(zd))='zYBf'
        zdm='中夜班津贴'      
   case uppe(allt(zd))='zBgr'
        zdm='中班工日'    
   case uppe(allt(zd))='YBgr'
        zdm='夜班工日'                      
endcase                        
brow titl '按[↑↓]审核；按[Ctrl+End]继续....' field rybh:h='编 号',ryxm:h=' 姓  名 ',&ZD.:12:h='现&ZDm':10,b.&ZD1.:12:h='原&ZDm':10,;
 colo scheme 5;
 when TS_TS()
 
