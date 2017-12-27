*******************************
* .\MNSH.PRG (Build 20040517)
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
if  not file('GZ'+CN+CM+'.DBF')
 =MESSAGEBOX(CN+'年'+CM+'月工资没有计算！', 16,"提示")
 return
else
use gz&cn1&cm1
endif
copy to abc field RYBH , RYXM , BZGZ , ;
 GWJT ,zf, YLBX,ybx,sybx,sds
use gz&cn&cm
copy to abc1 field RYBH , RYXM , BZGZ ;
, gWJT ,zf, YLBX,ybx,sybx,sds
select 5
use ryzk
index on RYBH to ry
select 4
use kq&cn1&cm1
index on RYBH to gz1
select 3
use kq&cn&cm
index on RYBH to gz
select 2
use abc
index on RYBH to abc
go top
select 1
use abc1
set relation to RYBH into 5
set relation to RYBH into 4 additive
set relation to RYBH into 3 additive
set relation to RYBH into 2 additive
copy to abc2 for b.RYXM<>RYXM or  b.BZGZ<>BZGZ ;
 or b.GWJT<>GWJT ;
 or b.zf<>zf or b.YBX<>YBX or sds<0
copy to 错误 for ((b.ylbx<>ylbx or b.sybx<>sybx or b.ybx<>ybx) and month(date())<>3) or sds<0
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
   case uppe(allt(zd))='BZGZ'
        zdm='岗位工资'       
   case uppe(allt(zd))='GWJT'
        zdm='上岗津贴'       
   case uppe(allt(zd))='ZF'
        zdm='知补'       
   case uppe(allt(zd))='YLBX'
        zdm='养老保险'      
   case uppe(allt(zd))='YBX'
        zdm='医疗保险'    
   case uppe(allt(zd))='SYBX'
        zdm='失业保险'
   case uppe(allt(zd))='SDS'
        zdm='所得税'
        if &zd>=0
           exit
        endi                
   endcase                            
        brow field rybh:h='人员编号',ryxm:h='人员姓名',&ZD.:12:h='现&ZDm':10,b.&ZD1.:12:h='原&ZDm':10,c.jrjb:h='节日加班',;
        c.bj:h='病假',c.sj:h='事假',c.cj:h='产假',c.gs:h='工伤';
,c.tqj:h='探亲假',c.hsj:h='婚丧假',c.jl:h='拘留',c.kg:h='旷工' colo scheme 5 for recn()=jlh when TS_TS1() titl '下列职工工资同上月相比较有变化 按[ Esc ]终止；按[Ctrl+End]继续审核...'
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
select 5
use kq&cn&cm
index on RYBH to kq
select 4
use ryzk
index on RYBH to ryzk
select 3
use abc index abc
select 2
use abc1
index on RYBH to abc1
select 1
use abc2
set relation to RYBH into 5 
set relation to RYBH into 4 additive
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
 BZGZ :P = '@z 9999.99' :H = '岗位工资';
 , GWJT :H = '上岗津贴';
 :P = '@z 999.9';
 ,zf:h='知补' ;
 :P = '@z 999.9';
 , YBX :H = '医疗保险' :P = '@z 9999.99', YLBX :H = '养老保险' :P = '@z 9999.99', sybx :H = '失业保险' ;
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
 BZGZ :P = '@z 9999.99' :H = '岗位工资';
 , GWJT :H = '上岗津贴';
 :P = '@z 999.9';
 ,zf:h='知补' ;
 :P = '@z 999.9';
 , YBX :H = '医疗保险' :P = '@z 9999.99' , YLBX :H = '养老保险' :P = '@z 9999.99', sybx :H = '失业保险' ;
 :P = '@z 9999.99'  , SDS :H = '所得税' :P = '@z 9999.99' noedit color scheme;
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
wait wind alltrim(RYXM)+'上月工资情况('+allt(d.zgqk1)+'):'+'岗位工资:'+alltrim(str(c.BZGZ,6,1))+';上岗津贴:'+alltrim(str(c.GWJT,5,1))+';知补:'+alltrim(str(c.zf,5,1))+';医疗保险:'+alltrim(str(c.YBX,10,6)) nowa at 25,15


PROCEDURE TS_TS1
wait wind alltrim(RYXM)+'上月考勤情况('+allt(e.zgqk1)+'):'+'节假:'+alltrim(str(d.JrJB))+';病假:'+alltrim(str(d.BJ))+';'+'事假:'+alltrim(str(d.SJ))+';'+'产假:'+alltrim(str(d.CJ))+';工休:'+alltrim(str(d.GS))+';探亲:'+alltrim(str(d.TQJ))+';婚丧:'+alltrim(str(d.HSJ))+';拘留:'+alltrim(str(d.JL))+';旷工:'+alltrim(str(d.KG)) nowa at 15,15


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
   case uppe(allt(zd))='BZGZ'
        zdm='岗位工资'       
   case uppe(allt(zd))='GWJT'
        zdm='上岗津贴'       
   case uppe(allt(zd))='ZF'
        zdm='知补'       
   case uppe(allt(zd))='YLBX'
        zdm='养老保险'      
   case uppe(allt(zd))='YBX'
        zdm='医疗保险'    
   case uppe(allt(zd))='SYBX'
        zdm='失业保险'
   case uppe(allt(zd))='SDS'
        zdm='所得税'                
endcase                 
brow titl '按[↑↓]审核；按[Ctrl+End]继续....' field rybh:h='编 号',ryxm:h=' 姓  名 ',&ZD.:12:h='现&ZDm':10,b.&ZD1.:12:h='原&ZDm':10,e.jrjb:h='节日加班';
,e.bj:h='病假',e.sj:h='事假',e.cj:h='产假',e.gs:h='工伤';
,e.tqj:h='探亲',e.hsj:h='婚丧' ,e.jl:h='拘留',e.kg:h='旷工' colo scheme 5;
 when TS_TS()
 
