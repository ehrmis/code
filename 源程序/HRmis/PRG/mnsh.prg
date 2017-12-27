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
 =MESSAGEBOX(CN+'��'+CM+'�¹���û�м��㣡', 16,"��ʾ")
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
copy to ���� for ((b.ylbx<>ylbx or b.sybx<>sybx or b.ybx<>ybx) and month(date())<>3) or sds<0
YN = MESSAGEBOX('��Ҫȫ����Աһ������𣿣�...����ÿ�˷ֱ���ˣ�',4+32,'��ʾ')
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
        zdm='��λ����'       
   case uppe(allt(zd))='GWJT'
        zdm='�ϸڽ���'       
   case uppe(allt(zd))='ZF'
        zdm='֪��'       
   case uppe(allt(zd))='YLBX'
        zdm='���ϱ���'      
   case uppe(allt(zd))='YBX'
        zdm='ҽ�Ʊ���'    
   case uppe(allt(zd))='SYBX'
        zdm='ʧҵ����'
   case uppe(allt(zd))='SDS'
        zdm='����˰'
        if &zd>=0
           exit
        endi                
   endcase                            
        brow field rybh:h='��Ա���',ryxm:h='��Ա����',&ZD.:12:h='��&ZDm':10,b.&ZD1.:12:h='ԭ&ZDm':10,c.jrjb:h='���ռӰ�',;
        c.bj:h='����',c.sj:h='�¼�',c.cj:h='����',c.gs:h='����';
,c.tqj:h='̽�׼�',c.hsj:h='��ɥ��',c.jl:h='����',c.kg:h='����' colo scheme 5 for recn()=jlh when TS_TS1() titl '����ְ������ͬ������Ƚ��б仯 ��[ Esc ]��ֹ����[Ctrl+End]�������...'
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
browse field RYBH :H = '�� ��' , RYXM :H = ' ��  �� ' , ;
 BZGZ :P = '@z 9999.99' :H = '��λ����';
 , GWJT :H = '�ϸڽ���';
 :P = '@z 999.9';
 ,zf:h='֪��' ;
 :P = '@z 999.9';
 , YBX :H = 'ҽ�Ʊ���' :P = '@z 9999.99', YLBX :H = '���ϱ���' :P = '@z 9999.99', sybx :H = 'ʧҵ����' ;
 :P = '@z 9999.99'  noedit color scheme;
 5 when TS_TS() titl '�������������ְ������ ��[����]��ˣ���[ F1 ]��ѯ���ˣ���[Ctrl+End]����....'
sele 1
use ����
go top
set relation to RYBH into 4
set relation to RYBH into 3 additive
set relation to RYBH into 2 additive
if recc()>0
browse field RYBH :H = '�� ��' , RYXM :H = ' ��  �� ' , ;
 BZGZ :P = '@z 9999.99' :H = '��λ����';
 , GWJT :H = '�ϸڽ���';
 :P = '@z 999.9';
 ,zf:h='֪��' ;
 :P = '@z 999.9';
 , YBX :H = 'ҽ�Ʊ���' :P = '@z 9999.99' , YLBX :H = '���ϱ���' :P = '@z 9999.99', sybx :H = 'ʧҵ����' ;
 :P = '@z 9999.99'  , SDS :H = '����˰' :P = '@z 9999.99' noedit color scheme;
 5 when TS_TS() titl '����ְ�����ʺܿ��ܳ���'
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
delete file ����.dbf
delete file ����.idx

PROCEDURE TS_TS
wait wind alltrim(RYXM)+'���¹������('+allt(d.zgqk1)+'):'+'��λ����:'+alltrim(str(c.BZGZ,6,1))+';�ϸڽ���:'+alltrim(str(c.GWJT,5,1))+';֪��:'+alltrim(str(c.zf,5,1))+';ҽ�Ʊ���:'+alltrim(str(c.YBX,10,6)) nowa at 25,15


PROCEDURE TS_TS1
wait wind alltrim(RYXM)+'���¿������('+allt(e.zgqk1)+'):'+'�ڼ�:'+alltrim(str(d.JrJB))+';����:'+alltrim(str(d.BJ))+';'+'�¼�:'+alltrim(str(d.SJ))+';'+'����:'+alltrim(str(d.CJ))+';����:'+alltrim(str(d.GS))+';̽��:'+alltrim(str(d.TQJ))+';��ɥ:'+alltrim(str(d.HSJ))+';����:'+alltrim(str(d.JL))+';����:'+alltrim(str(d.KG)) nowa at 15,15


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
        zdm='��λ����'       
   case uppe(allt(zd))='GWJT'
        zdm='�ϸڽ���'       
   case uppe(allt(zd))='ZF'
        zdm='֪��'       
   case uppe(allt(zd))='YLBX'
        zdm='���ϱ���'      
   case uppe(allt(zd))='YBX'
        zdm='ҽ�Ʊ���'    
   case uppe(allt(zd))='SYBX'
        zdm='ʧҵ����'
   case uppe(allt(zd))='SDS'
        zdm='����˰'                
endcase                 
brow titl '��[����]��ˣ���[Ctrl+End]����....' field rybh:h='�� ��',ryxm:h=' ��  �� ',&ZD.:12:h='��&ZDm':10,b.&ZD1.:12:h='ԭ&ZDm':10,e.jrjb:h='���ռӰ�';
,e.bj:h='����',e.sj:h='�¼�',e.cj:h='����',e.gs:h='����';
,e.tqj:h='̽��',e.hsj:h='��ɥ' ,e.jl:h='����',e.kg:h='����' colo scheme 5;
 when TS_TS()
 
