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
 =MESSAGEBOX(CN+'��'+CM+'�¹���û�м��㣡', 16,"��ʾ")
 return
else
use lw&cn1&cm1
endif
copy to abc field RYBH , RYXM , gwGZ , ;
 �������, ���ܽ���,zybf,zbgr,ybgr
use lw&cn&cm
copy to abc1 field RYBH , RYXM , gwGZ , ;
 �������, ���ܽ���,zybf,zbgr,ybgr
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
 or b.�������<>������� ;
 or b.���ܽ���<>���ܽ��� or b.zybf<>zybf or b.zbgr<>zbgr or b.ybgr<>ybgr
copy to ���� for (b.gwgz<>gwgz or b.�������<>������� or b.���ܽ���<>���ܽ���)
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
   case uppe(allt(zd))='gwgz'
        zdm='��λ����'       
   case uppe(allt(zd))='�������'
        zdm='�������'       
   case uppe(allt(zd))='���ܽ���'
        zdm='���ܽ���'       
   case uppe(allt(zd))='zYBf'
        zdm='��ҹ�����'      
   case uppe(allt(zd))='zBgr'
        zdm='�а๤��'    
   case uppe(allt(zd))='YBgr'
        zdm='ҹ�๤��'                      
   endcase                            
        brow field rybh:h='��Ա���',ryxm:h='��Ա����',&ZD.:12:h='��&ZDm':10,b.&ZD1.:12:h='ԭ&ZDm':10;
         colo scheme 5 for recn()=jlh when TS_TS1() titl '����ְ������ͬ������Ƚ��б仯 ��[ Esc ]��ֹ����[Ctrl+End]�������...'
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
browse field RYBH :H = '�� ��' , RYXM :H = ' ��  �� ' , ;
 gwGZ :P = '@z 9999.99' :H = '��λ����';
 , �������;
 :P = '@z 999.9';
 ,���ܽ���;
 :P = '@z 999.9';
 , zYBf :H = '��ҹ�����' :P = '@z 9999.99', zBgr :H = '�а๤��' :P = '@z 9999.99', ybgr :H = 'ҹ�๤��' ;
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
 gwGZ :P = '@z 9999.99' :H = '��λ����';
 , �������;
 :P = '@z 999.9';
 ,���ܽ��� ;
 :P = '@z 999.9';
 , zYBf :H = '��ҹ�����' :P = '@z 9999.99' , zBgr :H = '�а๤��' :P = '@z 9999.99', ybgr :H = 'ҹ�๤��' ;
 :P = '@z 9999.99' noedit color scheme;
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
wait wind alltrim(RYXM)+'���¹������:'+'��λ����:'+alltrim(str(c.gwGZ,6,1))+';�������:'+alltrim(str(c.�������,5,1))+';���ܽ���:'+alltrim(str(c.���ܽ���,5,1))+';��ҹ�����:'+alltrim(str(c.zYBf,10,6)) nowa at 25,15


PROCEDURE TS_TS1
wait wind alltrim(RYXM)+'���¿������:'+'�а๤��:'+alltrim(str(b.zbgr))+';ҹ�๤��:'+alltrim(str(b.ybgr)) nowa at 15,15


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
        zdm='��λ����'       
   case uppe(allt(zd))='�������'
        zdm='�������'       
   case uppe(allt(zd))='���ܽ���'
        zdm='���ܽ���'       
   case uppe(allt(zd))='zYBf'
        zdm='��ҹ�����'      
   case uppe(allt(zd))='zBgr'
        zdm='�а๤��'    
   case uppe(allt(zd))='YBgr'
        zdm='ҹ�๤��'                      
endcase                        
brow titl '��[����]��ˣ���[Ctrl+End]����....' field rybh:h='�� ��',ryxm:h=' ��  �� ',&ZD.:12:h='��&ZDm':10,b.&ZD1.:12:h='ԭ&ZDm':10,;
 colo scheme 5;
 when TS_TS()
 
