* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
*  �ļ���: LY_YLBX.PRG <-- ���ļ��� UnFoxAll ����
* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
 CLOSE ALL
 ON KEY LABEL F1 do grcx
 set safety off
 DH = ''
 XM = ''
 DYCJ = ''
 DYBM = ''
 LY=ND
 IF  .NOT. FILE('YL' + LY + '.dbf')
    USE ylbxk
     copy to yl&LY stru
     use yl&LY
    IF FILE('zr' + LY + '.dbf')
        append from zr&LY
    ELSE 
        MESSAGEBOX('���Ƚ���' + LY + '�깤���ܶ������','��ʾ')
       RETURN 
    ENDIF 
    REPLACE ��� WITH (LY)
    REPLACE ZBH WITH ALLTRIM(CJDM) + '000' + STR(RECNO(),1) FOR RECNO() < 10
    REPLACE ZBH WITH ALLTRIM(CJDM) + '00' + STR(RECNO(),2) FOR  ;
         RECNO() >= 10 AND RECNO() < 100
    REPLACE ZBH WITH ALLTRIM(CJDM) + '0' + STR(RECNO(),3) FOR  ;
         RECNO() >= 100 AND RECNO() < 1000
    REPLACE ZBH WITH ALLTRIM(CJDM) + STR(RECNO(),4) FOR RECNO() >= 1000 AND RECNO() < 10000
 ENDIF 
 DO DHKDY
************************ͨ�ù��ܣ��غż��***********
use yl&LY
repl a with 1 all
sort on rybh to temp
use temp excl
go top
do while !eof()
   bh=rybh
   skip
   if bh=rybh
      repl a with 0
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a=0 to aaa
if aaa>0
   inde on bmbh+rybh to temp
   go top
   brow for a=0 titl '����'+allt(str(aaa,4))+'���غţ����޸���ȷ��[ Ctrl+T ]ɾ�������¼'
   pack
   delete file yl&LY..dbf
   sort on bmbh,rybh to yl&LY
endif
close all
use yl&LY excl
 DO CASE 
 CASE DH = 2
     set filter to CJdm='&dycj'
 CASE DH = 3
     set filter to bmbh='&dybm'
 ENDCASE 
 inde on bmbh+rybh to yl&ly
 GO TOP
 for zdz=1 to fcount()
  zd=fiel(zdz)
  zdz1=(fsize(zd)+fsize(zd))*2.4
  if uppe(allt(zd))='RYXM' 
     exit
  endif
  endfor   
 BROWSE part zdz1 FIELDS CJMC :H = '��������' : 12 :R , BMMC :H = '��������' : 12 :R ,  ;
      RYXM :H = '  ����  ' : 8 :R , �ɷ����� : 8 :H = '�ɷ�����' , �ɷѻ���  ;
      : 8 :H = '��ɷѻ���' , ��λ�ɷ� : 12 :H = '���굥λ�ɷ�' , ��λ��Ϣ : 12  ;
      :H = '���굥λ��Ϣ' , ���˽ɷ� : 12 :H = '������˽ɷ�' , ������Ϣ :  ;
      12 :H = '���������Ϣ' , ����ϼ� : 10 :H = '����ϼ�' , ���굥λ�� :  ;
      12 :H = '���굥λ�ɷ�' , ������Ϣ1 : 12 :H = '���굥λ��Ϣ' ,  ;
      ������˽� : 12 :H = '������˽ɷ�' , ������Ϣ2 : 12 :H = '���������Ϣ' ,  ;
      ���ۼ� : 10 :H = '��    ��' TITLE  ;
      LY + '������ʻ�����[ F1 = ������Ա    Esc = �˳� ]'
 SET FILTER TO
 CLOSE ALL
USE deset
LOCATE FOR oop='backup'
I=seted
  use yl&LY excl
  pack
 SORT ON BMBH , RYBH TO temp
 USE temp
  copy to yl&LY
 CLOSE ALL
use yl&LY
WHNY=LY+'�����ϱ���'
DBFF1='yl'
do Qdir
close all
 DELETE File temp.dbf
 ON KEY LABEL F1 
retu
*
