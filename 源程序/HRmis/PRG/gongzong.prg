close all
USE  ryzk
COPY TO temp FIELDS gz,gz1,gw,gwjb,gwjb1,bzgz,gwgz,gwfl1,zyflm,zymc,a
USE temp
INDEX on gz1+gw TO temp
TOTAL ON gz1+gw TO temp1
USE temp1
REPLACE gwgz WITH bzgz/a all
GO top
brow fiel gz:h='���ִ���',gz1:h='����',gw:h='��λ',gwjb1:h='��λ����',gwgz:h='ƽ����λ����',a:h='���±���λ����' titl '����λ��ǰ���ָ�λ������ܲ�ѯ'
copy to &bf.\���ִ��� type xl5
yn = MESSAGEBOX("����λ��ǰ���ָ�λ���ܵ����ɹ����ִ�ʹ����",4+32,"��ʾ")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\���ִ���.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
ENDIF
close all
***************************
yn = MESSAGEBOX("��Ҫ�ñ���λԱ����ǰ��λ�����׼���ʸ��¡����ִ���⡱��",4+32,"��ʾ")
IF yn = 6
CLOSE all
SELECT 3
USE gwdmk
INDEX on gwdm TO gwdmk
sele 2
use temp1
scan
sele 1
use gongzong
 loca for allt(name)=allt(b.gz1) and allt(gwname)=allt(b.gw) 
 WAIT WINDOW NOWAIT '����ά��'+gwdm+'��'+allt(name)+'��'+allt(gwname)+'<���ִ����>... ...' 
 repl gwdm with b.gwjb
 SET RELATION TO gwdm INTO 3
 repl gwdj with c.gd01,bzgz WITH c.bzgz 
 ****��λ��λ�ȼ�����ά��**201719****
ENDSCAN
ENDIF
CLOSE all
use gongzong EXCLUSIVE
go top
brow fiel code:h='���ִ���',name:h='����',gwname:h='��λ',gwdm:h='�ڵȴ���',gwdj:h='��λ�ȼ�',bzgz:h='��׼����' titl '�Զ���淶��������λ��ǰ���ָ�λ���ƺʹ��루����󽫰��������ݿ⹤�ִ����Զ����¸�λ��ҵ��Ա���ִ��룩'
************����Աרҵά�����루����λ���ָ�λʵ�����ͳһά����****20160829******
pack
repl a with 1 all
sort on code to temp
close all
use temp excl
go top
do while !eof()
   bh=code
   skip
   if bh=code
      repl a with 0
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a=0 to aaa
if aaa>0
   inde on code to temp
   go top
   brow for a=0 titl '����'+allt(str(aaa,4))+'�����ִ����غţ����޸���ȷ��ɾ�������¼' 
endif
repl a with 1 all
sort on code to data\gongzong
CLOSE all
yn = MESSAGEBOX("�ù��ָ�λ���ݿ��û�����λԱ����ǰ���ִ�����",4+32,"��ʾ")
IF yn = 6
CLOSE all
sele 1
use ryzk
scan
 sele 2
 use gongzong
 loca for allt(name)=allt(A.gz1) and allt(gwname)=allt(A.gw) 
 WAIT WINDOW NOWAIT '����ά��'+A.cjdm+'��'+allt(A.cjmc)+'��'+allt(A.bmmc)+'��'+allt(A.ryxm)+'<���ִ���>... ...' 
 repl A.gz with code
 ****��λ���ִ���ά��**20160830****
ENDSCAN
SELECT  1
go top
brow fiel dwmc:h='��λ',cjmc:h='����',bmmc:h='����',ryxm:h='����',gz:h='���ִ���',gz1:h='����',gw:h='��λ',bzgz:h='��׼����',gwgz:h='��λ����'titl '�û��󱾵�λԱ����ǰ���ָ�λ'
close all
ENDIF

