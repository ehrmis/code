*******************************
*��λ����.prg (Build 20170719)
*******************************
set escape off
dhgw=''
dh14=''
dh1a=''
dhgwdy=''
on key label F1 do grcx
YF1=month(date())
ND1=year(date())
ND = str(ND1,4)
if YF1>9
  YF = str(YF1,2)
else
  YF = '0'+str(YF1,1)
endif
NY = ND+YF
********��ǰϵͳ����*************
NYR = Right(NY,4)
YFs = iif(val(yf)>9,str(val(yf)-1,2),'0'+str(val(yf)-1,1))
IF val(yf)=1
   NY1 = str(year(date())-1,4)+'12'
   YFs = '12'
else   
   if val(yf)=10
      NY1 = str(year(date()),4)+'09'
   ELSE
      NY1 = str(year(date()),4)+YFs
    endif   
ENDIF
********��ǰϵͳ����[���±��ʽ]*************
wait window '��  ��  ��  �ݡ���, ��  ��  ��... ...' nowait
CLOSE all
USE ryzk
*********�����ǵ�ǰryzk��ry&ny�޹�*****20170320**************************
COPY TO ryzktemp
USE ryzktemp EXCLUSIVE
COUNT for year(���ʱ��)=<year(date()) and year(���ʱ��)>0 TO lgrs
IF lgrs>0 
DELETE for year(���ʱ��)=<year(date()) and year(���ʱ��)>0
BROWSE FIELDS dwmc,cjmc,bmmc,ryxm,���ʱ�� for year(���ʱ��)=<year(date()) and year(���ʱ��)>0 TITLE '������Ƿ�ɾ�������Ա��'
*****************************Ӧɾ����Ա*************************************** 
yn = MESSAGEBOX("�Ƿ�ɾ�������Ա��",4+32,"��ʾ")
IF yn = 6
PACK
ELSE
RECALL
ENDIF
ENDIF
******��ʱ����Ԥɾ����Ա**************
close all
use ��λ���� excl
zap
appe from ryzktemp
repl �ڵ� with zcdj1 for allt(gwfl1)='������λ' 
repl gwgz with bzgz for allt(ygxz1)='��ͬ��ְ��' 
repl ���� with 1 , A with 1 all
IF !files("gw"+ny1+".dbf")
   COPY to gw&NY1
ENDIF  
COPY to gw&NY
close all
wait window '��  ��  ��  �� , ��  ��  ��... ...' nowait
sele 1
use ��λ����
repl �ڵ� with '' all
go top
sele 2
use gwdmk
SUM �Զ��� TO zdy
scan
 sele 1 
 repl gwdm with b->gwdm for allt(gwfl1)='������λ' and gwjb = b->gwdm
 repl gwdm with b->gwdm for allt(gwfl1)='������λ' and gwjb = b->gwdm
 repl gwdm with b->gwdm for allt(gwfl1)='�����λ' and gwjb = b->gwdm
 IF zdy=0
 ******��λ�Զ��幤������ֵ��****20170322**************
 repl �ڵ� with b->gd03 for allt(gwfl1)='������λ' and bzgz = b->bzgz2014
 repl �ڵ� with b->gd04 for allt(gwfl1)='�����λ' and bzgz = b->bzgz2014
 ************************************��λ�Զ��幤�ʱ�׼******20170322********************************
 ELSE 
  repl �ڵ� with zcdj1 for allt(gwfl1)='������λ'
  repl �ڵ� with zw1 for allt(gwfl1)='�����λ'
 ************************************���Ź��ʱ�׼******20170322********************************
 ENDIF 
ENDSCAN
********************�ڵȣ���λ�Զ����뼯�Ź����ƶȵı�׼ְ�ƹ���Ҫ���ֿ���********20170112***************
close all
********************��bzgz2014�汾У���ڵ�*******20170322****************
use ��λ����
repl �ڵ� with 'ҵ��Ա' for allt(gwfl1)='�����λ' and gwdm = '06'
repl �ڵ� with 'ҵ������' for allt(gwfl1)='�����λ' and gwdm = '07'
repl �ڵ� with 'ҵ������' for allt(gwfl1)='�����λ' and gwdm = '08'
repl �ڵ� with 'ҵ�񸱾���' for allt(gwfl1)='�����λ' and gwdm='09' and zw='07'
repl �ڵ� with 'ҵ����' for allt(gwfl1)='�����λ' and gwdm='11' and zw='06'
repl �ڵ� with '�߼�ҵ����' for allt(gwfl1)='�����λ' and gwdm='13' 
repl �ڵ� with '���ܾ���' for allt(gwfl1)='�����λ' and gwdm='15' and zw='03'
repl �ڵ� with '�ܾ���' for allt(gwfl1)='�����λ' and gwdm='17' and zw='02'
***********************��.���ά��*****************
inde on bmbh+zw+rybh to ��λ����
for zdz=1 to fcount()
  zd=fiel(zdz)
  zdz1=(fsize(zd)+fsize(zd))*2
  if uppe(allt(zd))='RYXM' 
     exit
  endif
endfor   
go top
browse part zdz1 field dwmc:h='��λ':r,cjmc:h='����':r,bmmc:h='����':r,ryxm:h='����':r,zw1:r:h='����ְ��',bzf:h='ע��೤(Y)',gz:h='���ִ���',;
gz1:h='����':r,gw:h='��λ':r,gwjb1:r:h='�ȼ�',gwdc1:h='��λ����':r,gwgz:h='��λ����':r,�ڵ�:h='��Ƹ��λ�ȼ�',sfgj:h='�ϸ��ڼ�':r,;
bzgz:h='��׼����':r,gwdm:h='�ڵȴ���',gl:h='����':r,pxdj:h='��ѵ�ȼ�':r,zcdj1:r:h='����ְ��',zc:r:h='ְ��',gwfl:h='��λ������':r,;
gwfl1:h='��λ����':r,zymc:h='ְҵ��չ' titl '[�� F1 �����] [�� Esc �ر�] ��������˺�ά������ע��೤�ڹ��ִ������Ƹ��λ�ȼ�'
repl gwfl1 with '�����λ' for gwfl='01'
repl gwfl1 with '������λ' for gwfl='02'
repl gwfl1 with '������λ' for gwfl='03'
CLOSE all
do ���
do ��ѵ
close all
sele 2
use ��λ����
inde on rybh to ��λ����
sele 1
use ryzk &&�����������
inde on rybh to ryzk
upda on rybh from B repl �ڵ� with b.�ڵ�,bzf with b.bzf,gz with b.gz
**************************��.�������***********************�������¹��ָ�λ����20170320*****************************
close all
use ��λ����
inde on gwjb+gwdc to ��λ����
tota on gwjb+gwdc to �ڼ�ͳ��
copy to gw&NY
use gw&NY
repl ���� with 1 , A with 1 all
inde on gz to gw&NY
tota on gz to temp
copy to ������Ա for val(gwfl)<3
******************************ԭ������Ա(2003��6��)**************
copy to ������Ա for val(gwfl)=3
use ������Ա excl
copy to ��֤��Ա for empty(pxdj) and '����'$zcdj1
delete for empty(pxdj) and gwfl='03'
pack
inde on gwfl1+�ڵ�+bzf to ������Ա
tota on gwfl1+�ڵ�+bzf to ��������
use ��֤��Ա excl
inde on gwfl1+�ڵ�+bzf to ��֤��Ա
tota on gwfl1+�ڵ�+bzf to ��֤����
use ������Ա
inde on gwfl1+�ڵ�+�ڵ� to ������Ա
tota on gwfl1+�ڵ�+�ڵ� to �������
close all
yn = MESSAGEBOX("��λ�䶯����λ�������Ʊ������ɹ����ִ�ʹ����",4+32,"��ʾ")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\��λ�䶯.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\��׼��������.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\ƽ����������.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
ENDIF

proc ��ѵ

*****��ÿ�˵���߼������򱣴��ͳ����ѵ���****
use ryzk
repl a with 1 all
copy to temp for !empty(pxdj) fiel cjmc,bmmc,rybh,ryxm,gz1,gw,pxmc,pxdj,zcfl1,zcdj1,gwfl1,a
use temp
copy to temp1 fiel gwfl1,gz1,gw,pxmc,pxdj,a
use temp1
inde on gwfl1+gz1+gw+pxmc+pxdj to temp1
tota on gwfl1+gz1+gw+pxmc+pxdj to ��ѵͳ��
close all

proc ���
close all
select 1
use gw&NY
repl a with 1 all
select 2
use gw&NY1
scan
wait window '�� �� �� �� �� �� �� λ �� ׼,�� �� ��... ...' nowait
  select 1
  repl a with 0 for allt(rybh)=allt(b.rybh) and bzgz<b.bzgz &&����
  repl a with 2 for allt(rybh)=allt(b.rybh) and bzgz>b.bzgz &&����
endscan
close all
***************************
use gongzong EXCLUSIVE
yn = MESSAGEBOX("��Ҫά�����ִ������",4+32,"��ʾ")
IF yn = 6
   APPEND BLANK
go top
brow fiel code:h='���ִ���',name:h='����',gwname:h='��λ',gwdj:h='��λ�ȼ�':r,bzgz:h='��׼����':r,pj:h='ƽ������':r,num:h='�ڸ�����':r titl '��ǰ���ָ�λ�Զ�����뼰���׼����������󽫰���׼�����Զ����¸�λ��ҵ��Ա���ָ�λ���ƣ�'
************����Աרҵά�����루����λ���ָ�λʵ�����ͳһά������ֻ�����ԣ���֧���޸Ĵ���****20160829******
DELETE FOR EMPTY(code)
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
COPY TO data\gongzong
************�������޸Ĺ��ִ����ʱ���򱣴�****20170322*******************
ENDIF
CLOSE all
USE ryzk
REPLACE zyflm WITH  '',zymc with '' all 
use zydm EXCLUSIVE
go top
brow fiel ����,ְҵ,����,��λ,ְ�� titl '�����ά����ǰְҵ����й��ָ�λ��ְҵ��չ���ݿ�'
************����Աרҵά�����루����λ���ָ�λʵ�����ͳһά������ֻ�����ԣ���֧���޸Ĵ���****20160829******
pack
repl a with 1 all
sort on ���� to temp1
close all
use temp1 excl
go top
do while !eof()
   bh=����
   skip
   if bh=����
      repl a with 0
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a=0 to aaa
if aaa>0
   inde on ���� to temp1
   go top
   brow for a=0 titl '����'+allt(str(aaa,4))+'�������غţ����޸���ȷ��ɾ�������¼' 
endif
repl a with 1 all
sort on ���� to data\zydm
CLOSE all
sele 1
use ryzk
scan
 sele 2
 use zydm
 loca for allt(����)$allt(A.gz1) or allt(��λ)$allt(A.gw) or allt(ְ��)$allt(A.zc)
 WAIT WINDOW NOWAIT '����ά��'+A.cjdm+'��'+allt(A.cjmc)+'��'+allt(A.bmmc)+'��'+allt(A.ryxm)+'<ְҵ��չ>... ...' 
 repl A.zyflm with allt(����),A.zymc with allt(ְҵ)
 ***************�滮ְҵ����*********20160830******************************
 use gongzong
 loca for allt(code)=allt(A.gz)
 WAIT WINDOW NOWAIT '����ά��'+A.cjdm+'��'+allt(A.cjmc)+'��'+allt(A.bmmc)+'��'+allt(A.ryxm)+'<���֡���λ>... ...' 
 repl A.gz1 with allt(name),A.gw with allt(gwname)
 ***************1.��׼���֡���λ����ά��2.�������޸Ĺ��ִ���Ĺ��֡���λ�����Զ����³ɹ��ִ����Ӧ�Ĺ��ָ�λ**************20170320************
endscan
close all
use gw&NY
REPLACE  a WITH 1 all
INDEX on gz TO gw&NY
tota on gz TO gztemp
SELECT 3
USE gwdmk
INDEX on gwdm TO gwdmk
select 1
use gongzong
REPLACE gwdm WITH '',gwdj WITH '',pj WITH 0,bzgz WITH 0,num WITH 0 all
**********���ټ�¼������¶��¼����ʱ��������ձ����±�*********20170322**********
select 2
use gztemp EXCLUSIVE
DELETE FOR EMPTY(gz)
PACK
scan
  WAIT WINDOW NOWAIT '����ͳ�ƹ��ָ�λ����... ...'
  select 1
  LOCATE for code=b.gz
  repl gwdm with b.gwjb, pj with b.bzgz/b.a,num with b.a 
  ******************************************��Ŀǰƽ����λ�����ж϶����ʺ϶Ȼ�������Ʒ���******20170719*******
 WAIT WINDOW NOWAIT '����ά��'+gwdm+'��'+allt(name)+'��'+allt(gwname)+'<���ִ����>... ...' 
 repl gwdm with b.gwjb
 SET RELATION TO gwdm INTO 3
 repl gwdj with c.gd01,bzgz WITH c.bzgz 
 ****��λ��λ�ȼ�����ά��**201719****
ENDSCAN
****************************
close all
select 1
use gw&NY
repl a with 1 all
inde on gwjb to gw&NY
tota on gwjb to gwjbtemp fiel gwjb,a
use gwjbtemp
select 2
use gwdmk
scan
  WAIT WINDOW NOWAIT '����ͳ�Ƹ�λ����... ...'
  select 1
  repl b.num with a for gwjb=b.gwdm
endscan
close all
use gw&NY
repl a with 1 all
copy to ��������
select 1
use �������� excl
delete for bzgz=0
pack
select 2
use gongzong
scan
wait window '���ڼ������ָ�λ���λ��׼���������,���Ժ�... ...' nowait
  select 1
  LOCATE for allt(gz)=allt(b.code) and bzgz<b.bzgz
  *****************BUGɨ�����б�����LOCATE for ������for *********20170719********************************
  repl a with 0
   LOCATE for allt(gz)=allt(b.code) and bzgz=b.bzgz
  *****************BUGɨ�����б�����LOCATE for ������for *********20170719********************************
  repl a with 1
  LOCATE for allt(gz)=allt(b.code) and bzgz>b.bzgz 
  repl a with 2 
  LOCATE for allt(gz)=allt(b.code) and bzgz<b.pj
  *****************BUGɨ�����б�����LOCATE for ������for *********20170719********************************
  repl a with 3
  LOCATE for allt(gz)=allt(b.code) and bzgz=b.pj
  repl a with 4 
   LOCATE for allt(gz)=allt(b.code) and bzgz>b.pj
  *****************BUGɨ�����б�����LOCATE for ������for *********20170719********************************
  repl a with 5
  repl gwgz with b.pj for allt(gz)=allt(b.code)
endscan
close all
use �������� excl
coun to xy for a=0
coun to dy for a=2
coun to xy1 for a=3
coun to dy1 for a=5
IF xy+dy+xy1+dy1>0
REPLACE gz WITH '����' FOR a=2 
REPLACE gz WITH '����' FOR a=1 
REPLACE gz WITH 'С��' FOR a=0 
go top
 brow part 35 for a=0 or a=2 fiel  cjmc:r:h='����',bmmc:r:h='����',ryxm:r:h='����',gz:r:h='����',gz1:r:h='����',gw:r:h='��λ',gwjb1:h='����':r,;
�ڵ�:r,sfgj:h='�ϸ��ڼ�':r,gwgz:r:h='ƽ����λ����',bzgz:r:h='��׼����',zymc:h='ְҵ��չ' titl 'ְ����λ���ʴ��ڱ����ָ�λ��׼���ʣ�'+str(dy,3)+'�ˣ�ְ����λ����С�ڱ����ָ�λ��׼���ʣ�'+str(xy,3)+'�� ��'+str(dy+xy,4)+'��'
copy to &bf.\��׼�������� FIELDS rybh,ryxm,cjmc,bmmc,�ڵ�,gwgz,gz,gz1,gw,gwjb1,bzgz type xl5
REPLACE gz WITH '����' FOR  a=5
REPLACE gz WITH '����' FOR  a=4
REPLACE gz WITH 'С��' FOR  a=3
 GO top
 brow part 35 for a=3 or a=5 fiel cjmc:r:h='����',bmmc:r:h='����',ryxm:r:h='����',gz:r:h='����',gz1:r:h='����',gw:r:h='��λ',gwjb1:h='����':r,;
�ڵ�:r,sfgj:h='�ϸ��ڼ�':r,gwgz:r:h='ƽ����λ����',bzgz:r:h='��׼����',zymc:h='ְҵ��չ' titl '��λ�������Ʒ�����ְ����λ���ʴ��ڱ����ָ�λƽ�����ʣ�'+str(dy1,3)+'�ˣ�ְ����λ����С�ڱ����ָ�λƽ�����ʣ�'+str(xy1,3)+'�� ��'+str(dy1+xy1,4)+'��'
copy to &bf.\ƽ���������� FIELDS rybh,ryxm,cjmc,bmmc,�ڵ�,gwgz,gz,gz1,gw,gwjb1,bzgz type xl5
ENDIF
**************��λ�������Ʒ���****20170719**********************
close all
use gw&NY
repl a with 1 all
INDEX on bmbh+rybh TO gw&NY
brow part 35 fiel cjmc:r:h='����',bmmc:r:h='����',ryxm:r:h='����',gz1:r:h='����',gw:r:h='��λ',gwjb1:h='����':r,;
�ڵ�:r,gwdc1:h='��λ����':r,gwgz:r:h='��λ����',zymc:h='ְҵ��չ' titl 'Ա��ְҵ���Ĺ滮��˵���������𡱡����ڵȡ�������׼���ʡ���ָ���ְҵ��չĿ�꣩'
copy to temp 
jg=0
sg=0
 sele 1
   use temp
 sele 2
   use gw&NY1
   inde on rybh to gw&NY1  
scan
wait window '�� �� �� �� �� �� �� λ �� �� �� ��,�� �� ��... ...' nowait
  select 1
  repl a with 0 for allt(rybh)=allt(b.rybh) and bzgz<b.bzgz
  repl a with 2 for allt(rybh)=allt(b.rybh) and bzgz>b.bzgz
endscan
close all
use temp excl
delete for a=1
pack
if recc()>0
�˿�='�䶯'
sort on bmbh,rybh to temp1 
   close all
   sele 2
   use gw&NY1
   inde on rybh to gw&NY1
   sele 1 
   use temp1 excl
   repl ��� with str(recn(),4) all
   count for a=0 to jg
   count for a=2 to sg
   inde on rybh to temp1
   upda on rybh from B repl dwmc with b.gz1,bmmc with b.gw,�ڵ� with b.gwjb1,gwgz with b.bzgz
   inde on ��� to temp1
   go top
   brow fiel ���,ryxm:h='����',erprybh:h='����',cjmc:h='��λ',dwmc:h='ԭ����',bmmc:h='ԭ��λ',�ڵ�:h='ԭ�ڼ�',gwgz:h='ԭ��λ����',gz1:h='�ֹ���',gw:h='�ָ�λ',gwjb1:h='�ָڼ�',bzgz:h='�ָ�λ����' titl '������Ա��λ�����б䶯(��'+str(recc(),4)+'��;����'+str(sg,4)+'��,����'+str(jg,4)+'��)'
   copy to &bf.\��λ�䶯 FIELDS ���,ryxm,erprybh,cjmc,dwmc,bmmc,�ڵ�,gwgz,gz1,gw,gwjb1,bzgz type xl5
ENDIF
close all

proc grcx
do while .t.
      XM = '        '
      do xmsr.spr
      locate for upper(XM)=RYXM or upper(XM)=RYBH
      if eof()
        do dhkjg.spx
    M = readkey()
    if M=12 or M=268
      return
    endif
      else
        exit
      endif
enddo