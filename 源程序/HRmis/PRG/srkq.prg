***************************
* .\SRKQ.PRG(Build20161206)
***************************
on key label F1 do kqgrcx
on key label F2 do sykqts
on key label F3 do plsr
clear
_SCREEN.PICTURE = ''
dh=''
XH1 = 'N'
LY = ''
T123 = '  '
close all
USE dmk
�ֳ���=�Ʒ�����
use deset
GO top
LOCATE FOR seted=0 AND 'sfgz'$oop
IF !EOF()
   sfgz1='����'   
ELSE
sfgz1='����'
USE dmk
��׼=��Զ����
����=�Ʒ�����
�ձ�=round(��׼/����,2)
����1=allt(�Ʒ�����)
����2=allt(�۷�����)
����3=allt(��������)
ENDIF
set confirm off
LY = ND
****************ǰ���Զ�����****************
if val(yf)<3
  JJND = str(val(nd)-1,4)
  JJYF = str(10+val(yf),2)
else
  JJND = ND
  JJYF = iif(val(YF)-2>9,str(val(YF)-2,2),'0'+str(val(YF)-2,1))
endif
****************�����Զ�����****************
if val(YF)-1>9
  jjYF = str(val(YF)-1,2)
else
  jjYF = iif(val(YF)-1=0,'12','0'+str(val(YF)-1,1))
endif
if jjyf='12'
   syjj=str(val(nd)-1,4)+jjYF
else
   syjj = ND+jjYF
endif
****************��ʼ��*******20161020*********
if i=99
do case
   case VAL(dwdm1)=0 and VAL(cjdm1)>0 and val(bmbh1)=0 
   dh=2
   ******��������******
   T123 = cjdm1
   case VAL(dwdm1)>=0 and VAL(cjdm1)>0 and val(bmbh1)>0
   dh=3 
   *************��������**********   
   T123 = bmbh1
   **********��ȷɸѡ������****************
   case VAL(dwdm1)>0 and VAL(cjdm1)>0 and val(bmbh1)=0 
   dh=2
   T123 = cjdm1
   *************ģ����λ����ѡ��������**********
   case VAL(dwdm1)>0 and VAL(cjdm1)=0 and val(bmbh1)>0 
   dh=3 
   T123 = bmbh1
   *************ģ����λ����ѡ��������********** 
   OTHERWISE 
   T123 = '  '
   dh=1
   *************ģ������**********
endcase
ELSE
T123 = '  '
dh=1
ENDIF
close all
clear
XH1 = 'N'
if !file('KQ'+ny+'.DBF')
  use KQK excl
  ZAP
 IF FILE('ry'+ny+'.dbf')
  append from ry&ny for ygxz='01' and !'��'$zgqk1
  ********************�����������ڲ�ְ��****20151217****************
 ELSE
  append from ryzk for ygxz='01' and !'��'$zgqk1
 ENDIF 
 SORT ON bmbh,zw,rybh  TO KQ&ny
 *********ֱ���������ɵ��¿���*****20161026********************
endif
close all
*********************��ʼ��****************
************������ʼ��ֱ��ȫ��׷������****20161205******************************
select 2
 IF FILE('ry'+ny+'.dbf')
    use ry&ny 
    repl a with 1 all
    index on RYBH to ry&ny
 ELSE
    use ryzk 
    repl a with 1 all
    index on RYBH to ryzk
 ENDIF
select 1
use kq&ny EXCLUSIVE
repl ���ݼƻ� with CTOD(''),a with 0 all
************������ʼ��ֱ��ȫ��׷������*****��鱾��Ӧɾ��/������Ա��ʱ����***20160613*********
inde on rybh to kq&ny
update on RYBH from B replace CJDM with b.CJDM , BMBH;
 with b.BMBH , CJMC with b.CJMC , BMMC with b.BMMC ,erprybh with b.erprybh,bjgz with b.bjgz,bjdj with b.bjdj,bjjb with b.bjjb,;
 zbjt with b.zbjt,ybjt with b.ybjt,gz1 with b.gz1,zw with b.zw,��ҵ���� with b.��ҵ����,zgqk with b.zgqk,a with b.a
GO top
loca for a=0 
IF !eof()
   delete for a=0 
   brow for a=0 titl '���������������Ա�Ƿ�ɾ��'
    yn = MESSAGEBOX("��Ҫɾ�����������Ա��",4+32,"��ʾ")
   IF yn = 6
      PACK
****���ܣ�����ʾɾ���ٹ���������Ա״����ȷ���Ƿ�����������Ա��******20160613*************
   ELSE
       RECALL all
   ENDIF         
ENDIF
CLOSE ALL
select 2
use kq&ny
repl a with 0 all
index on RYBH to kq&ny
select 1
IF FILE('ry'+ny+'.dbf')
    use ry&ny 
    repl a with 1 all
    index on RYBH to ry&ny
 ELSE
    use ryzk 
    repl a with 1 all
    index on RYBH to ryzk
 ENDIF 
update on RYBH from B replace a with b.a
GO top
loca for a=1 and ygxz='01' 
if !eof()
   brow for a=1 and ygxz='01' titl '��������������������뿼�ڿ�����Ա'
   CLOSE all
   use kq&ny
   IF FILE('ry'+ny+'.dbf')
    APPEND FROM ry&ny for a=1 and ygxz='01'
   ELSE
    APPEND FROM ryzk for a=1 and ygxz='01'
   ENDIF  
ENDIF
CLOSE all
use kq&ny 
repl a with 1 all
 IF FILE('ry'+ny+'.dbf')
    use ry&ny 
    repl a with 1 all
 ELSE
    use ryzk 
    repl a with 1 all
 ENDIF 
**************���뻹ԭĬ�ϱ�ʶ*****20160613************************      
close all
IF files('kq'+syjj+'.dbf')
select 4
use kQ&syJJ
index on RYBH to temp
ENDIF 
select 3
IF FILE('ry'+ny+'.dbf')
    use ry&ny 
    index on RYBH to ry&ny
 ELSE
    use ryzk 
    index on RYBH to ryzk
 ENDIF 
select 1
use SMK1
select 2
use kQ&ny
set relation to RYBH into 3
IF files('kq'+syjj+'.dbf')
set relation to RYBH into 4 additive
ENDIF 
index on BMBH+zw+RYBH to kq&ny
if !empty(T123)
  go bottom
do case
   case dh=2
   set filter to cjdm=T123
   case dh=3
   set filter to bmbh=T123
endcase
  go top
  NUM1 = recno()
  go bottom
  NUM2 = recno()
  go top
else
  go top
  NUM1 = recno()
  go bottom
  NUM2 = recno()
  go top
endif
XH1 = 'Y'
DEFINE WINDOW temp FROM INT((SROWS()-40)/2)  , INT((SCOLS()-160)/2)  TO INT((SROWS()+40)/2) , INT((SCOLS()+160)/2) title 'DOSʱ������С���̿��ٿ�������ģʽ [ ʹ��С���̿������룬�ûس���ȷ�ϱ������� ]'
activate window temp
do while upper(XH1)='Y'
  XH2 = 'N'
  select 1
  go top
  KKK1 = 5
  LLL1 = 3
  do while  not eof()
    @ KKK1 , 20+LLL1 say CHINES font '����' , 14
    KKK1 = KKK1+2
    if KKK1>20
      KKK1 = 5
      LLL1 = LLL1+35
    endif
    skip
  enddo
  @ 23 , 33 say 'PgDn = ��һ����PgUp = ��һ��  Enter = ����' font '����' , 15
  @ 25 , 33 say 'F1 = ���Ҹ��ˡ�F2 = ���¿��ڡ�F3 = �������롡Esc = �˳�' font '����' , 15
  do while upper(XH2)='N'
    select 1
    go top
    KKK1 = 5
    LLL1 = 3
    do while  not eof()
      RYHDS = FIELD_NAME
      TM123 = HHH1
      select 2
      if upper(trim(RYHDS))='RYBH' or upper(trim(RYHDS))='RYXM' or upper(trim(RYHDS))='BMMC';
 or upper(trim(RYHDS))='CJMC'
        RYHDS=&RYHDS
        @ KKK1 , 35+LLL1 say RYHDS font '����' , 14
      else
        @kkk1,35+lll1 get &ryhds font '����' , 14 
        *PICT &TM123
      endif
      KKK1 = KKK1+2
      if KKK1>20
        KKK1 = 5
        LLL1 = LLL1+35
      endif
      select 1
      skip
    enddo
    read
    M = readkey()
    select 2  
    do case
    case M=6 or M=262
      if recno()<>NUM1
        skip -1
      endif
      loop
    case M=7 or M=263
      if recno()=NUM2
        go bottom
      else
        skip
      endif
      loop
    case M=2 or M=258
      go top
      loop
    case M=12 or M=268
      exit
    endcase
  enddo
  FF = readkey()
  if FF=12 or FF=268
    exit
  endif
ENDDO
replace ���ݼƻ� with c.���ݼƻ� FOR YEAR(c.���ݼƻ�)=YEAR(date()) AND YEAR(��ֹʱ��)=YEAR(date()) AND rybh=c.rybh
*********************�Զ���д�����ݼ���Ա�ĵ������ݼƻ�****************
CLEAR WINDOWS temp
********************1.�������ݺ͸��Դ�������***********************
on key label F2
on key label F3
************�ͷ��Զ����ȼ�****20161128**************
select 2
   rs=recc()
   rs=allt(str(rs))
   sum fzsd,kk to sfzsd , skk
   count for bj>0 to bjj
   bjj=allt(str(bjj))
   count for sj>0 to sjj
   sjj=allt(str(sjj))
if dh<>1
   if dh=2
   count for left(BMBH,2)=T123  to rs1
   rs1=allt(str(rs1))
   sum fzsd,kk to sfzsd1,skk1 for left(BMBH,2)=T123 
   count for left(BMBH,2)=T123 and bj>0 to bjj1
   bjj1=allt(str(bjj1))
   count for left(BMBH,2)=T123 and sj>0 to sjj1
   sjj1=allt(str(sjj1))
   go top
   loca for left(BMBH,2)=T123
   cjm=allt(cjmc)
   IF files('kq'+syjj+'.dbf')
   brow part 20 when sykqts() fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',sbts:h='��������',xcts:h='�ֳ�����',zbgr:h='�а�',ybgr:h='ҹ��',gj:h='���ݼ�',���ݼƻ�,��ʼʱ��,��ֹʱ��,bj1:h='����',jrjb:h='���ռӰ�',jjb:h='���ռӰ�',xxpx:h='������ѵ',bj:h='����',sj:h='�¼�',cj:h='�㻤/����',gs:h='����',tqj:h='̽��',hunj:h='���',sangj:h='ɥ��',gx:h='����',ly:h='����',jl:h='����',kg:h='����',fzsd:h='����ˮ��',kk:h='�ۿ�';
    titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs1+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'�ˣ����У�'+cjm+'��'+rs1+'�ˣ�����='+bjj1+'�ˣ��¼�='+sjj1+'�ˣ�����ˮ�磽'+allt(str(sfzsd1,10,1))+'Ԫ���ۿ'+allt(str(skk1,10,1))+'Ԫ   ��[ Esc ]���˳�'
   ELSE
   brow part 20 fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',sbts:h='��������',xcts:h='�ֳ�����',zbgr:h='�а�',ybgr:h='ҹ��',gj:h='���ݼ�',gj:h='���ݼ�',���ݼƻ�,��ʼʱ��,��ֹʱ��,bj1:h='����',jrjb:h='���ռӰ�',jjb:h='���ռӰ�',xxpx:h='������ѵ',bj:h='����',sj:h='�¼�',cj:h='�㻤/����',gs:h='����',tqj:h='̽��',hunj:h='���',sangj:h='ɥ��',gx:h='����',ly:h='����',jl:h='����',kg:h='����',fzsd:h='����ˮ��',kk:h='�ۿ�';
    titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs1+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'�ˣ����У�'+cjm+'��'+rs1+'�ˣ�����='+bjj1+'�ˣ��¼�='+sjj1+'�ˣ�����ˮ�磽'+allt(str(sfzsd1,10,1))+'Ԫ���ۿ'+allt(str(skk1,10,1))+'Ԫ   ��[ Esc ]���˳�'
   ENDIF
  else
   count for BMBH=T123  to rs1
   rs1=allt(str(rs1))
   sum fzsd,kk to sfzsd2,skk2 for BMBH=T123 
   count for BMBH=T123 and bj>0 to bjj1
   bjj1=allt(str(bjj1))
   count for BMBH=T123 and sj>0 to sjj1
   sjj1=allt(str(sjj1))
   go top
   loca for BMBH=T123
   bm=allt(bmmc)
   IF files('kq'+syjj+'.dbf')
  brow part 20 when sykqts() fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',sbts:h='��������',xcts:h='�ֳ�����',zbgr:h='�а�',ybgr:h='ҹ��',gj:h='���ݼ�',���ݼƻ�,��ʼʱ��,��ֹʱ��,bj1:h='����',jrjb:h='���ռӰ�',jjb:h='���ռӰ�',xxpx:h='������ѵ',bj:h='����',sj:h='�¼�',cj:h='�㻤/����',gs:h='����',tqj:h='̽��',hunj:h='���',sangj:h='ɥ��',gx:h='����',ly:h='����',jl:h='����',kg:h='����',fzsd:h='����ˮ��',kk:h='�ۿ�';
   titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs1+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'�ˣ����У�'+bm+'��'+rs1+'�ˣ�����='+bjj1+'�ˣ��¼�='+sjj1+'�ˣ�����ˮ�磽'+allt(str(sfzsd2,10,1))+'Ԫ���ۿ'+allt(str(skk2,10,1))+'Ԫ  ��[ Esc ]���˳�'
   ELSE
   brow part 20 fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',sbts:h='��������',xcts:h='�ֳ�����',zbgr:h='�а�',ybgr:h='ҹ��',gj:h='���ݼ�',���ݼƻ�,��ʼʱ��,��ֹʱ��,bj1:h='����',jrjb:h='���ռӰ�',jjb:h='���ռӰ�',xxpx:h='������ѵ',bj:h='����',sj:h='�¼�',cj:h='�㻤/����',gs:h='����',tqj:h='̽��',hunj:h='���',sangj:h='ɥ��',gx:h='����',ly:h='����',jl:h='����',kg:h='����',fzsd:h='����ˮ��',kk:h='�ۿ�';
   titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs1+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'�ˣ����У�'+bm+'��'+rs1+'�ˣ�����='+bjj1+'�ˣ��¼�='+sjj1+'�ˣ�����ˮ�磽'+allt(str(sfzsd2,10,1))+'Ԫ���ۿ'+allt(str(skk2,10,1))+'Ԫ  ��[ Esc ]���˳�'
   ENDIF
   endif
else
   go top
  IF files('kq'+syjj+'.dbf')  
  brow part 20 when sykqts() fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',sbts:h='��������',xcts:h='�ֳ�����',zbgr:h='�а�',ybgr:h='ҹ��',gj:h='���ݼ�',���ݼƻ�,��ʼʱ��,��ֹʱ��,bj1:h='����',jrjb:h='���ռӰ�',jjb:h='���ռӰ�',xxpx:h='������ѵ',bj:h='����',sj:h='�¼�',cj:h='�㻤/����',gs:h='����',tqj:h='̽��',hunj:h='���',sangj:h='ɥ��',gx:h='����',ly:h='����',jl:h='����',kg:h='����',fzsd:h='����ˮ��',kk:h='�ۿ�';
   titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'�ˣ�����ˮ�磽'+allt(str(sfzsd,10,1))+'Ԫ���ۿ'+allt(str(skk,10,1))+'Ԫ  ��[ Esc ]���˳�'
  ELSE
  brow part 20 fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',sbts:h='��������',xcts:h='�ֳ�����',zbgr:h='�а�',ybgr:h='ҹ��',gj:h='���ݼ�',���ݼƻ�,��ʼʱ��,��ֹʱ��,bj1:h='����',jrjb:h='���ռӰ�',jjb:h='���ռӰ�',xxpx:h='������ѵ',bj:h='����',sj:h='�¼�',cj:h='�㻤/����',gs:h='����',tqj:h='̽��',hunj:h='���',sangj:h='ɥ��',gx:h='����',ly:h='����',jl:h='����',kg:h='����',fzsd:h='����ˮ��',kk:h='�ۿ�';
   titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'�ˣ�����ˮ�磽'+allt(str(sfzsd,10,1))+'Ԫ���ۿ'+allt(str(skk,10,1))+'Ԫ  ��[ Esc ]���˳�'
  ENDIF
ENDIF
index on BMBH+zw+RYBH to kq&ny
REPLACE xcts WITH 0,sbts WITH 0 FOR zgqk>'02'
*************���ڸ���Ա��������*****20170726*****************
count for sbts=0 OR xcts>25 or xcts<�ֳ��� or xcts<>sbts or sj>7 TO rs1
**********�����쳣��ʾ****20160902*************** 
if rs1>0
 IF files('kq'+syjj+'.dbf')
   brow part 20 when sykqts() field cjmc:h='��λ',bmmc:h='����',ryxm:h='����',sbts:h='��������',xcts:h='�ֳ�����',bj:h='����',sj:h='�¼�',cj:h='�㻤/����',gs:h='����',tqj:h='̽��',gx:h='����',jl:h='����' for sbts=0 OR xcts>25 or xcts<�ֳ��� or xcts<>sbts titl '��ȷ��������Ա�ֳ�����'
 ELSE
   brow part 20 field cjmc:h='��λ',bmmc:h='����',ryxm:h='����',sbts:h='��������',xcts:h='�ֳ�����',bj:h='����',sj:h='�¼�',cj:h='�㻤/����',gs:h='����',tqj:h='̽��',gx:h='����',jl:h='����' for sbts=0 OR xcts>25 or xcts<�ֳ��� or xcts<>sbts titl '��ȷ��������Ա�ֳ�����'
 ENDIF
   REPLACE a WITH 8 for sbts=0 OR xcts>25 or xcts<�ֳ��� or xcts<>sbts or sj>7
   *****************�������쳣��ʶa=8Ϊ���ڲ��*******20161027***********************************
endi
count for !'�����'$��ҵ���� and (zbgr=0 OR ybgr=0) TO rs1
****************�ǳ������Ա������������ҹ��*****************************************
if rs1>0
   IF files('kq'+syjj+'.dbf')
      brow part 20 when sykqts() fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',��ҵ����,sbts:h='��������',xcts:h='�ֳ�����',zbgr:h='�а�',ybgr:h='ҹ��' for !'�����'$��ҵ���� and (zbgr=0 OR ybgr=0) titl '��ȷ������'+allt(STR(rs1))+'�˷ǳ���࿼��'
   ELSE
      brow part 20 fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',��ҵ����,sbts:h='��������',xcts:h='�ֳ�����',zbgr:h='�а�',ybgr:h='ҹ��' for !'�����'$��ҵ���� and (zbgr=0 OR ybgr=0) titl '��ȷ������'+allt(STR(rs1))+'�˷ǳ���࿼��'
   ENDIF
    REPLACE a WITH 8 for !'�����'$��ҵ���� and (zbgr=0 OR ybgr=0)
   *****************�������쳣��ʶa=8Ϊ���ڲ��*******20161027***********************************
endif
count for '�����'$��ҵ���� and zbgr+ybgr>=13 TO rs1
****************�������Ա������������ҹ��*****************************************
if rs1>0
   IF files('kq'+syjj+'.dbf')   
      brow part 20 when sykqts() fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',��ҵ����,sbts:h='��������',xcts:h='�ֳ�����',zbgr:h='�а�',ybgr:h='ҹ��' for '�����'$��ҵ���� and zbgr+ybgr>=13 titl '��ȷ������'+allt(STR(rs1))+'�˳���࿼��'
   ELSE
      brow part 20 fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',��ҵ����,sbts:h='��������',xcts:h='�ֳ�����',zbgr:h='�а�',ybgr:h='ҹ��' for '�����'$��ҵ���� and zbgr+ybgr>=13 titl '��ȷ������'+allt(STR(rs1))+'�˳���࿼��'
   ENDIF
    REPLACE a WITH 8 for '�����'$��ҵ���� and zbgr+ybgr>=13
   *****************�������쳣��ʶa=8Ϊ���ڲ��*******20161027***********************************
endif
***********************2.��������/��ҹ�࿼���쳣��˸���********20161024*************
index on BMBH+zw+RYBH to kq&NY 
count for bj+sj+cj+gs+tqj+hunj+sangj+jl+kg+gj+gx+ly+xxpx>0 TO rs1 
zrs=RECCOUNT()
zrs=ALLTRIM(STR(zrs))
********20160627***********
if rs1>0
rs1=ALLTRIM(STR(rs1))
SCAN
DO CASE 
   CASE bj>0
        REPLACE ȱ����� WITH '����'+ALLTRIM(STR(bj))+'��'
   CASE sj>0
        REPLACE ȱ����� WITH '�¼�'+ALLTRIM(STR(sj))+'��' 
   CASE cj>0
        IF xb1='Ů'
           REPLACE ȱ����� WITH '����'+ALLTRIM(STR(cj))+'��'
        ELSE
           REPLACE ȱ����� WITH '�㻤��'+ALLTRIM(STR(cj))+'��'
        ENDIF            
   CASE gs>0
        REPLACE ȱ����� WITH '���˼�'+ALLTRIM(STR(gs))+'��'
   CASE tqj>0
        REPLACE ȱ����� WITH '̽�׼�'+ALLTRIM(STR(tqj))+'��'
   CASE gj>0
        REPLACE ȱ����� WITH '���ݼ�'+ALLTRIM(STR(gj))+'��' 
   CASE hunj>0
        REPLACE ȱ����� WITH '���'+ALLTRIM(STR(hunj))+'��'
   CASE sangj>0
        REPLACE ȱ����� WITH 'ɥ��'+ALLTRIM(STR(sangj))+'��'     
   CASE jl>0
        REPLACE ȱ����� WITH '����'+ALLTRIM(STR(jl))+'��'     
   CASE kg>0
        REPLACE ȱ����� WITH '����'+ALLTRIM(STR(kg))+'��' 
   CASE gx>0
        REPLACE ȱ����� WITH '����'+ALLTRIM(STR(gx))+'��'
   CASE ly>0
        REPLACE ȱ����� WITH '����'+ALLTRIM(STR(ly))+'��'
   CASE xxpx>0
        REPLACE ȱ����� WITH '������ѵ'+ALLTRIM(STR(xxpx))+'��'                   
 ENDCASE
 ENDSCAN              
 brow part 20 fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',xcts:h='����',ȱ�����,bj:h='����',sj:h='�¼�',cj:h='����',gs:h='���˼�',tqj:h='̽�׼�',gj:h='���ݼ�',hunj:h='���',sangj:h='ɥ��',jl:h='����',kg:h='����',gx:h='����',ly:h='����',xxpx:h='��ѵ';
   titl nd+'��'+allt(str(val(yf)))+'�·�ȱ����Ա����'+rs1+'/'+zrs+'�� ��[ Esc ]���˳�' for  bj+sj+cj+gs+tqj+hunj+sangj+jl+kg+gj+gx+ly+xxpx>0
ENDIF
repl ERP��� with erprybh,�������� with sbts,�ֳ����� with xcts,�а� with zbgr,ҹ�� with ybgr,�¾� with xjgr,���ռӰ� with jrjb,���ռӰ� with jjb,������ѵ WITH xxpx,���� with bj,�¼� with sj,���� with cj,���� with gs,̽�׼� with tqj,��ɥ�� with hsj,��� with hunj,ɥ�� with sangj,���� with ly,���� WITH gx,���� with jl,���� with kg,���ݼ� with gj,������ with bj1,�ϸ����� with sfgz,����ˮ�� with fzsd,�ۿ� with kk,������ with fzj all
copy to &bf.\ȱ���쳣��� fiel dwmc,cjmc,bmmc,ryxm,xb1,��ҵ����,��������,�ֳ�����,�а�,ҹ��,ȱ�����,������ѵ,����,�¼�,����,����,̽�׼�,��ɥ��,���,ɥ��,����,����,����,����,���ݼ� type XL5 for bj+sj+cj+gs+tqj+hunj+sangj+jl+kg+gj+gx+ly+xxpx>0 OR a=8
*************************************************3.�˹����뿼�ں�ȫ���û�����ERPϵͳ���ں����ֶα���****************20160812*********************�ṩȱ�����������쳣�б��˹���***************
REPLACE hsj WITH hunj+sangj,a WITH 1 all
************��١�ɥ�ٵ���******20160405**********************
go top
loca for jrjb+jjb+bj+sj+cj+gs+tqj+hsj+jl+kg+gj>0
if !eof()
 IF files('kq'+syjj+'.dbf') 
 brow part 20 when sykqts() fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',xb1,��ҵ����,��������,�ֳ�����,�а�,ҹ��,ȱ�����,jrjb:h='���ռӰ�',jjb:h='���ռӰ�',xxpx:h='������ѵ',bj:h='����',sj:h='�¼�',cj:h='�㻤/����',gs:h='���˼�',tqj:h='̽�׼�',gj:h='���ݼ�',hunj:h='���',sangj:h='ɥ��',gx:h='����',ly:h='����',jl:h='����',kg:h='����';
   titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'��  ��[ Esc ]���˳�' for jrjb+bj+sj+cj+gs+tqj+hsj+jl+kg+gj>0
 ELSE
  brow part 20 fiel cjmc:h='��λ',bmmc:h='����',ryxm:h='����',xb1,��ҵ����,��������,�ֳ�����,�а�,ҹ��,ȱ�����,jrjb:h='���ռӰ�',jjb:h='���ռӰ�',xxpx:h='������ѵ',bj:h='����',sj:h='�¼�',cj:h='�㻤/����',gs:h='���˼�',tqj:h='̽�׼�',gj:h='���ݼ�',hunj:h='���',sangj:h='ɥ��',gx:h='����',ly:h='����',jl:h='����',kg:h='����';
   titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'��  ��[ Esc ]���˳�' for jrjb+bj+sj+cj+gs+tqj+hsj+jl+kg+gj>0
 ENDIF
ENDIF
***********************���¿����쳣����������********20161024*************
close all
select 2
 IF FILE('ry'+ny+'.dbf')
    use ry&ny 
    index on RYBH to ry&ny
 ELSE
    use ryzk 
    index on RYBH to ryzk
 ENDIF 
 **************����ȡ�������������䣨������㣩****���ϸ��������벻�����ꡱ*****20161121*************
scan FOR 'ְ��'$ygxz1
����=INT(gl)
DO case
CASE ����<1
     REPLACE sfbl WITH 0
CASE ����=1
     REPLACE sfbl WITH 20
CASE ����=2
     REPLACE sfbl WITH 40
CASE ����=3
     REPLACE sfbl WITH 60
CASE ����=4
     REPLACE sfbl WITH 80
CASE ����>=5
     REPLACE sfbl WITH 100
ENDCASE
endscan              
select 1
use kq&ny
REPLACE sfgz WITH 0 all
IF sfgz1='����'
inde on rybh to kq&ny
set relation to rybh into 2
repl sfgz with ��׼*b.sfbl/100 for &����1
*****�����Ʒ���ʽ******************xcts>=15*******20160115*************
count for &����2 to rs1
if rs1>0
   repl sfgz with xcts*�ձ�*b.sfbl/100 for &����2
*****�۷���ʽ*******************************xcts<15*******20160115*************
   brow field cjmc:h='��λ',bmmc:h='����',ryxm:h='����',b.cjgzrq:h='�μӹ�������',b.��������,b.sfbl:h='�ϸ�������%��',gj:h='���ݼ�',xcts:h='�ֳ�����',sfgz:h='�ϸ�����' for &����2 titl '������Ա�ֳ������������ڱ�׼�Ʒ�������'+����1+'���涨Ӧ�۷��ϸ�����'
endif
count for &����3 to rs2
if rs2>0
   repl sfgz with 0 for &����3
*****�۷���ʽ**********tqj+sj>15*******20160115*************
   brow field cjmc:h='��λ',bmmc:h='����',ryxm:h='����',b.cjgzrq:h='�μӹ�������',b.��������,b.sfbl:h='�ϸ�������%��',gj:h='���ݼ�',xcts:h='�ֳ�����',sfgz:h='�ϸ�����' for &����3 titl '������Աȫ��ȱ�ڳ�����׼�Ʒ�������'+����3+'���ļ��涨�����ϸ�����'
ENDIF
count for sfgz=0 to rs3
count for sfgz>0 to rs
if rs>0
   INDEX on bmbh+zw+rybh TO kq&ny
   GO top
   brow field cjmc:h='��λ',bmmc:h='����',ryxm:h='����',b.cjgzrq:h='�μӹ�������',b.��������,b.sfbl:h='�ϸ�������%��',gj:h='���ݼ�',xcts:h='�ֳ�����',sfgz:h='�ϸ�����' titl '����Ӧ�Ʒ��ϸ�����'+allt(STR(rs))+'�ˣ����У��۷��ϸ�����'+allt(STR(rs1))+'�ˣ������ϸ�����'+allt(STR(rs2))+'�ˣ����������޲��M1����'+allt(STR(rs3))+'��'
ENDIF
ENDIF
REPLACE bj1 WITH xcts*bjjb all
*************1.������������׼���ֳ�����****************
yn = MESSAGEBOX("���ڼ������������"+bf+"\��ȱ���쳣��������ӱ��У��ִ�ʹ����",4+32,"��ʾ")
IF yn = 6
**************��׼Excel��ȡ�ļ�*************
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\ȱ���쳣���.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
ENDIF
close all
on key label F1 do grcx
*********��ԭ�ȼ�����*********************
i=1
*********��ԭȫ�ֱ���i=1*********************
CLEAR
*************************������ͼƬ�����÷�׽���*****************
use ccrq
go 8
��ֽ��=fdmax
I=val(ͼƬ)
pic=ALLTRIM(ͼƬ)
repl ͼƬ with ALLTRIM(str(I+1))
IF I>=��ֽ��
   repl ͼƬ with '0'
ENDIF
**********�û�������־�����׽���*********
pict_fd='fd'+pic+'.jpg'
_SCREEN.PICTURE = '&pict_fd'
close all
RETURN


******************
PROCEDURE KQgrcx
******************
do while .t.
      do srxm.spr
      sele 2
      go top
      locate for allt(RYXM)=allt(XM) or allt(RYBH)=allt(XM)
      JLH=recn()
      if eof()
        do dhkjg.spx
    M = readkey()
    if M=12 or M=268
      return
    endif
     else
      wait wind '��Ҫ���ҵ���Ա���ҵ������Ȱ�[PgDn][PgUp]������ˢ��һ��' nowa 
      exit    
     endif
ENDDO
******************
PROCEDURE plsr
******************
select 2
jlh=RECNO()
sb=sbts
xc=xcts
zb=zbgr
yb=ybgr
jr=jrjb
yn = MESSAGEBOX("Ԥͳһ����ǰ��¼����������������",4+32,"��ʾ")
IF yn = 6
   REPLACE sbts WITH sb,xcts WITH xc,zbgr WITH zb,ybgr WITH yb,jrjb WITH jr FOR sbts=0
ENDIF
go jlh
******************
PROCEDURE SYKQTS
******************
wait wind alltrim(b.RYXM)+'[����]��'+'�а�:'+alltrim(str(d.zbgr))+'ҹ��:'+alltrim(str(d.ybgr))+'����:'+alltrim(str(d.Bj1))+'�Ӱ�:'+alltrim(str(d.JrJB))+';����:'+alltrim(str(d.BJ))+';'+'�¼�:'+alltrim(str(d.SJ))+';'+'�㻤/����:'+alltrim(str(d.CJ))+';����:'+alltrim(str(d.GS))+';̽��:'+alltrim(str(d.TQJ))+';���:'+alltrim(str(d.HunJ))+';ɥ��:'+alltrim(str(d.SangJ))+';����:'+alltrim(str(d.JL))+';����:'+alltrim(str(d.KG)) nowa at 50,120

******************
PROCEDURE SYKQTS1
******************
wait wind alltrim(b.RYXM)+'[����]��'+'�а�:'+alltrim(str(b.zbgr))+'ҹ��:'+alltrim(str(b.ybgr))+';����:'+alltrim(str(b.BJ))+';'+'�¼�:'+alltrim(str(b.SJ))+';���:'+alltrim(str(b.HunJ))+';ɥ��:'+alltrim(str(b.SangJ))+';����:'+alltrim(str(b.KG)) nowa at 50,120