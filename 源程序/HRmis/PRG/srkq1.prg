***************************
* .\kqSR.PRG(Build20140123)
***************************
on escape
set escape off
on key label F1 do grsr
on key label F2 do sykqts
close all
set safety off
set talk off
clear 
_SCREEN.PICTURE = ''
dh=''
dycj=''
dybm=''
ND = '  '
YF = '  '
XH1 = 'N'
LY = ''
T123 = '  '
close all
set confirm off
XH1 = 'N'
ND1 = year(date())
YF1 = month(date())
do srny2.spx
LYNY = ND+YF
****************ǰ���Զ�����****************
if val(yf)<3
  JJND = str(val(nd)-1,4)
  JJYF = str(10+val(yf),2)
else
  JJND = ND
  JJYF = iif(val(YF)-2>9,str(val(YF)-2,2),'0'+str(val(YF)-2,1))
endif
JJLY = JJND+JJYF
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
****************��ʼ��****************
T123 = '  '
do dhkdy
do case
  case dh=2
  T123 = dyCJ
  case dh=3
  T123 = dybm
endcase
close all
clear
XH1 = 'N'
if !file('LW'+LYNY+'.DBF')
  use lwpqgzk excl
  zap
  COPY TO lw&LYNY
  USE lw&LYNY
  append from ryzk for ygxz='02'
endif
close all
*********************��ʼ��****************
select 2
use ryzk
index on RYBH to ryzk
select 1
use lw&LYNY
inde on rybh to lw&LYNY
update on RYBH from B replace CJDM with b.CJDM , BMBH;
 with b.BMBH , CJMC with b.CJMC , BMMC with b.BMMC ,erprybh with b.erprybh,gz1 with b.gz1,gw with b.gw
*********************�˹����뿼��****************
close all
select 4
use lw&syJJ
index on RYBH to temp
select 3
use ryzk
index on RYBH to ryzk
select 1
use SMK1
select 2
use lw&LYNY
set relation to RYBH into 3
set relation to RYBH into 4 additive
replace BJ with 21 for c.RYFL='33'
index on BMBH+RYBH to lw&LYNY
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
  @ 23 , 33 say 'PgDn=��һ����¼��PgUp=��һ����¼' font '����' , 15
  @ 25 , 33 say 'F1=����ָ����Ա��F2=��ѯ���¿��ڣ�Esc=�����˳�' font '����' , 15
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
        @kkk1,35+lll1 get &ryhds font '����' , 14 PICT &TM123
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
enddo
********************�������ݺ͸��Դ�������***********************
on key label F1 do grcx
on key label F2
select 2
*replace JRJB with 0,zbgr with 0,ybgr with 0,xcts with 0 for '��'$c.ZW1
*************************���⴦��*************************************
index on BMBH+RYBH to lw&LYNY
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
   brow part 20 when sykqts() fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',sbts:h='��������',zbgr:h='�а�',ybgr:h='ҹ��',bj1:h='����',jrjb:h='�Ӱ�',bj:h='����',sj:h='�¼�',cj:h='����',gs:h='����',tqj:h='̽��',hsj:h='��ɥ',jl:h='����',kg:h='����',fzsd:h='����ˮ��',kk:h='�ۿ�';
    titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs1+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'�ˣ����У�'+cjm+'��'+rs1+'�ˣ�����='+bjj1+'�ˣ��¼�='+sjj1+'�ˣ�����ˮ�磽'+allt(str(sfzsd1,10,1))+'Ԫ���ۿ'+allt(str(skk1,10,1))+'Ԫ   ��[ Esc ]���˳�'
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
  brow part 20 when sykqts() fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',sbts:h='��������',zbgr:h='�а�',ybgr:h='ҹ��',bj1:h='����',jrjb:h='�Ӱ�',bj:h='����',sj:h='�¼�',cj:h='����',gs:h='����',tqj:h='̽��',hsj:h='��ɥ',jl:h='����',kg:h='����',fzsd:h='����ˮ��',kk:h='�ۿ�';
   titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs1+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'�ˣ����У�'+bm+'��'+rs1+'�ˣ�����='+bjj1+'�ˣ��¼�='+sjj1+'�ˣ�����ˮ�磽'+allt(str(sfzsd2,10,1))+'Ԫ���ۿ'+allt(str(skk2,10,1))+'Ԫ  ��[ Esc ]���˳�'
   endif
else
   go top
  brow part 20 when sykqts() fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',sbts:h='��������',zbgr:h='�а�',ybgr:h='ҹ��',bj1:h='����',jrjb:h='�Ӱ�',bj:h='����',sj:h='�¼�',cj:h='����',gs:h='����',tqj:h='̽��',hsj:h='��ɥ',jl:h='����',kg:h='����',fzsd:h='����ˮ��',kk:h='�ۿ�';
   titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'�ˣ�����ˮ�磽'+allt(str(sfzsd,10,1))+'Ԫ���ۿ'+allt(str(skk,10,1))+'Ԫ  ��[ Esc ]���˳�'
endif
go top
loca for jrjb+bj+sj+cj+gs+tqj+hsj+jl+kg+gj>0
if !eof()
 brow part 20 when sykqts() fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',jrjb:h='�Ӱ�',bj:h='����',sj:h='�¼�',cj:h='����',gs:h='���˼�',tqj:h='̽�׼�',gj:h='���ݼ�',hsj:h='��ɥ��',jl:h='����',kg:h='����';
   titl '���������'+nd+'��'+allt(str(val(yf)))+'�·ݿ��ڣ���'+rs+'�ˣ�����='+bjj+'�ˣ��¼�='+sjj+'��  ��[ Esc ]���˳�' for jrjb+bj+sj+cj+gs+tqj+hsj+jl+kg+gj>0
endif
close all
********************DHS*****************
*if val(dm111)=99
*sele 2
*use ryzk
*inde on rybh to ryzk
*select 1
*use lw&LYNY
*inde on rybh to lw&lyny
*set relation to rybh into 2
 *  repl sfgz with 800*b.sfbl/100 for xcts>=15
  * repl sfgz with xcts*round(800/15,2)*b.sfbl/100 for xcts<15
   *repl sfgz with 0 for tqj+gs>=21
   *repl jang2 with 240,sfgz with 250*c.sfbl/100 for bmbh='0114' and xcts>=15
   *repl jang2 with 240,sfgz with xcts*round(250/15,2)*c.sfbl/100 for bmbh='0114' and xcts<15
*endif
close all
delete file temp.idx
delete file temp.dbf
*************************��ԭͼƬ*****************
clear
 _SCREEN.WINDOWSTATE = 2
close all
use ccrq
go 8
repl ͼƬ with str(val(ͼƬ)+1,2)
pict_fd='fd'+allt(ͼƬ)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
retu


******************
PROCEDURE SYKQTS
******************
wait wind alltrim(b.RYXM)+'[����]��'+'�а�:'+alltrim(str(d.zbgr))+'ҹ��:'+alltrim(str(d.ybgr))+'����:'+alltrim(str(d.Bj1))+'�Ӱ�:'+alltrim(str(d.JrJB))+';����:'+alltrim(str(d.BJ))+';'+'�¼�:'+alltrim(str(d.SJ))+';'+'����:'+alltrim(str(d.CJ))+';����:'+alltrim(str(d.GS))+';̽��:'+alltrim(str(d.TQJ))+';��ɥ:'+alltrim(str(d.HSJ))+';����:'+alltrim(str(d.JL))+';����:'+alltrim(str(d.KG)) nowa at 34,10
