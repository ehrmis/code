***************************
*jjSR1.PRG( Build20081015 )
***************************
on escape
set escape off
on key label F1 do grsr
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
T123 = '  '
close all
TS = 0
TS1 = 0
XH1 = 'N'
set confirm off
ND1 = year(date())
YF1 = month(date())
*if yf1=0
*   yf1=12
*   nd1=nd1-1
*endif   
clear
do srny.spx
NY = ND + YF
if !file('KQ'+NY+'.DBF')
  use KQK
  COPY TO KQ&NY stru
  USE KQ&NY
  append from RYZK for zgqk='01'  
endif
*********************��ʼ��****************
close all
select 2
use ryzk
index on RYBH to ryzk
select 1
use kq&NY
inde on rybh to kq&NY
update on RYBH from B replace CJDM with b.CJDM , BMBH;
 with b.BMBH , CJMC with b.CJMC , BMMC with b.BMMC ,erprybh with b.erprybh
close all
do dhkdy
do case
   case dh=2
  T123 = dyCJ
  case dh=3
  T123 = dybm
endcase 
use kq&NY
on key label F1 do grcx
do case
case dh=1
clear
go top
brow titl '����������'+nd+'��'+allt(str(val(yf)))+'�·ݼ�Ч���ʣ���'+allt(str(recc()))+'�� ��[ F1 ]������ ��[ Esc ]���˳�';
fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',jang1:h='��Ч����',jang2:h='���'
count for jang2>0 to rs1
sum jang2 to xyj
if rs1>0
   pjxyj=round(xyj/rs1,1)
   pjxyj=allt(str(pjxyj,10,1))
else
   xyj=0
   pjxyj='0'
endif
count for jang1>0 to rs
sum jang1 to zhj1
zhj=allt(str(zhj1,10,2))
if rs>0
   pjzhj=round(zhj1/rs,1)
   pjzhj=allt(str(pjzhj,10,1))
   pjj=round((zhj1+xyj)/rs,1)
   pjj=allt(str(pjj,10,1))   
   rs=str(rs,4)
else
   zhj='0'
   pjzhj='0'
   pjj='0'
   rs='0'
endif
go top
brow titl '��'+rs+'�ˣ���Ч����='+zhj+'Ԫ(�˾�='+pjzhj+'Ԫ/��)�����='+allt(str(xyj,10,2))+'Ԫ(�˾�='+pjxyj+'Ԫ/��)���ϼƣ�'+allt(str(zhj1+xyj,10,2))+'Ԫ(�˾�='+pjj+'Ԫ/��)';
fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',jang1:h='��Ч����',jang2:h='���' 
case dh=2
count for cjdm=T123 to rs
go top
brow titl '����������'+nd+'��'+allt(str(val(yf)))+'�·ݼ�Ч���ʣ���'+allt(str(rs))+'�� ��[ F1 ]������ ��[ Esc ]���˳�';
fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',jang1:h='��Ч����',jang2:h='���' for cjdm=T123
   count for cjdm=T123 and jang2>0 to rs1
   sum jang2 to xyj for cjdm=T123
   if rs1>0
       pjxyj=round(xyj/rs1,0)
       pjxyj=allt(str(pjxyj,10,1))
   else
       xyj=0
       pjxyj='0'
   endif
count for cjdm=T123 and jang1>0 to rs 
   sum jang1 to zhj1 for cjdm=T123
   zhj=allt(str(zhj1,10,2))
   if rs>0
   pjzhj=round(zhj1/rs,1)
   pjzhj=allt(str(pjzhj,10,1))
   pjj=round((zhj1+xyj)/rs,1)
   pjj=allt(str(pjj,10,1))   
   rs=str(rs,4)
else
   zhj='0'
   pjzhj='0'
   pjj='0'
   rs='0'
endif
   go top
   loca for cjdm=T123   
   brow titl '��'+rs+'�ˣ���Ч����='+zhj+'Ԫ(�˾�='+pjzhj+'Ԫ/��)�����='+allt(str(xyj,10,2))+'Ԫ(�˾�='+pjxyj+'Ԫ/��)���ϼƣ�'+allt(str(zhj1+xyj,10,2))+'Ԫ(�˾�='+pjj+'Ԫ/��)';
fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',jang1:h='��Ч����',jang2:h='���' for cjdm=T123
set filter to
case dh=3
count for bmbh=T123 to rs
go top
brow titl '����������'+nd+'��'+allt(str(val(yf)))+'�·ݼ�Ч���ʣ���'+allt(str(rs))+'�� ��[ F1 ]������ ��[ Esc ]���˳�';
fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',jang1:h='��Ч����',jang2:h='���' for bmbh=T123
   count for bmbh=T123 and jang2>0 to rs1
   sum jang2 to xyj for bmbh=T123
   if rs1>0
      pjxyj=round(xyj/rs1,0)
      pjxyj=allt(str(pjxyj,10,1))
   else
      xyj=0
      pjxyj='0'     
   endif
count for bmbh=T123 and jang1>0 to rs  
   sum jang1 to zhj1 for bmbh=T123
   zhj=allt(str(zhj1,10,2))
if rs>0
   pjzhj=round(zhj1/rs,1)
   pjzhj=allt(str(pjzhj,10,1))
   pjj=round((zhj1+xyj)/rs,1)
   pjj=allt(str(pjj,10,1))   
   rs=str(rs,4)
else
   zhj='0'
   pjzhj='0'
   pjj='0'
   rs='0'
endif
   go top
   loca for bmbh=T123   
brow titl '��'+rs+'�ˣ���Ч����='+zhj+'Ԫ(�˾�='+pjzhj+'Ԫ/��)�����='+allt(str(xyj,10,2))+'Ԫ(�˾�='+pjxyj+'Ԫ/��)���ϼƣ�'+allt(str(zhj1+xyj,10,2))+'Ԫ(�˾�='+pjj+'Ԫ/��)';
 fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',jang1:h='��Ч����',jang2:h='���' for bmbh=T123
set filter to
endcase
*******�Զ�ɸѡERP��������*********
repl ERP��� with erprybh,�а� with zbgr,ҹ�� with ybgr,���ռӰ� with jrjb,���� with bj,�¼� with sj,���� with cj,���� with gs,̽�׼� with tqj,;
��ɥ�� with hsj,���� with jl,���� with kg,���ݼ� with gj,������ with bj1*bjjb,jang with jang1+jang2,�ϸ����� with sfgz,��Ч���� with jang,����ˮ�� with fzsd,������ with fzj all
ERPKQ='ryxm,ERP���,�а�,ҹ��'
J = 8
go top
do while j<24
  tt=fiel(j)
IF type(field(j))='N'
sum &tt to stt
if stt>0
     ERPKQ=allt(ERPKQ)+','+'&tt'
endif
ENDIF
  j=j+1
enddo
************����ERP����************
index on BMBH+RYBH to kq&NY
 yn = MESSAGEBOX("��Ҫ����"+ND+"��"+allt(str(val(YF)))+"�·ݼ�Ч����������",4+32,"��ʾ")
IF yn = 6 
 old_dirs='d:'
Declare integer Qdir_m In Qdir.dll string @ck,;
integer yy,;
string cc
*ע��Qdir_m()��@ck=���ص��ļ���·����yy=�ļ������,cc=��ʾ
*==========================================================
ck=spac(64)      &&����ļ���·��
yy=1
cc='��ѡ����ļ�����Ŀ¼,����ȷ����ť��'
xx=Qdir_m(@ck,yy,cc)   
*����Qdir_m()��ѡ��Ľ����ck�з���
dirs=allt(subs(ck,2,64))
if len(dirs)=0
   dirs=old_dirs
endif
if right(dirs,1)<>'\'
   dirs=dirs+'\'
endif
wjm='KQ'
copy to &dirs.&wjm.&nd.&yf TYPE fox2x
wjm='ERP���ڽ���'
copy to &dirs.&wjm fiel &ERPKQ for zbgr+ybgr+jrjb+bj+sj+cj+gs+tqj+hsj+jl+kg+bj1+gj+sfgz+JANG+FZSD+KK+FZJ>0 and ygxz='01' and zgqk='01' type xl5
 =messagebox("����ERP��Ч���������ѳɹ���������",48,"��ϲ")  
ENDIF
***********************************
clear
 _SCREEN.WINDOWSTATE = 2
close all
use ccrq
go 8
repl ͼƬ with str(val(ͼƬ)+1,2)
pict_fd='fd'+allt(ͼƬ)+'.jpg'
close all
_SCREEN.PICTURE = '&pict_fd'
ND=str(year(date()),4)
YF=iif(month(date())>9,str(month(date()),2),'0'+str(month(date()),1))
close all
retu
