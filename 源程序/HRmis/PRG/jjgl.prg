*************************
* jjgl.prg(Build20090117)
*************************
close all
set safety off
ND1 = year(date())
YF1 = month(date())
clear
do srny.spx
ly=nd
ly1=str(val(nd)-1,4)
gzsrk=sblj+'zr1'+ly
if  not file(gzsrk+'.DBF') 
    MESSAGEBOX('���Ƚ���'+ly+'�깤���ܶ������','��ʾ')
    retu
endif
wait window '�������ݴ���!���Ժ�......' nowait
close all
sele 2
use &sblj.zr&ly
inde on rybh to &sblj.zr&ly
sele 1
use &sblj.zr&ly1
inde on rybh to &sblj.zr&ly1
upda on rybh from B repl cjdm with b.cjdm , cjmc with b.cjmc , bmbh with b.bmbh , bmmc with b.bmmc
copy to &sblj.zr11&ly1
copy to &sblj.zr11temp
close all
*********������ͬ�ڱȽϹ����ܶ�************
  use &sblj.zr11&ly1
  replace MOU with 0 , HJ with 0 all
  for IJB = 1 to yf1
    CMJB = iif(IJB>9,str(IJB,2),'0'+str(IJB,1))
    if  not file('GZ'+ly+CMJB+'.DBF')
      exit
    endif
    replace HJ with round(HJ+J&cmjb,0),MOU with val(CMJB) all
    repl PJ with round(HJ/mou,0) , A with 1 for mou>0
  endfor 
  sum hj,mou to hjtem,moutem
  if hjtem=0 or moutem=0 
     use &sblj.zr11temp
     copy to &sblj.zr11&ly1
  endif    
close all
sele 5
use ryzk
index on RYBH to ryzk
select 4
use &sblj.zr11&ly1 excl 
repl A with 1 all
index on rybh to &sblj.zr11&LY1
set relation to rybh into 5
repl cjdm with e.cjdm,m03 with val(e.zw) for rybh=e.rybh
repl cjdm with '88',cjmc with '�������Ǹ�' for m03=7
repl cjdm with '99',cjmc with '�Ƽ��ɲ�' for m03=5 or m03=6
delete for jjhj=0
pack
index on CJDM to temp
total on CJDM to &sblj.zr11 
select 3
use &sblj.zr&LY
copy to temp
use temp
repl hj with jjhj,PJ with round(HJ/mou,0) , A with 1 for mou>0
set relation to rybh into 5
repl cjdm with e.cjdm,m03 with val(e.zw) for rybh=e.rybh
repl cjdm with '88',cjmc with '�������Ǹ�' for m03=7
repl cjdm with '99',cjmc with '�Ƽ��ɲ�' for m03=5 or m03=6
index on CJDM to &sblj.zr&LY
total on CJDM to &sblj.zr1 
select 2
use &sblj.zr11
replace PJ with round(HJ/A,0) all
sum to LJ1,RS1 HJ,A
index on CJDM to &sblj.zr11
select 1
use &sblj.zr1 excl
delete for jjhj=0
pack
sum jjHJ,A to LJ,RS
set relation to cjdm into 2
replace PJ with round(jjHJ/A,0) , m01 with round(PJ-b.PJ,0) all
replace m02 with round(m01/b.PJ*100,1) all
appe blank
go bott
repl cjmc with 'ȫ�������ۼ�',hj with LJ,pj with round(lj/rs,2),m01 with round(LJ/RS,1)-round(LJ1/RS1,1),m02 with (round(LJ/RS,1)-round(LJ1/RS1,1))/round(LJ1/RS1,1)*100 
copy to jjtemp fiel cjmc,hj,pj,m01,m02
use ����ͳ�� excl
zap
appe from jjtemp
repl ���� with cjmc,��Ч���� with hj,�˾� with pj,���ʶ� with m01,���� with m02 all
select 4
index on RYBH to zr11temp
select 3
set relation to RYBH into 4 
  clear
  for zdz=1 to fcount()
  zd=fiel(zdz)
  zdz1=(fsize(zd)+fsize(zd))*2.4
  if uppe(allt(zd))='RYXM' 
     exit
  endif
  endfor 
  repl m01 with d.hj ,m02 with d.pj,m03 with round((PJ-m02)/m02*100,1) for d.pj>0  
  copy to  jjtemp fiel cjmc,bmmc,ryxm,hj,m01,pj,m02,m03
use �������� excl
zap
appe from jjtemp
repl ���� with cjmc,���� with bmmc,��Ч���� with hj,ƽ�� with pj,ƽ�����ʶ� with pj-m02,���� with m03 all
close all
use gzjj excl
zap
appe from &gzsrk for jjhj>0 and mou>0
sele 3
use &sblj.zr&LY
inde on rybh to zr&LY
sele 2
use ryzk
inde on rybh to ryzk
sele 1
use gzjj
inde on rybh to gzjj
upda on rybh from C repl ylbx with c.ylbx,ybx with c.ybx,sybx with c.sybx,gjj with c.gjj,sds with c.sds,sfje with c.sfje 
repl pj with round(hj/mou,0),jfjs with round(jjhj/mou,0),a with 1 all
inde on bmbh to gzjj
tota on bmbh to bmjj
use bmjj
sum hj,jjhj,a,ylbx,ybx,sybx,gjj,sds,sfje to gzsr,jjsr,rs,yl,yb,sy,zf,ss,sf 
repl pj with round(hj/a,0),jfjs with round(jjhj/a,0) all
appe blan
go bott
repl cjmc with '����  ��    ',hj with gzsr,jjhj with jjsr,;
pj with round(hj/rs,0),jfjs with round(jjhj/rs,0),ylbx with yl,ybx with yb,sybx with sy,gjj with zf,sds with ss,sfje with sf
***********************************************************
use gzjj
inde on cjdm to gzjj
tota on cjdm to cjjj
use cjjj
sum hj,jjhj,a,ylbx,ybx,sybx,gjj,sds,sfje to gzsr,jjsr,rs,yl,yb,sy,zf,ss,sf 
repl pj with round(hj/a,0),jfjs with round(jjhj/a,0) all
appe blan
go bott
repl cjmc with '����  ��    ',hj with gzsr,jjhj with jjsr,;
pj with round(hj/rs,0),jfjs with round(jjhj/rs,0),ylbx with yl,ybx with yb,sybx with sy,gjj with zf,sds with ss,sfje with sf
***********************************************************
use gzjj
set relation to rybh into 2
repl a with 0 for val(b.zw)>4 and val(b.zw)<7
copy to kjjj for a=0
use kjjj
set relation to rybh into 2
repl cjmc with b.zw1,a with 1 all
sum hj,jjhj,a,ylbx,ybx,sybx,gjj,sds,sfje to gzsr,jjsr,rs,yl,yb,sy,zf,ss,sf 
appe blan
go bott
repl cjmc with '����  ��    ',hj with gzsr,jjhj with jjsr,;
pj with round(hj/rs,0),jfjs with round(jjhj/rs,0),ylbx with yl,ybx with yb,sybx with sy,gjj with zf,sds with ss,sfje with sf
sort on hj to �Ƽ�����
***********************************************************
close all
select 1
use jj excl
pack
replace ���ع��� with ���� for ���='����' and ���ط�=.T.
replace ���ص�� with ���� for !'����'$��� and ���ط�=.T.
replace ���� with 0-����,�������� with ���� for ���='����' and ���ط�=.F. and  ����>0
replace ���� with 0-����,������� with ���� for !'����'$��� and ���ط�=.F. and  ����>0
inde on cjdm+bmbh to jj
repl �����ܼ� with 0 all
tota on cjdm+bmbh to cjhz field �������,��������,����
use cjhz excl
delete for val(cjdm)=0
pack
repl �����ܼ� with ���� all
sum �������,��������,�����ܼ� to dxj,zhj,zj
appe blan
go bott
repl �������� with '����  ��    ',������� with dxj,�������� with zhj,�����ܼ� with zj 
use ��λ���� excl
zap
appe from cjhz
use jj
inde on allt(�·�) to jj
tota on allt(�·�) to dyhj field ����,���ع���,���ص��,��������,�������
select 2
use dyhj
repl ���·����� with ��������+������� all
repl �������غ� with ���ع���+���ص�� all
go top
do while !eof()
select 1
repl �������ع��� with b.���ع��� for allt(�·�)=allt(b.�·�)
repl �������ص�� with b.���ص�� for allt(�·�)=allt(b.�·�)
repl ���·������� with b.�������� for allt(�·�)=allt(b.�·�)
repl ���·������ with b.������� for allt(�·�)=allt(b.�·�)
repl ���·����ϼ� with b.���·����� for allt(�·�)=allt(b.�·�)
repl �������غϼ� with b.�������غ� for allt(�·�)=allt(b.�·�)
select 2  
skip
enddo
close all
use jj
calculate for ���ط�=.T. and ���='����' to PHZHJ sum(����)
calculate for ���ط�=.F. and ���='����' to FCZHJ sum(����)
calculate for ���ط�=.T. and !'����'$��� to PHDXJ sum(����)
calculate for ���ط�=.F. and !'����'$��� to FCDXJ sum(����)
calculate for ���ط�=.T. to PHZJ sum(����)
calculate for ���ط�=.F. to FCZJ sum(����)
calculate for ����='���ս�' to NZJ sum(����)
calculate for ����='���Ƚ�' to JDJ sum(����)
repl �ۼ����ع��� with PHZHJ,�ۼ����ص�� with PHDXJ,�ۼƷ������� with FCZHJ,�ۼƷ������ with FCDXJ,�����ܼ� with PHZJ,�����ܼ� with FCZJ all
repl �ۼ� with 0 all
inde on allt(����) to jj
repl �ۼ� with 0 all
tota on allt(����) to lj1 field ����,�ۼ�
use lj1
repl �ۼ� with ���� all
use jj
inde on allt(cjdm)+allt(bmbh)+allt(����) to jj
repl �ۼ� with 0 all
tota on allt(cjdm)+allt(bmbh)+allt(����)+allt(�·�) to lj field ����,�ۼ�
close all
sele 2
use lj
repl �ۼ� with ���� all
inde on allt(cjdm)+allt(bmbh)+allt(����)+allt(�·�) to lj
sele 1
use jj
set relation to allt(cjdm)+allt(bmbh)+allt(����)+allt(�·�) into 2
go top
do while !eof()
repl �ۼ� with b.���� 
skip
enddo
close all
IF !DBUSED('JJGL') 
      OPEN DATABASE JJGL 
 ENDIF   
  IF !DBUSED('DMK') 
      OPEN DATABASE DMK 
 ENDIF   
 use jj 
 sort on �·�,cjdm,bmbh to ������ϸ field �·�,��������,��������,����,���,���ط�,����,�ۼ�
 sort on cjdm,bmbh,����,���,�·� to ��λ��ϸ field ��������,��������,����,���,����,�ۼ�
 inde on allt(cjdm)+allt(bmbh)+allt(�·�)+allt(����)+allt(���) to jj
 sort on cjdm,bmbh,�·�,����,��� to str(ND1,4)+'��н��'
 use lj1 
 sort on ����,��� to �����ۼ� field ����,���,���ط�,�ۼ�
 use lj 
 sort on cjdm,bmbh,���� to ����λн�� field ��������,��������,����,�ۼ�
 close all
 delete file lj.dbf
 delete file lj1.dbf
  IF !DBUSED('JJGL') 
      OPEN DATABASE JJGL 
 ENDIF   
  IF !DBUSED('DMK') 
      OPEN DATABASE DMK 
 ENDIF   
do form jj
