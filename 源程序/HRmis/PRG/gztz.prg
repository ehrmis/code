******************************
* .\GZtz.PRG(Build20070128)
******************************
nd=''
yf=''
ND1 = year(date())
YF1 = month(date())
close all
delete file gztemp.dbf
delete file temp.dbf
do srny.spx
ly=nd
ny=''
dh=''
dycj=''
dybm=''
close all
set safety off
jb1='skc' &&ѡ��
use gztz excl
zap
copy to dw&JB1 
copy to gztemp
n1=LY+'01'
ni=val(n1)
n2=LY+YF
nii=val(n2)
do dhkdy
for ny1=ni to nii
  ny=allt(str(ny1)) 
  USE gztz excl
  if file('GZ'+NY+'.DBF')
  do case
case DH=1
    APPE FROM GZ&NY fiel hj,bzgz,jbgz,jang,ZYBF,HT,ZF,GWJT,sfGZ,kk,STJT,BJ,ylbx,sybx,ybx,gjj,sds,sfje,a,rybh,dwdm,cjdm,cjmc,bmbh,bmmc
case DH=2
    APPE FROM GZ&NY fiel hj,bzgz,jbgz,jang,ZYBF,HT,ZF,GWJT,sfGZ,kk,STJT,BJ,ylbx,sybx,ybx,gjj,sds,sfje,a,rybh,dwdm,cjdm,cjmc,bmbh,bmmc for cjdm='&dycj'
case DH=3
    APPE FROM GZ&NY fiel hj,bzgz,jbgz,jang,ZYBF,HT,ZF,GWJT,sfGZ,kk,STJT,BJ,ylbx,sybx,ybx,gjj,sds,sfje,a,rybh,dwdm,cjdm,cjmc,bmbh,bmmc for bmbh='&dybm'
endcase
  endif
  yrs=recc()
  repl a with 0 all
  index on DWDM to gztz
  TOTAL ON DWDM TO DW&NY
  copy to temp 
ZAP
use gztemp
appe from temp
   use dw&JB1
   appe from dw&NY
    go bottom
    repl �·� with right(NY,2)
    repl a with yrs , ���� with '&NY' , YLJE with HJ , JTHJ with GLGZ+ZYBF+HT+ZF+JTBT+SBF+SDBT+MT+XLF+GWJT+sfGZ+kk+STJT+BJ
    do case
    case right(����,2)='03'
      total on DWDM to dw1 for val(�·�)>=1 and val(�·�)<=3
      select 1
      use dw&JB1
      append blank
      select 2
      use dw1
      for I = 1 to fcount()
        FIELDNAME = alltrim(field(I))
        select 1
        if empty(field(I))
          exit
        endif
        repl &fieldname with b.&fieldname
      endfor
      replace ���� with 'һ����' , A with int(A/3) , �·� with ''
    case right(����,2)='06'
      total on DWDM to dw2 for val(�·�)>=4 and val(�·�)<=6
      select 1
      use dw&JB1
      append blank
      select 2
      use dw2
      for I = 1 to fcount()
        FIELDNAME = alltrim(field(I))
        select 1
        if empty(field(I))
          exit
        endif
        repl &fieldname with b.&fieldname
      endfor
      replace ���� with '������' , A with int(A/3) , �·� with ''
    case right(����,2)='09'
      total on DWDM to dw3 for val(�·�)>=7 and val(�·�)<=9
      select 1
      use dw&JB1
      append blank
      select 2
      use dw3
      for I = 1 to fcount()
        FIELDNAME = alltrim(field(I))
        select 1
        if empty(field(I))
          exit
        endif
        repl &fieldname with b.&fieldname
      endfor
      replace ���� with '������' , A with int(A/3) , �·� with ''
    case right(����,2)='12'
      total on DWDM to dw4 for val(�·�)>=10 and val(�·�)<=12
      select 1
      use dw&JB1
      append blank
      select 2
      use dw4
      for I = 1 to fcount()
        FIELDNAME = alltrim(field(I))
        select 1
        if empty(field(I))
          exit
        endif
        repl &fieldname with b.&fieldname
      endfor
      replace ���� with '�ļ���' , A with int(A/3) , �·� with ''
      total on DWDM to dw for right(allt(����),4)='����'
      select 1
      append blank
      select 3
      use dw
      for I = 1 to fcount()
        FIELDNAME = alltrim(field(I))
        select 1
        if empty(field(I))
          exit
        endif
        repl &fieldname with c.&fieldname
      endfor
      replace ���� with '��  ��' , A with int(A/4) , �·� with ''
    endcase
endfor
close all
***************************
use gztemp
index on rybh to gztemp 
TOTAL ON rybh TO GZ&LY
use gz&LY
INDEX on cjdm TO cjgz
tota on cjdm TO cjgz
INDEX on bmbh TO bmgz
tota on bmbh TO bmgz
USE cjgz
COPY TO d:\���乤�� TYPE xl5
USE bmgz
COPY TO d:\���Ź��� TYPE xl5
use dw&JB1 
copy to temp for !empty(����)
use temp excl
delete for a=0
pack
go top
brow titl '����ˡ����䲿�Ź���̨�ʡ��������ɵ��ӱ�D��\����̨�ʡ�'
do case
case DH=1
    copy to d:\����̨�� fiel ����,hj,bzgz,jbgz,jang,ZYBF,HT,ZF,GWJT,sfGZ,kk,STJT,BJ,ylbx,sybx,ybx,gjj,sds,sfje,a  type xl5
case DH=2
    copy to d:\����̨�� fiel ����,cjmc,hj,bzgz,jbgz,jang,ZYBF,HT,ZF,GWJT,sfGZ,kk,STJT,BJ,ylbx,sybx,ybx,gjj,sds,sfje,a  type xl5
case DH=3
    copy to d:\����̨�� fiel ����,cjmc,bmmc,hj,bzgz,jbgz,jang,ZYBF,HT,ZF,GWJT,sfGZ,kk,STJT,BJ,ylbx,sybx,ybx,gjj,sds,sfje,a  type xl5
endcase
close all
yn = MESSAGEBOX("���ݵ����ɹ����ִ�ʹ����",4+32,"��ʾ")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("d:\����̨��")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("d:\���乤��")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("d:\���Ź���")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
ENDIF
*************************
close all
delete files cjgz.*
delete files bmgz.*
delete files dw*.*
delete files temp.*
delete files gztemp.*
return
