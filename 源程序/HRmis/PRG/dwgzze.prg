******************************
* .\DWGZZE.PRG(Build2007.4.13)
******************************
ny=''
JB1 =LY
use dwgzze excl
zap
copy to dw&JB1 
copy to gztemp
n1=LY+'01'
ni=val(n1)
YF0=iif(month(date())>9,str(month(date()),2),'0'+str(month(date()),1))
n2=allt(iif(val(LY)<year(date()),LY+'12',LY+YF0))
nii=val(n2)
for ny11=ni to nii
  ny=allt(str(ny11))
  wait window '������д'+LY+'��'+str(val(right(NY,2)),2)+'�¹����ܶ�̨��.....'  nowait
  USE DWGZZE excl
  if file('GZ'+NY+'.DBF')
     APPE FROM GZ&NY fiel hj,bzgz,jbgz,jang,ZYBF,HT,ZF,GWJT,sfgz,kk,STJT,BJ,ylbx,sybx,ybx,gjj,sds,sfje,a,rybh,dwdm
  endif
 repl dwdm with '99' all
  yrs=recc()
  repl a with 0 all
  index on DWDM to dwgzze
  TOTAL ON DWDM TO DW&NY
  copy to temp 
ZAP
use gztemp
appe from temp
   use dw&JB1
   appe from dw&NY
    go bottom
    repl �·� with right(NY,2)
    repl a with yrs , ���� with '&NY' , YLJE with HJ , JTHJ with GLGZ+ZYBF+HT+ZF+JTBT+SBF+SDBT+MT+XLF+GWJT+sfgz+kk+STJT+BJ
    do case
    case right(����,2)='03'
      total on DWDM to dw1 for val(�·�)>=1 and val(�·�)<=3
      appe from dw1
      go bott
      replace ���� with 'һ����' , A with int(A/3) , �·� with ''
      case right(����,2)='06'
      total on DWDM to dw2 for val(�·�)>=4 and val(�·�)<=6
       appe from dw2
       go bott
      replace ���� with '������' , A with int(A/3) , �·� with ''
    case right(����,2)='09'
      total on DWDM to dw3 for val(�·�)>=7 and val(�·�)<=9
       appe from dw3
       go bott
      replace ���� with '������' , A with int(A/3) , �·� with ''
    case right(����,2)='12'
      total on DWDM to dw4 for val(�·�)>=10 and val(�·�)<=12
      appe from dw4
       go bott
      replace ���� with '�ļ���' , A with int(A/3) , �·� with ''
      total on DWDM to dw for right(allt(����),4)='����'
      appe from dw
       go bott
      replace ���� with '��  ��' , A with int(A/4) , �·� with ''
    endcase
endfor
close all
use gztemp
index on rybh to gztemp 
TOTAL ON rybh TO GZ&LY
use dw&JB1 
copy to temp for !empty(����)
use temp
pf='d:\'+LY
copy to &pf.�깤��̨�� type xl5 
 =messagebox("�ѳɹ�������D��\����̨�ʵ��ӱ�",48,"��ϲ")
  yn = MESSAGEBOX("��Ҫʹ�õ��ӱ�༭��ӡ������",4+32,"��ʾ")
IF yn = 6
 myexcel=CREATEOBJECT("excel.application")
           pf1="&pf"
           myexcel.workbooks.open(pf1+"�깤��̨��.xls")
           myexcel.visible=.t.
           myexcel.caption="ʹ�õ��ӱ�༭��ӡ����" 
           release oleapp
           WAIT CLEAR
ENDIF     
close all
delete files dw*.*
delete files temp.*
delete files gztemp.*
return
