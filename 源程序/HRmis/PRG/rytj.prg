close all
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
use data\dmk
bf=LEFT(ALLTRIM(�̷�),3)
wait window '��  ��  ��  �ݡ���, ��  ��  ��... ...' nowait
use ��Ա�䶯 
copy to temp stru
use temp &&��ǰ��Ա
appe from ryzk
close all
sele 2
use rybdtemp &&���»��ϴμ����Ա,ÿ����һ�θ���һ��
repl a with 1 all
inde on rybh to rybdtemp
sele 1
use temp
repl a with 0 all
inde on rybh to temp
upda on rybh from B repl a with b.a 
use ��Ա�䶯 excl
zap
appe from temp for a=0
repl �䶯��� with '����',a with 1 all
close all
sele 2
use temp
repl a with 1 all
inde on rybh to temp
sele 1
use rybdtemp
repl a with 0 all
inde on rybh to rybdtemp
upda on rybh from B repl a with b.a 
use ��Ա�䶯 
appe from rybdtemp for a=0
repl �䶯��� with '����' for a=0
repl a with 1 all                          
use rybdtemp
repl a with 1 all
go top
scan
sele 2
use temp
 WAIT WINDOW NOWAIT '���ڼ�����Ա�䶯���'+A.cjdm+'��'+allt(A.cjmc)+'��'+allt(A.bmmc)+'��'+allt(A.ryxm)+'... ...' 
 repl a with 0 ,�䶯��� with '��λ�䶯'for dwdm<>A.dwdm and rybh= A->rybh
 repl a with 0 ,�䶯��� with '����䶯'for cjdm<>A.cjdm and rybh= A->rybh
 repl a with 0 ,�䶯��� with '���ű䶯'for bmbh<>A.bmbh and rybh= A->rybh
 repl a with 0 ,�䶯��� with 'ְ��䶯'for zw<>A.zw and rybh= A->rybh
 repl a with 0 ,�䶯��� with '�ڸڱ䶯'for zgqk<>A.zgqk and rybh= A->rybh  
 repl a with 0 ,�䶯��� with '��λ�䶯'for gz<>A.gz and rybh= A->rybh  
endscan
use ��Ա�䶯 
appe from temp for a=0
appe from temp for year(��ڼƻ�)=year(date()) and month(��ڼƻ�)=month(date())
repl a with 0 all
count for a=0 to jls
if jls>0 
inde on bmbh+rybh to ��Ա�䶯
for zdz=1 to fcount()
  zd=fiel(zdz)
  zdz1=(fsize(zd)+fsize(zd))*2
  if uppe(allt(zd))='RYXM' 
     exit
  endif
endfor   
go top
browse part zdz1 field dwmc:h='��λ',cjmc:h='����',bmmc:h='����',ryxm:h='����',zw1:h='ְ��',zgqk1:h='�ڸ����',�䶯���,��ڼƻ�,���ʱ�� titl '[�� F1 �����] [�� Esc ����] ��������˺�ά��'
endif
close all
********************
use ryzk
repl a with 1 all
copy to temp1 for zgqk='01'
copy to temp2 for !empty(bmmc) and zgqk='01'
copy to temp3 for left(zzdm,2)='02' and zgqk='01'
if jls>0
use ��Ա�䶯
copy to &bf.��Ա�䶯��� type xl5
 =messagebox("���¸���λ����ͳ�Ƶ�������"+&bf+"��Ա�䶯��������ӱ��У���",48,"��ϲ")
endif
use temp1
inde on cjdm to temp1
tota on cjdm to temp
use temp
copy to &bf.����λ����ͳ�� fiel cjmc,a type xl5
 =messagebox("���¸���λ����ͳ�Ƶ�������"+&bf+"����λ����ͳ�ơ����ӱ��У���",48,"��ϲ")
use temp1
inde on dwdm+cjdm to temp1
tota on dwdm+cjdm to temp
use temp
copy to &bf.����������ͳ�� fiel dwmc,cjmc,a type xl5
 =messagebox("���¸���λ����ͳ�Ƶ�������"+&bf+"����������ͳ�ơ����ӱ��У���",48,"��ϲ")
use temp2
inde on dwdm+cjdm+bmbh to temp2
tota on dwdm+cjdm+bmbh to temp
use temp
copy to &bf.����������ͳ�� fiel dwmc,cjmc,bmmc,a type xl5
 =messagebox("���¸���������ͳ�Ƶ�������"+&bf+"����������ͳ�ơ����ӱ��У���",48,"��ϲ")
use temp3
inde on xb1 to temp3
tota on xb1 to temp
use temp
copy to &bf.������λ�Ա���� fiel xb1,a type xl5
 =messagebox("���¸���������ͳ�Ƶ�������"+&bf+"������λ�Ա���ܡ����ӱ��У���",48,"��ϲ")
use temp3
inde on cjdm+bmbh+xb1 to temp3
tota on cjdm+bmbh+xb1 to temp
use temp
copy to &bf.������λ�Ա�ͳ�� fiel cjmc,bmmc,xb1,a type xl5
 =messagebox("���¸���������ͳ�Ƶ�������"+&bf+"������λ�Ա�ͳ�ơ����ӱ��У���",48,"��ϲ")
use ryzk
copy to temp for val(zw)<11 and zgqk='01'
use temp
inde on dwdm+cjdm+bmbh+zw to temp
tota on dwdm+cjdm+bmbh+zw to temp1
use temp1
copy to &bf.ְ��ͳ�� fiel dwmc,cjmc,bmmc,zw1,a type xl5
use ryzk
copy to temp fiel dwmc,cjmc,bmmc,zgqk,zgqk1,ryxm,zw1 for zgqk<>'01'
use temp
count for zgqk<>'01' to jls1
if jls1>0 
copy to &bf.�ڸ���� type xl5
 =messagebox("����ְ��ͳ�Ƶ�������"+&bf+"�ڸ���������ӱ��У���",48,"��ϲ")
endif
close all
yn = MESSAGEBOX("���ݵ����ɹ����ִ�ʹ����",4+32,"��ʾ")
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"��Ա�ֲ�ͳ�Ʊ�")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
if jls>0
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"��Ա�䶯���")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
endif
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"����λ����ͳ��")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"����������ͳ��")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"����������ͳ��")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"������λ�Ա����")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"������λ�Ա�ͳ��")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"ְ��ͳ��")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
if jls1>0 
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"�ڸ����")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
endif
ENDI  
use ��Ա�䶯 excl
copy to rybdtemp stru
use rybdtemp &&����Ϊ��ǰ��Ա
IF FILE('ry'+ny1+'.dbf')
   appe from ry&ny1
else   
   appe from ryzk 
ENDIF
delete file temp.dbf
delete file temp1.dbf