*******************************
* ���ϱ��սɷѻ���.prg(20151124)
*******************************
close all
set safety off
clear
_SCREEN.PICTURE = ''
pd='N'
temp='jfjstemp'
set talk off
@ 1,22 say '�Զ���дϵͳ��ɷѹ����걨����ٽ������' font '����',20
@ 4,3 say '��һ��������˳�����'+ly+'�깤���ܶ�' font '����',13
@ 6,3 say '�ڶ����������籣�������ġ��ɷѻ����걨�������̣����鱸�ݺá������걨�����ӱ�'font '����',13
@ 8,3 say '��������ϵͳ�Զ������籣�������������е�Ӧ��д����' font '����',13
@ 10,3 say '���Ĳ����˳�ϵͳ����ϱ��籣���������ļ����������걨�����ӱ�[.XLS]�����ά������ע����' font '����',12
@ 12,3 say '���³ɹ���γ��ϱ���.....�ٺ٣����̾�����ô�򵥣���' font '����',13
yn = MESSAGEBOX("�Զ���дϵͳ��ɷѹ����걨����",4+32,"��ʾ")
if yn=7
close all  
use zr1&ly
COPY TO &bf.\�����걨 FIELDS erprybh,scbh,ryxm,xb1,�ɷѹ���,jfjs TYPE XL5 
********************�˹�ά�����籣��������������************************************************
ELSE
close all
=MESSAGEBOX("��ѡ���籣�������ġ������걨�����ӱ�",48,"�Զ���д�걨����")
old_dirs='&xp'+'\�����걨.xls'
dirs=GETFILE("xls","�����걨���ӱ�")
if empty(dirs)
   dirs=old_dirs
endif
bakfile=allt(dirs)
if upper(right(bakfile,3))='XLS'
   use �����걨 excl
   zap 
   APPEND FROM &bakfile TYPE xl8 
   brow titl '�����籣�������������������ϱ���ԭʼ���ݣ�'  
else   
   =MESSAGEBOX("���ݵ���ʧ�ܣ�",48,"��ʾ")
   close all
   clear 
   retu
endif        
*****************����ά����Ĭ��ʹ����Ա״�����еġ����˱�š�***************
close all
sele 2
use ryzk
inde on rybh to ryzk
sele 1
use zr1&ly
repl scbh with space(10) all
inde on rybh to zr1&ly
upda on rybh from B repl scbh with b.scbh
go top
fd = MESSAGEBOX("ʹ���������ƽ�����ʷⶥ���״�����",4+32,"��ʾ")
******************�Զ����»���**************
close all
sele 2
use �����걨 excl
*******���ɷѻ����걨�����ӱ�***A���B���˱��C����D�������֤����E�¹���F��ע*********20151123***************               
repl e with '' all
*****�ϱ��������ӱ������������**********
i=1
 scan 
      sele 1
      use zr1&ly
      loca for val(scbh)=val(b.b) 
       wait window '�����ø��˱�Ź�����д��'+allt(str(i))+'/'+allt(str(recc()))+'�ˡ�'+ALLTRIM(Str(i/recc()*100))+'%���ɷѻ���,���Ժ�......' nowait
      if fd=6
      repl b.e with ALLTRIM(str(jfjs))         
         *********�ϱ�����*******
      else
      repl b.e with ALLTRIM(str(pj))        
        *********�걨����*******    
      endif
       i=i+1 
  endscan
close all
use �����걨 excl
count for empty(e) and len(allt(b))=10 to rs
if rs>0
  delete FOR empty(e) and len(allt(b))=10
   go top
   brow FOR empty(e) and len(allt(b))=10 fiel b:h='���˱��',c:h='����',d:h='���֤��',e:h='�½ɷѹ���' titl '��ά��������Ա'+nd1+'��ɷѻ���,������Ա����ɾ����'
   yn3 = MESSAGEBOX("ɾ���ޱ�ż�¼��",4+32,"��ʾ")
   if yn3=6
      pack
   else
      recall all
   endif   
endi
go top 
brow fiel  a:h='���',b:h='���˱��',c:h='����',d:h='���֤��',e:h='�½ɷѹ���',f:h='��ע' titl '��������ˡ����ϱ��սɷѻ������������(�ر�������Ա)�������ɵ��ӱ�'+bf+'\�����걨 ��XLS'
*********************��˻���*****************
COPY TO &bf.\�����걨 TYPE XL5 
********************�˹�ά�����籣��������������************************************************
close all  
use zr1&ly
************************ͨ�ù��ܣ��غż��***********
repl a with 1 all
sort on scbh to temp
use temp excl
go top
do while !eof()
   bh=scbh
   skip
   if bh=scbh
      repl a with 0
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a=0 to aaa
if aaa>0
   copy to jfjstemp for a=0
   temp='jfjstemp'
   pd='Y'
   use jfjstemp excl
   inde on bmbh+rybh to jfjstemp
   go top
   brow fiel scbh:h='���ϱ��ո��˱��',cjmc:h='����',bmmc:h='����',ryxm:h='����',sfz:h='���֤��' titl '����'+allt(str(aaa,4))+'�˸��˱���غŻ�պţ����޸���ȷ������'+bf+"\����غš�XLS��"   
   COPY TO &bf.\����غ�  TYPE XL5 FIELDS cjmc,bmmc,ryxm,rybh,sfz,scbh
   ********************�˹�ά������Ա״���ٲ���һ�Σ������ٵ����籣���������������еġ����˱�š���************************************************
endif
close all
ENDIF
clear
yn2 = MESSAGEBOX("�ɷѹ�����д�򵼳��ɹ�������ʹ�õ��ӱ�༭��ӡ��",4+32,"��ʾ")
IF yn2 = 6
  if yn=6
     myexcel=CREATEOBJECT("excel.application")
     myexcel.workbooks.open("&bakfile")
     *********����ļ���***ͨ�ñ��ʽ*****
     myexcel.visible=.t.
     myexcel.caption="ʹ�õ��ӱ�༭��ӡ����" 
  endif
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\�����걨.xls")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����" 
ENDI   
close all
clear
use ccrq
go 8
I=val(ͼƬ)
repl ͼƬ with str(I+1,2)
IF I>=54
   repl ͼƬ with '0'
ENDIF
pic=ALLTRIM(ͼƬ)
**********�û�������־�����׽���*********
pict_fd='fd'+pic+'.jpg'
_SCREEN.PICTURE = '&pict_fd'


