*********************************
* .\shbxJfjs.PRG(Build 20160621)
*********************************
close all
nd1=STR(VAL(nd)-1,4)
on key label F1 do grcx
if  not file('zr1'+nd1+'.DBF')
   =MESSAGEBOX(nd1+'��ȹ����ܶ����,������ٸ��£�',48,"��ʾ")
  return
else    
use zr1&nd1
go top
browse field  CJMC:h='����',bmmc:h='����',RYXM:h='����',HJ:h=nd1+'�깤���ܶ�',mou:h='�ɷ�����',PJ:h='��ƽ��',����������,У��:h='�ɷѹ���',jfjs:h=nd+'��ɷѻ���',ZR3:h='���ϱ���',qynj:h='��ҵ���',sybx:h='ʧҵ����',ybx:h='ҽ�Ʊ���',gjj:h='ס��������',HEJI:h='�����½��ɺϼ�' noedit titl nd+'��ְ������"��������"�½ɷ���� [�� F1 ������ĳһְ��]'
endif
close all
use sbdmk
������=�����·�
IF �˿�='����'
         IF month(date())<>������
      =MESSAGEBOX('�Բ���ϵͳ����Ϊ'+str(������,2)+'�·ݸ�����ᱣ�սɷѻ���,������'+str(month(date()),2)+'�£�',48,"��ʾ")
         retu
         ELSE
           do SBJS_GX
         ENDIF
yn = MESSAGEBOX("���˽ɷ����ݵ����ɹ����ִ�ʹ����",4+32,"��ʾ")         
FileName = GETFILE('XLS', '�����ļ���: ', 'ȷ��')
IF EMPTY(alltrim(FileName)) 
    FileName='&xp'+'\���˽ɷ�.xls'
ENDIF 
wjm=ALLTRIM(FileName)
**************��׼Excel��ȡ�ļ�*************      
         use zr1&nd1
         COPY TO &wjm FIELDS cjmc,bmmc,ryxm,jfjs,zr3,qynj,sybx,ybx,gjj TYPE XL5        
         close all
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&wjm")
myexcel.visible=.t.
myexcel.caption="ʹ�õ��ӱ�༭��ӡ����"
ENDIF
�˿�=''
ENDIF


proc SBJS_GX

close all
wait window '���ڸ���'+nd+'��ְ��������ᱣ�սɷѽ��......' nowait
select 2
use zr1&nd1
index on RYBH to zr1&nd1
go top
select 1
use ryzk
repl ylbx with 0 all
*****���ڴ������������Ա***
index on RYBH to ryzk
YN = MESSAGEBOX('��Ҫ����'+str(������,2)+'�·�ҽ�Ʊ�����',4+32,'��ʾ')
 if YN = 6
    YN1 = MESSAGEBOX('��Ҫһ���Դ������ֲ�����(12Ԫ/��.��)��',4+32,'��ʾ')
    if YN1 = 6
       upda on rybh from B replace YBJS with b.JFJS , JFJS with b.jfjs , YLBX with b.ZR3 ,qynj with b.qynj,SYBX with b.sybx , gjj with b.gjj,YBX with b.YBX+12      
    else   
       upda on rybh from B replace YBJS with b.JFJS , JFJS with b.jfjs , YLBX with b.ZR3,qynj with b.qynj, SYBX with b.sybx , gjj with b.gjj,YBX with b.YBX
    endif
else 
   upda on rybh from B replace YBJS with b.JFJS , JFJS with b.jfjs , YLBX with b.ZR3 ,qynj with b.qynj, SYBX with b.sybx,gjj with b.gjj   
endif
close all
use ryzk
index on bmbh+RYBH to ryzk
go top
brow fiel cjmc:h='����',bmmc:h='����',ryxm:h='����',jfjs:h='�½ɷѻ���' , ylbx:h='���ϱ���',qynj:h='��ҵ���',sybx:h='ʧҵ����',ybx:h='ҽ�Ʊ���',gjj:h='ס��������' noedit titl '&nd'+'����������ɷ���� [ �� F1 ��ѯ���� ]'
CLOSE ALL
wait window '���ڱ���'+str(year(date())-2,4)+'�ꡫ'+str(year(date()),4)+'���籣���ݵ���BACKUP����Ŀ¼��... ...' nowait
use sbdmk
DMND=allt(dw)
for BF=year(date())-2 to year(date())
    BFND=str(BF,4)
    BFND1=DMND+right(str(BF,4),2)
   if files("ZR"+BFND+".DBF")
      use zr&BFND
      copy to backup\zr&BFND1 
   endif 
   if file('zr1'+BFND+'.dbf')
      use zr1&BFND
      copy to backup\zr1&BFND1  
   endif 
   if file('yl'+BFND+'.dbf')
      use yl&BFND
      copy to backup\yl&BFND1 
   endif 
   if file('sy'+BFND+'.dbf')
      use sy&BFND
      copy to backup\sy&BFND1 
   endif 
endfor
 =MESSAGEBOX(nd+"����ᱣ�ո��˽ɷѽ����³ɹ���",48,"��ϲ")
return