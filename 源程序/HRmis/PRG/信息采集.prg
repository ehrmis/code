yn = MESSAGEBOX("�Ƿ����û���Ա��Ϣ�����ɼ���",4+32,"��ʾ")
IF yn = 7
   RETURN
ENDIF
do srny.spx
sn=STR(VAL(nd)-1,4)
NY = ND+YF
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
ENDIF
yf1=1
yf2=VAL(yf)
****************��ʼ��****************
USE ryzk
BROWSE TITLE '���ݲɼ�ȫ������Ӳ�����Ա���ryzk.dbfά������ԣ�1�����մ�����˹�ά����2����������ְ����룻3����������zydm'
DO jszc
DO zy
DO zwdm
*****************
PROCEDURE zwdm
**************ɨ�跨����ά��ְ�����****����ά���������*****20151126*******************
CLOSE all
sele 2
USE ryzk
scan
 sele 1
 use zwdm
 WAIT WINDOW NOWAIT '����ά��'+b.cjdm+'��'+allt(b.cjmc)+'��'+allt(b.bmmc)+'��'+allt(b.ryxm)+'<ְ�����>... ...' 
 loca for ALLTRIM(b.zw1)=allt(name)
 repl b.zw with code     
endscan       

******************
PROCEDURE jszc
******************
**********ɨ�跨����ά�����ָ�λ************
CLOSE all
sele 2
USE ryzk
scan
 sele 1
 use jszc
 WAIT WINDOW NOWAIT '����ά��'+b.cjdm+'��'+allt(b.cjmc)+'��'+allt(b.bmmc)+'��'+allt(b.ryxm)+'<����ְ��>... ...' 
 IF !EMPTY(b.zcbh)
    loca for val(b.zcbh)=val(code)
    repl b.zc with allt(zc)
 ELSE    
    loca for  ALLTRIM(b.zc)=allt(zc)
    repl b.zcbh with code
    ***********����ְ���Զ����*******20151105***************
 ENDIF    
endscan
******************
PROCEDURE zy
******************
**********ɨ�跨����ά����ѧרҵ************
CLOSE all
sele 2
USE ryzk
repl zydm with '',zymc with '',zyfl with '', kbdm with '',kbzy with '',a with 1 all
scan
 sele 1
 use zyfl
 WAIT WINDOW NOWAIT '����ά��'+b.cjdm+'��'+allt(b.cjmc)+'��'+allt(b.bmmc)+'��'+allt(b.ryxm)+'<רҵ����>... ...' 
 ****��һרҵ*******
 loca for b.ygxz='01' and (allt(רҵ����)$allt(b.zy) or allt(רҵ����1)$allt(b.zy) or allt(רҵ����2)$allt(b.zy))
 repl b.kbdm with allt(����),b.kbzy with allt(רҵ����)
 if empty(רҵ����1) and  !empty(רҵ����2)
    repl b.kbdm with allt(����),b.kbzy with allt(רҵ����2)
 endif
 ****�ڶ�רҵ*******
 if !empty(b.zy1)
    loca for allt(רҵ����)$allt(b.zy1) or allt(רҵ����1)$allt(b.zy1) or allt(רҵ����2)$allt(b.zy1)
    repl b.zydm with allt(����),b.zyfl with allt(רҵ����)
    if empty(רҵ����1) and  !empty(רҵ����2)
       repl b.zydm with allt(����),b.zyfl with allt(רҵ����2)
    endif   
 endif 
endscan
repl kbdm with '88' for val(kbdm)=0
repl zydm with '99' for val(zydm)=0
