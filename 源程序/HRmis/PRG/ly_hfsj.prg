***************************
* .\LY_HFSJ.PRG
***************************
set safety off
 CLOSE ALL
 YN0 = MESSAGEBOX('��Ҫ�ָ��Զ��屸�������𣿣�...�������̻�ѡ���Զ����ļ��У�',291,'��ʾ')
 DO CASE 
 CASE YN0 = 6
 old_dirs='d:'
Declare integer Qdir_m In Qdir.dll string @ck,;
integer yy,;
string cc
*ע��Qdir_m()��@ck=���ص��ļ���·����yy=�ļ������,cc=��ʾ
*==========================================================
ck=spac(64)      &&����ļ���·��
yy=1
cc='��ѡ���Ҫ�ָ��ļ��ı���Ŀ¼,����ȷ����ť��'
xx=Qdir_m(@ck,yy,cc)   
*����Qdir_m()��ѡ��Ľ����ck�з���
dirs=allt(subs(ck,2,64))
if len(dirs)=0
   dirs=old_dirs
endif
if right(dirs,1)<>'\'
   dirs=dirs+'\'
endif
if file(dirs+'YL'+ly+'.dbf')
yn = MESSAGEBOX('�ָ�'+alltrim(LY)+'�����ϱ����ļ�('+dirs+'YL'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
IF yn = 6
  wait window '���ڻָ����ϱ��տ�... ...' nowait
    use &dirs.YL&ly
    copy to YL&LY
=MESSAGEBOX("���ݻָ��ɹ�!",48,"��ϲ")
endif
endif
if file(dirs+'zr'+ly+'.dbf') 
yn = MESSAGEBOX('�ָ�'+alltrim(LY)+'�깤���ܶ��ļ�('+dirs+'ZR'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
    IF yn = 6
   wait window '���ڻָ������ܶ��... ...' nowait
    if file(dirs+'zr'+ly+'.dbf')
       use &dirs.ZR&ly
       copy to ZR&LY
    endif
    endif
endif
if file(dirs+'zr1'+ly+'.dbf')
yn = MESSAGEBOX('�ָ�'+str(val(alltrim(LY)+1),4)+'��ɷѻ����ļ�('+dirs+'ZR'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
IF yn = 6
    if file(dirs+'zr1'+ly+'.dbf')
       use &dirs.ZR1&ly
       copy to ZR1&LY
    endif
    =MESSAGEBOX("���ݻָ��ɹ�!",48,"��ϲ")
endif
endif
 CASE YN0 = 7
   if file('backup\yl'+ly+'.dbf')
    yn = MESSAGEBOX('�ָ����ݿ���'+alltrim(LY)+'����ᱣ���ļ�(backup\YL'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
IF yn = 6
  wait window '���ڻָ����ϱ��տ�... ...' nowa
  use backup\yl&ly
  copy to sy&LY
   =MESSAGEBOX("���ݻָ��ɹ�!",48,"��ϲ")
  wait window '���ڻָ�ʧҵ���տ�... ...' nowa
  use backup\sy&ly
  copy to sy&LY
   =MESSAGEBOX("���ݻָ��ɹ�!",48,"��ϲ")
endif 
   endif 
   if file('backup\zr'+ly+'.dbf')
    yn = MESSAGEBOX('�ָ����ݿ���'+alltrim(LY)+'�깤���ܶ��ļ�(backup\ZR'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
IF yn = 6
  wait window '���ڻָ������ܶ��... ...' nowa
  use backup\zr&ly
  copy to zr&LY
   =MESSAGEBOX("���ݻָ��ɹ�!",48,"��ϲ")
endif
   endif
   xn=str(val(LY)+1,4)
   if file('backup\zr1'+ly+'.dbf')
    yn = MESSAGEBOX('�ָ����ݿ���'+alltrim(xn)+'��ɷѻ����ļ�(backup\ZR1'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
IF yn = 6
  wait window '���ڻָ��ɷѻ�����... ...' nowa
  use backup\zr1&ly
  copy to zr1&LY
   =MESSAGEBOX("���ݻָ��ɹ�!",48,"��ϲ")
endif 
   endif    
 CASE YN0 = 2
    RETURN 
 ENDCASE