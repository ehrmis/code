IF I=1
YN = MESSAGEBOX('��Ҫ����'+WHNY+'��...��ѡ���ļ�����Ŀ¼',4+32,'��ʾ')
IF YN = 6
 on error do copyerr 
 old_dirs='backup\'
Declare integer Qdir_m In Qdir.dll string @ck,;
integer yy,;
string cc
*ע��Qdir_m()��@ck=���ص��ļ���·����yy=�ļ������,cc=��ʾ
*==========================================================
ck=spac(64)      &&����ļ���·��
yy=1
cc='ѡ��'+WHNY+'����Ŀ¼����ȡ������Ĭ�ϱ��ݵ�BACKUP��Ŀ¼�У�'
xx=Qdir_m(@ck,yy,cc)   
*����Qdir_m()��ѡ��Ľ����ck�з���
dirs=allt(subs(ck,2,64))
if len(dirs)=0
   dirs=old_dirs
ENDIF
if right(dirs,1)<>'\'
   dirs=dirs+'\'
endif
    use dbf()
    copy to &dirs.&DBFF1.&ND
    if �˿�='ʧ��'
       =MESSAGEBOX("���ݱ���ʧ�ܣ�",48,"��ϲ")
    else
      =MESSAGEBOX("���ݱ��ݳɹ���",48,"��ϲ")
    endif
close all
on error do CCCL with error() , message() , message(1) , sys(16) , lineno()
ENDIF
ELSE
 on error do copyerr 
 old_dirs='backup\'
Declare integer Qdir_m In Qdir.dll string @ck,;
integer yy,;
string cc
*ע��Qdir_m()��@ck=���ص��ļ���·����yy=�ļ������,cc=��ʾ
*==========================================================
ck=spac(64)      &&����ļ���·��
yy=1
cc='ѡ��'+WHNY+'����Ŀ¼����ȡ������Ĭ�ϱ��ݵ�BACKUP��Ŀ¼�У�'
xx=Qdir_m(@ck,yy,cc)   
*����Qdir_m()��ѡ��Ľ����ck�з���
dirs=allt(subs(ck,2,64))
if len(dirs)=0
   dirs=old_dirs
ENDIF
if right(dirs,1)<>'\'
   dirs=dirs+'\'
endif
    use dbf()
    copy to &dirs.&DBFF1.&ND
    if �˿�='ʧ��'
       =MESSAGEBOX("���ݱ���ʧ�ܣ�",48,"��ϲ")
    else
      =MESSAGEBOX("���ݱ��ݳɹ���",48,"��ϲ")
    endif
close all
on error do CCCL with error() , message() , message(1) , sys(16) , lineno()
ENDIF