*************************ͨ�ù��ܣ��غż��***********
if empty(dbff1)
   dbff1='&SJLJ.data\ryzk'  
endif
use '&dbff1' excl 
repl a with 1 all
F_LIST = ''
        for iii = 1 to fcount()
        if iii>1
          F_LIST = alltrim(F_LIST)+','+rtrim(field(iii))
        else
          F_LIST = alltrim(field(1))+alltrim(F_LIST)+','+rtrim(field(iii))
        endif
      endfor
WHNY=''
go top
do while .T.
getexpr '�����������µ����ݣ��ֶΣ�:' to WHNY
WHNY=allt(WHNY)
IF  "'"$WHNY or '"'$WHNY or ','$WHNY
     MESSAGEBOX('���������в��ܺ�'+"'"+'"'+','+'�ַ���','��ʾ')
else
     if alltrim(upper(WHNY))$alltrim(F_LIST) AND !EMPTY(WHNY)
       exit
    ELSE
       IF !EMPTY(WHNY)
          MESSAGEBOX('�������ݲ�������ѡ���ݿ��У�','��ʾ')
       ELSE
          EXIT
       ENDIF     
    endif
ENDIF
enddo
sort on &whny to temp
use temp excl
     go top
     do while !eof()
        bh=&whny
        skip
        if bh=&whny
           repl a with 0
           skip-1
           repl a with 0
           skip
        endi
     enddo
coun for a=0 to aaa
if aaa>0
   go top
   if uppe(whny)='C'
      brow for a=0 titl '����'+allt(str(aaa,4))+'�����������޸���ȷ��ɾ�������¼' 
   else
      brow for a=0 titl '����'+allt(str(aaa,4))+'���غţ����޸���ȷ��ɾ�������¼' 
   endif
   pack
   copy to &dbff1  
else
if uppe(whny)='C'
   =MESSAGEBOX("û�з�����������",48,"��ʾ")
else
   =MESSAGEBOX("û�з������غţ�",48,"��ʾ")
endif
endif
use '&dbff1' excl 
repl a with 1 all
go top