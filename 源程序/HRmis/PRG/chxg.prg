*************************通用功能：重号检查***********
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
getexpr '请输入主更新的内容（字段）:' to WHNY
WHNY=allt(WHNY)
IF  "'"$WHNY or '"'$WHNY or ','$WHNY
     MESSAGEBOX('输入内容中不能含'+"'"+'"'+','+'字符！','提示')
else
     if alltrim(upper(WHNY))$alltrim(F_LIST) AND !EMPTY(WHNY)
       exit
    ELSE
       IF !EMPTY(WHNY)
          MESSAGEBOX('输入内容不含在所选数据库中！','提示')
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
      brow for a=0 titl '发现'+allt(str(aaa,4))+'人重名！请修改正确或删除多余记录' 
   else
      brow for a=0 titl '发现'+allt(str(aaa,4))+'人重号！请修改正确或删除多余记录' 
   endif
   pack
   copy to &dbff1  
else
if uppe(whny)='C'
   =MESSAGEBOX("没有发现人重名！",48,"提示")
else
   =MESSAGEBOX("没有发现人重号！",48,"提示")
endif
endif
use '&dbff1' excl 
repl a with 1 all
go top