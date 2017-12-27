***************************
* .\LY_SCSJ.PRG
***************************
if file('yl'+LY+'.dbf') or file('sy'+LY+'.dbf')
yn = MESSAGEBOX('删除'+alltrim(LY)+'年养老保险文件(YL'+alltrim(LY)+')？',4+32,"提示")
IF yn = 6
    delete file yl&LY..dbf
    delete file yl&LY..idx
    delete file ylbx&LY..dbf
    delete file ylbx&LY..idx
    delete file lnk&LY..dbf
    delete file lnk&LY..idx
=MESSAGEBOX("数据删除成功！",48,"恭喜")
endif
if file('sy'+LY+'.dbf')
yn = MESSAGEBOX('删除'+alltrim(LY)+'年失业保险文件(SY'+alltrim(LY)+'.DBF)？',4+32,"提示")
IF yn = 6
    delete file sy&LY..dbf
    delete file sy&LY..idx     
=MESSAGEBOX("数据删除成功！",48,"恭喜")
endif
endif
else
=MESSAGEBOX("数据不存在！",48,"提示")
endif