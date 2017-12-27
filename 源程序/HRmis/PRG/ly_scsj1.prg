***************************
* .\LY_SCSJ.PRG
***************************
if file('zr'+LY+'.dbf')
yn = MESSAGEBOX('删除'+alltrim(LY)+'年工资总额文件(ZR'+alltrim(LY)+'.DBF)？',4+32,"提示")
IF yn = 6
    delete file zr&LY..dbf
    delete file zr&LY..idx     
=MESSAGEBOX("数据删除成功！",48,"恭喜")
endif
xn=str(val(LY)+1,4)
if file('zr1'+LY+'.dbf')
yn = MESSAGEBOX('删除'+alltrim(xn)+'年缴费基数文件(ZR1'+alltrim(LY)+'.DBF)？',4+32,"提示")
IF yn = 6
    delete file zr1&LY..dbf
    delete file zr1&LY..idx
    delete file gzze&LY..dbf
    delete file gzze&LY..idx
=MESSAGEBOX("数据删除成功！",48,"恭喜")
endif
endif
if file('xzjs'+LY+'.dbf')
yn = MESSAGEBOX('删除'+alltrim(LY)+'年新增人员缴费基数文件(XZJS'+alltrim(LY)+'.DBF)？',4+32,"提示")
IF yn = 6
    delete file xzjs&LY..dbf
    delete file xzjs&LY..idx     
=MESSAGEBOX("数据删除成功！",48,"恭喜")
endif
endif
else
=MESSAGEBOX("数据不存在！",48,"提示")
endif