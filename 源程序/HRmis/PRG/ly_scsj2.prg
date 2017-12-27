***************************
* .\LY_SCSJ.PRG
***************************
if file('kq'+LY+'.dbf')
yn = MESSAGEBOX('删除'+alltrim(LY)+'年失业保险文件(SY'+alltrim(LY)+'.DBF)？',4+32,"提示")
IF yn = 6
    delete file sy&LY..dbf
    delete file sy&LY..idx     
=MESSAGEBOX("数据删除成功！",48,"初始化")
endif
*xn=str(val(LY)+1,4)
*if file('zr1'+LY+'.dbf')
*yn = MESSAGEBOX('删除'+alltrim(xn)+'年缴费基数文件(ZR1'+alltrim(LY)+'.DBF)？',4+32,"提示")
*IF yn = 6
 *   delete file zr1&LY..dbf
  *  delete file zr1&LY..idx
   * delete file gzze&LY..dbf
    *delete file gzze&LY..idx
*=MESSAGEBOX("数据删除成功！",48,"初始化")
*endif
*endif
else
=MESSAGEBOX("数据不存在！",48,"提示")
endif