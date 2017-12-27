* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
*  文件名: GRCX.PRG <-- 本文件由 UnFoxAll 创建
* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
   DO srxm.spr 
   LOCATE FOR ALLTRIM(ryxm) = ALLTRIM(XM) .OR. ALLTRIM(rybh) = ALLTRIM(XM)
    **********************注意：在当前表中查询***********先人工借助系统左下脚提示栏信息确定是否在此表中？再按F1查找*****************
   IF EOF()
       DO dhkjg.spx
       M = READKEY()
       IF M = 12 .OR. M = 268
          EXIT  
       ENDIF 
   ENDIF 
*
