* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
*  文件名: 导出至EXCEL.PRG <-- 本文件由 UnFoxAll 创建
* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-


 LOCAL I , J , XLSHEET , XLAPP , TMPSHEET
 WAIT WINDOW AT INT(SROWS() / 2) + 2 , INT(SCOLS() / 2) - 15 NOWAIT  ;
      '启动 Microsoft Excel 中...'
 SELECT * INTO ARRAY ARRZGTZKDATA FROM cjk.DBF
 TMPSHEET = GETOBJECT('','excel.sheet')
 XLAPP = TMPSHEET.APPLICATION
 XLAPP.VISIBLE = .T.
 XLAPP.ACTIVEWINDOW.WINDOWSTATE = -4137
  XLAPP.WORKBOOKS.ADD()
 XLSHEET = XLAPP.ACTIVESHEET
 USE ('cjk')
 FOR Z = 1 TO FCOUNT()
    XLSHEET.CELLS( 1 , Z ).VALUE = FIELD(Z)
 ENDFOR 
 XLSHEET.COLUMNS( 'A:AL' ).COLUMNWIDTH = 10
 FOR I = 1 TO ALEN(ARRZGTZKDATA,1)
    FOR J = 1 TO ALEN(ARRZGTZKDATA,2)
       XLSHEET.CELLS( I + 1 , J ).VALUE = ARRZGTZKDATA(I,J)
    ENDFOR 
 ENDFOR 
 CLOSE DATABASES ALL
 WAIT CLEAR
*
