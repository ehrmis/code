* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
*  文件名: FJDY.PRG <-- 本文件由 UnFoxAll 创建
* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
 CLOSE ALL
 define window MYWIN1 from 10 , 50 to 30 , 120  colo 7+/2 shadow title '分级打印:';
 footer '[可用鼠标操作]'
activate window MYWIN1
 @ 1 , 8 SAY '请选择打印范围[可用“←↑↓→”操作]:' font '宋体',13
 STORE 1 TO DH
 @ 5 , 8 GET DH SIZE 2 , 12 , 2 FUNCTION '*h 全  体;车  间;班  组' font '宋体',12 
 READ 
 DO CASE 
 CASE DH = 1
    CLOSE ALL
    clear WINDOW MYWIN1
    RETURN 
 CASE DH = 2
    USE data\cjk
    GO TOP
    DEFINE POPUP MYPOP1 FROM 1 , 0 TO 26 , 75 PROMPT FIELDS CJMC TITLE '车间(科室)' font '宋体',12 SHADOW 
    ON SELECTION POPUP MYPOP1 DEACT POPUP MYPOP1
    ACTIVATE POPUP MYPOP1
    @ 0 , 10 SAY CJMC
    DYCJ = ALLTRIM(CJDM)
    CLOSE ALL
   CLEAR WINDOW MYWIN1
   CLEAR WINDOW MYPOP1
    RETURN 
 CASE DH = 3
    SELECT 1
    USE data\cjk
    GO TOP
    SET COLOR TO W+/G
    DEFINE POPUP MYPOP1 FROM 1 , 0 TO 26 , 75 PROMPT FIELDS CJMC TITLE '车间(科室)'  SHADOW 
    ON SELECTION POPUP MYPOP1 DEACT POPUP MYPOP1
    ACTIVATE POPUP MYPOP1
    @ 0 , 10 SAY CJMC
    DM1 = CJDM
    SELECT 2
    USE data\BMDM
     set filter to left(bmbh,2)='&dm1'
    GO TOP
    DEFINE POPUP MYPOP1 FROM 1 , 0 TO 26 , 75 PROMPT FIELDS BMMC TITLE '部门(班组)'  SHADOW 
    ON SELECTION POPUP MYPOP1 DEACT POPUP MYPOP1
    ACTIVATE POPUP MYPOP1
    @ 0 , 20 SAY BMMC
    DYBM = ALLTRIM(BMBH)
 ENDCASE 
 SET FILTER TO
 CLEAR WINDOW MYWIN1
 CLEAR WINDOW MYPOP1
 CLOSE ALL
 RETURN 
*
