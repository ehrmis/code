* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
*  文件名: LY_YLBX.PRG <-- 本文件由 UnFoxAll 创建
* -=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
 CLOSE ALL
 ON KEY LABEL F1 do grcx
 set safety off
 DH = ''
 XM = ''
 DYCJ = ''
 DYBM = ''
 LY=ND
 IF  .NOT. FILE('YL' + LY + '.dbf')
    USE ylbxk
     copy to yl&LY stru
     use yl&LY
    IF FILE('zr' + LY + '.dbf')
        append from zr&LY
    ELSE 
        MESSAGEBOX('请先进行' + LY + '年工资总额处理！！！','提示')
       RETURN 
    ENDIF 
    REPLACE 年份 WITH (LY)
    REPLACE ZBH WITH ALLTRIM(CJDM) + '000' + STR(RECNO(),1) FOR RECNO() < 10
    REPLACE ZBH WITH ALLTRIM(CJDM) + '00' + STR(RECNO(),2) FOR  ;
         RECNO() >= 10 AND RECNO() < 100
    REPLACE ZBH WITH ALLTRIM(CJDM) + '0' + STR(RECNO(),3) FOR  ;
         RECNO() >= 100 AND RECNO() < 1000
    REPLACE ZBH WITH ALLTRIM(CJDM) + STR(RECNO(),4) FOR RECNO() >= 1000 AND RECNO() < 10000
 ENDIF 
 DO DHKDY
************************通用功能：重号检查***********
use yl&LY
repl a with 1 all
sort on rybh to temp
use temp excl
go top
do while !eof()
   bh=rybh
   skip
   if bh=rybh
      repl a with 0
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a=0 to aaa
if aaa>0
   inde on bmbh+rybh to temp
   go top
   brow for a=0 titl '发现'+allt(str(aaa,4))+'人重号！请修改正确或按[ Ctrl+T ]删除多余记录'
   pack
   delete file yl&LY..dbf
   sort on bmbh,rybh to yl&LY
endif
close all
use yl&LY excl
 DO CASE 
 CASE DH = 2
     set filter to CJdm='&dycj'
 CASE DH = 3
     set filter to bmbh='&dybm'
 ENDCASE 
 inde on bmbh+rybh to yl&ly
 GO TOP
 for zdz=1 to fcount()
  zd=fiel(zdz)
  zdz1=(fsize(zd)+fsize(zd))*2.4
  if uppe(allt(zd))='RYXM' 
     exit
  endif
  endfor   
 BROWSE part zdz1 FIELDS CJMC :H = '车间名称' : 12 :R , BMMC :H = '部门名称' : 12 :R ,  ;
      RYXM :H = '  姓名  ' : 8 :R , 缴费月数 : 8 :H = '缴费月数' , 缴费基数  ;
      : 8 :H = '年缴费基数' , 单位缴费 : 12 :H = '当年单位缴费' , 单位利息 : 12  ;
      :H = '当年单位利息' , 个人缴费 : 12 :H = '当年个人缴费' , 当年利息 :  ;
      12 :H = '当年个人利息' , 当年合计 : 10 :H = '当年合计' , 历年单位缴 :  ;
      12 :H = '历年单位缴费' , 历年利息1 : 12 :H = '历年单位利息' ,  ;
      历年个人缴 : 12 :H = '历年个人缴费' , 历年利息2 : 12 :H = '历年个人利息' ,  ;
      总累计 : 10 :H = '总    计' TITLE  ;
      LY + '年个人帐户输入[ F1 = 查找人员    Esc = 退出 ]'
 SET FILTER TO
 CLOSE ALL
USE deset
LOCATE FOR oop='backup'
I=seted
  use yl&LY excl
  pack
 SORT ON BMBH , RYBH TO temp
 USE temp
  copy to yl&LY
 CLOSE ALL
use yl&LY
WHNY=LY+'年养老保险'
DBFF1='yl'
do Qdir
close all
 DELETE File temp.dbf
 ON KEY LABEL F1 
retu
*
