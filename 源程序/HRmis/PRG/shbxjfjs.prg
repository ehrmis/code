*********************************
* .\shbxJfjs.PRG(Build 20160621)
*********************************
close all
nd1=STR(VAL(nd)-1,4)
on key label F1 do grcx
if  not file('zr1'+nd1+'.DBF')
   =MESSAGEBOX(nd1+'年度工资总额不存在,处理后再更新！',48,"提示")
  return
else    
use zr1&nd1
go top
browse field  CJMC:h='车间',bmmc:h='班组',RYXM:h='姓名',HJ:h=nd1+'年工资总额',mou:h='缴费月数',PJ:h='月平均',工资外收入,校正:h='缴费工资',jfjs:h=nd+'年缴费基数',ZR3:h='养老保险',qynj:h='企业年金',sybx:h='失业保险',ybx:h='医疗保险',gjj:h='住房公积金',HEJI:h='个人月缴纳合计' noedit titl nd+'年职工个人"五险两金"月缴费情况 [按 F1 键搜索某一职工]'
endif
close all
use sbdmk
更新月=更新月份
IF 人可='更新'
         IF month(date())<>更新月
      =MESSAGEBOX('对不起！系统设置为'+str(更新月,2)+'月份更新社会保险缴费基数,现在是'+str(month(date()),2)+'月！',48,"提示")
         retu
         ELSE
           do SBJS_GX
         ENDIF
yn = MESSAGEBOX("个人缴费数据导出成功！现打开使用吗？",4+32,"提示")         
FileName = GETFILE('XLS', '导出文件名: ', '确定')
IF EMPTY(alltrim(FileName)) 
    FileName='&xp'+'\个人缴费.xls'
ENDIF 
wjm=ALLTRIM(FileName)
**************标准Excel读取文件*************      
         use zr1&nd1
         COPY TO &wjm FIELDS cjmc,bmmc,ryxm,jfjs,zr3,qynj,sybx,ybx,gjj TYPE XL5        
         close all
IF yn = 6
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&wjm")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表"
ENDIF
人可=''
ENDIF


proc SBJS_GX

close all
wait window '正在更新'+nd+'年职工个人社会保险缴费金额......' nowait
select 2
use zr1&nd1
index on RYBH to zr1&nd1
go top
select 1
use ryzk
repl ylbx with 0 all
*****用于处理本年度新增人员***
index on RYBH to ryzk
YN = MESSAGEBOX('需要更新'+str(更新月,2)+'月份医疗保险吗？',4+32,'提示')
 if YN = 6
    YN1 = MESSAGEBOX('需要一次性代扣特种病保险(12元/人.年)吗？',4+32,'提示')
    if YN1 = 6
       upda on rybh from B replace YBJS with b.JFJS , JFJS with b.jfjs , YLBX with b.ZR3 ,qynj with b.qynj,SYBX with b.sybx , gjj with b.gjj,YBX with b.YBX+12      
    else   
       upda on rybh from B replace YBJS with b.JFJS , JFJS with b.jfjs , YLBX with b.ZR3,qynj with b.qynj, SYBX with b.sybx , gjj with b.gjj,YBX with b.YBX
    endif
else 
   upda on rybh from B replace YBJS with b.JFJS , JFJS with b.jfjs , YLBX with b.ZR3 ,qynj with b.qynj, SYBX with b.sybx,gjj with b.gjj   
endif
close all
use ryzk
index on bmbh+RYBH to ryzk
go top
brow fiel cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',jfjs:h='月缴费基数' , ylbx:h='养老保险',qynj:h='企业年金',sybx:h='失业保险',ybx:h='医疗保险',gjj:h='住房公积金' noedit titl '&nd'+'年五险两金缴费情况 [ 按 F1 查询个人 ]'
CLOSE ALL
wait window '正在备份'+str(year(date())-2,4)+'年～'+str(year(date()),4)+'年社保数据到“BACKUP”子目录中... ...' nowait
use sbdmk
DMND=allt(dw)
for BF=year(date())-2 to year(date())
    BFND=str(BF,4)
    BFND1=DMND+right(str(BF,4),2)
   if files("ZR"+BFND+".DBF")
      use zr&BFND
      copy to backup\zr&BFND1 
   endif 
   if file('zr1'+BFND+'.dbf')
      use zr1&BFND
      copy to backup\zr1&BFND1  
   endif 
   if file('yl'+BFND+'.dbf')
      use yl&BFND
      copy to backup\yl&BFND1 
   endif 
   if file('sy'+BFND+'.dbf')
      use sy&BFND
      copy to backup\sy&BFND1 
   endif 
endfor
 =MESSAGEBOX(nd+"年社会保险个人缴费金额更新成功！",48,"恭喜")
return