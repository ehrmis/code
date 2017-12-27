USE  培训登记
jls=RECCOUNT()
if jls=0
=MESSAGEBOX('请先在“人员状况”数据维护中进行培训登记……', 16,"提示")
 RETURN
ELSE
********快捷报表中调用“培训登记数据处理”过程，但又不导出电子表数据***20170906************
CLOSE all
SELECT 2 
USE ryzk
REPLACE a WITH 1 all  
INDEX on rybh TO ryzk
SELECT 1
use 培训登记 excl
repl a with 0 all
SET RELATION TO rybh inTO 2
REPLACE dwdm with b.dwdm,dwmc with b.dwmc,cjdm with b.cjdm,cjmc with b.cjmc,bmbh with b.bmbh,bmmc with b.bmmc,xb WITH b.xb,xb1 WITH b.xb1,zw WITH b.zw,zw1 WITH b.zw1,csrq WITH b.csrq,;
nl WITH b.nl,cjgzrq WITH b.cjgzrq,gl WITH b.gl,gz1 WITH b.gz1,gw WITH b.gw,gwfl1 with b.gwfl1,a with b.a FOR rybh=b.rybh
*******必须使用for条件****因为同一个编号有多条记录，不能用UPDATE语句*****UPDATE on rybh from B repl更新语句在A区有多条同号记录时只更新第一条记录****20171130***************
GO top
loca for a=0 
IF !eof()
   delete for a=0 
   brow for a=0 titl '请认真审核下列人员是否删除'
    yn = MESSAGEBOX("需要删除上述标记人员吗？",4+32,"提示")
   IF yn = 6
      PACK
****功能：先提示删除再关联当月人员状况库确定是否有减少人员？******20171130*************
   ELSE
       RECALL all
   ENDIF         
ENDIF
repl a with 1 all
CLOSE ALL
**********多本证书者只更新保存最高级培训(第一行记录)***********  
use 培训登记 excl
inde on 姓名 to 培训登记
DELETE FOR ryxm<>姓名 
*********ryxm可能是自动复制其他人的记录时生成的记录，或者是误输入了一条空记录(只点击了“培训登记”并没有输入任何数据)*****20151115修复“培训登记”Bug***提升系统稳定性**************
COUNT FOR ryxm<>姓名 TO rs
IF rs>0
  GO top
  BROWSE FIELDS 序号,单位,部门,班组,姓名,培训名称,特种作业,项目,徒弟,师带徒工种,起始日期,终止日期,培训等级,证书名称,证书编号,发证单位,发证日期,fszq:h='复审周期',近期复审 FOR ryxm<>姓名 titl '请审核培训情况[按F1查找个人  可使用“表”菜单功能增加或删除冗余记录]'
ENDIF
inde on 序号 to 培训登记 
GO top 
BROWSE FIELDS 序号,单位,部门,班组,姓名,培训名称,特种作业,项目,徒弟,师带徒工种,起始日期,终止日期,培训等级,证书名称,证书编号,发证单位,发证日期,fszq:h='复审周期',近期复审 titl '请审核培训情况[按F1查找个人  可使用“表”菜单功能增加或删除冗余记录]'
repl djdm with '08' for '高级'$培训等级 and gwfl1='操作岗位'
repl djdm with '09' for '中级'$培训等级 and gwfl1='操作岗位'
repl djdm with '10' for '初级'$培训等级 and gwfl1='操作岗位'
*****************操作岗位人员职业资格证处理完后，再对不同岗位人员多重身份人员分类******20170906*****************************************
repl djdm with '99' for gwfl1<>'操作岗位' OR VAL(djdm)=0
repl djdm with '06' for ALLTRIM(培训等级)='高级技师'
repl djdm with '07' for ALLTRIM(培训等级)='技师'
**********************下面处理特殊等级代码***********20170906*********************************
REPLACE 培训等级 WITH '特种作业',djdm with '11',证书名称 with '安全操作证' FOR !EMPTY(特种作业)
REPLACE 培训等级 WITH '师带徒',djdm with '12',培训名称 with '师带徒培训' FOR !EMPTY(师带徒工种)
repl a with 1,证书编号 with allt(证书编号) all
**************培训分类*************20170906**********************************
copy to temp for EMPTY(证书编号) AND VAL(djdm)<>12
************先过滤无编号培训证书***********************
sort on 证书编号 to temp1 for !EMPTY(证书编号)
copy to temp2 for EMPTY(证书编号) AND VAL(djdm)=12
************一分为二先过滤无编号培训证书，清理问题记录(证书重号)*****标准登记法，有证必有编号****20170906**************
use temp1 excl
go top
do while !eof()
   bh=allt(培训名称)+allt(证书编号)+allt(培训等级)
   skip
   if bh=allt(培训名称)+allt(证书编号)+allt(培训等级)
      repl a with 2
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a<>1 to aaa
IF aaa>0
   inde on 证书编号 to temp1
   go top
   DELETE FOR a=2
   brow for a<>1 titl '发现'+allt(str(aaa,4))+'证书重号！请修改正确或删除多余记录' 
   yn = MESSAGEBOX("是否删除重号人员？",4+32,"提示")
IF yn = 6
PACK
ELSE
RECALL
ENDIF
ENDIF
CLOSE all
use temp
********无证号*************
appe from temp1
********有证号*******************
appe from temp2
********************“师带徒”记录***********20170906**********************
sort on rybh,djdm to temp3
use temp3 excl
delete for empty(培训名称)
***********没登记培训内容*******20170906****************
pack 
use 培训登记 excl 
zap
appe from temp3
********等级排序更新人员状况库******20170906***************
REPLACE 单位 with ALLTRIM(dwmc),部门 with ALLTRIM(cjmc),班组 with ALLTRIM(bmmc),pxmc with ALLTRIM(培训名称),pxdj with ALLTRIM(培训等级),a with 1 all
*********要求：一人多证按最高级更新RYZK.DBF*********************
inde on 姓名 to 培训登记
GO top
BROWSE FIELDS 姓名,单位,部门,班组,培训名称,特种作业,项目,徒弟,师带徒工种,起始日期,终止日期,培训等级,证书名称,证书编号,发证单位,发证日期,复审周期,近期复审 TITLE '请认真审核员工培训登记情况'
CLOSE ALL
SELECT 2 
USE ryzk  
go top
do while !eof()
   SELECT 1
   use 培训登记    
   go top
   loca for rybh=b.rybh and gwfl1='操作岗位' and VAL(djdm)<11
   ***************培训信息保存在人员状况库中**************“特种作业（11）”用“师带徒（12）”培训除外********20170906**********************************
   if !eof()
      repl b.pxdj with ALLTRIM(pxdj),b.pxmc with ALLTRIM(pxmc),b.证书编号 with ALLTRIM(证书编号),b.发证单位 with ALLTRIM(发证单位),b.发证日期 with 发证日期,b.zcrq with 发证日期
**********多本证书操作岗位人员只更新保存最高级职业资格培训(第一行记录)***********20151115*******操作岗位的职业资格等级证只有发证日期，也就是“职称/资格认定日期”*****20160816***   
   endif
  sele 2
  wait window NOWAIT '正在保存培训登记......' 
  skip 
enddo 
CLOSE all
USE ryzk 
repl zcdj1 with allt(pxdj) for gwfl1='操作岗位'
REPLACE zc with LEFT(zcdj1,4)+allt(pxmc) for gwfl1='操作岗位' AND !'技师'$pxdj
REPLACE zc with allt(pxmc)+LEFT(zcdj1,8) for gwfl1='操作岗位' AND '技师'$pxdj
REPLACE zc WITH ALLTRIM(zc) all
SELECT 2
USE jsdj
scan
  select 1
  USE ryzk
  wait window NOWAIT '正在保存职业资格等级......' 
  repl zcdj with b.code for ALLTRIM(zcdj1)=ALLTRIM(b.name)
ENDSCAN
CLOSE ALL
USE 培训登记
SORT ON bmbh,rybh,djdm TO temp
********按等级代码排序****20170906********************
USE deset
LOCATE FOR oop='backup'
I=seted
USE temp
REPLACE 序号 with str(recn(),5),单位 with ALLTRIM(dwmc),部门 with ALLTRIM(cjmc),班组 with ALLTRIM(bmmc),a WITH 1,人数 WITH 1 all
COPY TO data\培训登记
********细节：还原母库*************
DBFF1='培训登记'
WHNY=ND+'培训登记'
do Qdir
********细节：提示备份*************
CLOSE ALL
ENDIF