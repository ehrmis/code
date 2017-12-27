****************************************
* SJCSH.prg( Build 20170505 )
****************************************
close all
set safety off
set talk off
ND=str(year(date()),4)
YF=iif(month(date())>9,str(month(date()),2),'0'+str(month(date()),1))
ny=nd+yf
****全局变量赋值*****20161019*************
set date long
***********************************************************
USE ryzk1 EXCLUSIVE
PACK
USE ryzk EXCLUSIVE
PACK
repl a with 1 all
sort on rybh to temp
close all
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
   inde on rybh to temp
   go top
   brow noedit for a=0 titl '发现'+allt(str(aaa,4))+'人重号，请在“人员状况维护”中修改正确或删除多余记录！' 
endif
repl a with 1 all
close all
delete file temp?.dbf
delete file temp?.idx
******************
    use ryzk 
    COPY TO temp FIELDS dwdm,dwmc,CJDM,CJMC,xb,a
    COPY TO temp1 FIELDS dwdm,dwmc,a
    COPY TO cj FIELDS dwdm,dwmc,CJDM , CJMC , a
    COPY TO bm FIELDS  dwdm,dwmc,CJDM , CJMC ,BMBH , BMMC , a
***********按照维护创建自动生成父子筛选分级汇总及自然状况统计数据******20170104**********      
   USE temp1
    INDEX ON dwdm TO temp1
    TOTAL ON dwdm TO DATA\dwdm
    USE data\dwdm
    单位名称=ALLTRIM(dwmc)
    USE data\dmk
    REPLACE mc WITH 单位名称
    USE temp
    INDEX ON dwdm+CJDM+xb TO temp
    TOTAL ON dwdm+CJDM+xb TO temp1
   close all  
    SELECT 1
    USE EXCLUSIVE xb
    ZAP
    SELECT 2
    USE temp1
    GO TOP
    DO WHILE  .NOT. EOF()
       SELECT 1
       APPEND BLANK
       REPLACE dwdm WITH B.dwdm , dwmc with b.dwmc , cjdm WITH B.CJDM , cjmc WITH B.CJMC , code WITH B.xb , num WITH  ;
            B.a
       SELECT 2
       SKIP
    ENDDO
    SELECT 1
    REPLACE name WITH '女' FOR code='0'
    REPLACE name WITH '男' FOR code='1'
close all    
    USE cj
    INDEX ON dwdm+CJDM TO cj
    TOTAL ON dwdm+CJDM TO DATA\cjk
    *********父子级筛选：分级分类汇总******20170104****************
    USE bm
    INDEX ON dwdm+cjdm+BMBH TO bm
    TOTAL ON dwdm+cjdm+BMBH TO DATA\bmdm
    *********父子级筛选：分级分类汇总******20170104****************
    SELECT 1
    USE EXCLUSIVE 车间
    ZAP
    SELECT 2
    USE cjk
    GO TOP
    DO WHILE  .NOT. EOF()
       SELECT 1
       APPEND BLANK
       REPLACE 单位代码 WITH B.dwdm , 单位名称 with b.dwmc , 车间代码 WITH B.CJDM , 车间名称 WITH B.CJMC , 人数 WITH  ;
            B.a
       SELECT 2
       SKIP
    ENDDO
close all
    SELECT 1
    USE EXCLUSIVE 部门
    ZAP
    SELECT 2
    USE bmdm
    GO TOP
    DO WHILE  .NOT. EOF()
       SELECT 1
       APPEND BLANK
       REPLACE 单位代码 WITH B.dwdm , 单位名称 with b.dwmc ,车间代码 WITH B.CJDM , 车间名称 WITH B.CJMC , 部门编号 WITH B.BMBH , 部门名称 WITH B.BMMC , 人数 WITH B.a
       SELECT 2
       SKIP
    ENDDO
close all
**********************
use ryzk
count for year(csrq)=0 OR year(cjgzrq)=0 OR  empty(erprybh) to rs
if rs>0
   go top
   brow for year(csrq)=0 OR year(cjgzrq)=0 OR  empty(erprybh) fiel cjmc:r:h='部门',bmmc:r:h='班组',ryxm:r:h='姓名',erprybh:h='ERP编号',csrq:h='出生日期',cjgzrq:h='工作日期' titl '请输入下列'+allt(str(rs))+'人的重要信息 [ 按Esc键退出 ]'
ENDIF
count FOR empty(ygxz1) or empty(zgqk1) to rs
if rs>0
   loca FOR empty(ygxz1) or empty(zgqk1)
   brow FOR empty(ygxz1) or empty(zgqk1) fiel dwmc:r:h='单位',cjmc:r:h='部门',bmmc:r:h='班组',ryxm:r:h='姓名',ygxz1:r:h='用工性质',zgqk1:r:h='在岗情况' titl '请在人员状况维护窗口维护下列'+allt(str(rs))+'人“用工性质或在岗情况” [ 按Esc键退出 ]'
endif
count FOR gwgz=0 AND ygxz1='劳务派遣工' to rs
if rs>0
   loca FOR gwgz=0 AND ygxz1='劳务派遣工'
   brow FOR gwgz=0 AND ygxz1='劳务派遣工' fiel dwmc:r:h='单位',cjmc:r:h='部门',bmmc:r:h='班组',ryxm:r:h='姓名',gz:r:h='工种代码',gz1:r:h='工种',gw:r:h='岗位',gwdc1:r:h='岗位档差',gwgz:r:h='岗位工资' titl '请在人员状况维护窗口维护下列'+allt(str(rs))+'人“工种岗位或岗位工资” [ 按Esc键退出 ]'
ENDIF
count FOR cjgz>0 AND zgqk1='在岗人员' to rs
if rs>0
   loca FOR cjgz>0 AND zgqk1='在岗人员'
   brow FOR cjgz>0 AND zgqk1='在岗人员' fiel dwmc:r:h='单位',cjmc:r:h='部门',bmmc:r:h='班组',ryxm:r:h='姓名',cjgz:h='特殊工资',gz1:r:h='工种',gw:r:h='岗位',zgqk1:r:h='在岗情况' titl '请审核下列'+allt(str(rs))+'人“特殊工资” [ 按Esc键退出 ]'
endif
*sum jfjs to js
*if js>0
*   count for jfjs=0 to rs
*   if rs>0
*      go top
*      brow for jfjs=0 fiel cjmc:r:h='部门',bmmc:r:h='班组',ryxm:r:h='姓名',csrq:r:h='出生日期',cjgzrq:r:h='工作日期'whcd1:r:h='学历',jfjs:h='缴费基数' titl '请输入下列'+allt(str(rs))+'人“缴费基数” [ 按Esc键退出 ]'
*   endif
*endif
*******老职工工龄工资处理*****************
replace YLF with 2002-year(CJGZRQ)+1+20 for CJGZRQ<ctod('2002-7-1')
replace gwjt with (2003-year(CJGZRQ)+1)*15 for CJGZRQ<ctod('2003-7-1') and allt(zgqk1)='在岗人员'
**************计算年龄（周岁）工龄（周年）计算******保留两位小数****************
replace nL with ROUND(year(date())-year(Csrq)+(month(date())-month(Csrq))/12,2) for month(Csrq)-month(date())<=0 and year(csrq)>0
replace nL with ROUND(year(date())-year(Csrq)-(month(Csrq)-month(date()))/12,2) for month(Csrq)-month(date())>0 and year(csrq)>0
*************************************周岁年龄（周岁）计算***********************************************************************************
*replace GL with ND1-year(Cjgzrq) for month(Cjgzrq)-month(date())<=0 and year(cjgzrq)>0
*replace GL with ND1-year(Cjgzrq) for month(Cjgzrq)-month(date())>0 and year(cjgzrq)>0 
***********************************仅仅是针对工龄工资计算时使用虚龄，占着一天算一年的计算方法******20170125************************************************* 
replace gL with ROUND(year(date())-year(Cjgzrq)+(month(date())-month(Cjgzrq))/12,2) for month(Cjgzrq)-month(date())<=0 and year(cjgzrq)>0
replace gL with ROUND(year(date())-year(Cjgzrq)-(month(Cjgzrq)-month(date()))/12,2) for month(Cjgzrq)-month(date())>0 and year(cjgzrq)>0 
*************************************年休假使用实际工龄周年计算***********************************************************************************
replace 周岁年龄 with ALLTRIM(LEFT(STR(nl,5,2),2))+'岁'+ALLTRIM(STR(VAL(right(STR(nl,5,2),2))/100*12))+'个月' all
replace 周岁年龄 with ALLTRIM(STR(VAL(right(STR(nl,5,2),2))/100*12))+'个月' for ALLTRIM(LEFT(STR(nl,5,2),2))='0'
replace 周岁年龄 with ALLTRIM(LEFT(STR(nl,5,2),2))+'岁' for ALLTRIM(STR(VAL(right(STR(nl,5,2),2))/100*12))='0'
*************************************工作年限（周年）计算***********************************************************************************
replace 工作年限 with ALLTRIM(LEFT(STR(gl,5,2),2))+'年'+ALLTRIM(STR(VAL(right(STR(gl,5,2),2))/100*12))+'个月' all
replace 工作年限 with ALLTRIM(STR(VAL(right(STR(gl,5,2),2))/100*12))+'个月' for ALLTRIM(LEFT(STR(gl,5,2),2))='0'
replace 工作年限 with ALLTRIM(LEFT(STR(gl,5,2),2))+'年' for ALLTRIM(STR(VAL(right(STR(gl,5,2),2))/100*12))='0'
replace NLD with '20岁以下' for NL<21
replace NLD with '21—25岁' for NL>=21 and NL<26
replace NLD with '26—30岁' for NL>=26 and NL<31
replace NLD with '31—35岁' for NL>=31 and NL<36
replace NLD with '36—40岁' for NL>=36 and NL<41
replace NLD with '41—45岁' for NL>=41 and NL<46
replace NLD with '46—50岁' for NL>=46 and NL<51
replace NLD with '51—55岁' for NL>=51 and NL<56
replace NLD with '56—60岁' for NL>=56 
replace GLD with '25年以下' for GL<25
replace GLD with '25—29年' for GL>=25 and GL<30
replace GLD with '30—34年' for GL>=30 and GL<35
replace GLD with '35年以上' for GL>=35
COPY TO temp FIELDS dwdm,dwmc,CJDM,CJMC,nld,a
COPY TO temp1 FIELDS dwdm,dwmc,CJDM,CJMC,gld,a
USE temp
INDEX ON dwdm+CJDM+nld TO temp
TOTAL ON dwdm+CJDM+nld TO nld
USE temp1
INDEX ON dwdm+CJDM+gld TO temp1
TOTAL ON dwdm+CJDM+gld TO gld 
USE ryzk
sort on bmbh,zw,rybh to temp 
use temp
GO top
bh=rybh
REPLACE a WITH 1,人数 WITH 1 all
COPY TO ry&ny
*************保存和更新人员信息到本月备份ryzk.dbf中****bug*******
COPY TO 人员状况 
COPY TO data\ryzk
COPY TO data\tmpdbf
close all
USE gzze1 EXCLUSIVE
ZAP
USE lwtemp EXCLUSIVE
ZAP
use ccrq
go 8
壁纸数=fdmax
I=val(图片)
pic=ALLTRIM(图片)
repl 图片 with ALLTRIM(str(I+1))
IF I>=壁纸数
   repl 图片 with '0'
ENDIF
**********用户清新杂志封面封底界面*********
pict_fd='fd'+pic+'.jpg'
_SCREEN.PICTURE = '&pict_fd'
close all
