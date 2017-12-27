*******************************
* 养老保险缴费基数.prg(20151124)
*******************************
close all
set safety off
clear
_SCREEN.PICTURE = ''
pd='N'
temp='jfjstemp'
set talk off
@ 1,22 say '自动填写系统外缴费工资申报表快速解决方案' font '黑体',20
@ 4,3 say '第一步：请先顺利完成'+ly+'年工资总额' font '宋体',13
@ 6,3 say '第二步：插入社保部拷来的“缴费基数申报”数据盘，建议备份好“基数申报”电子表！'font '宋体',13
@ 8,3 say '第三步：系统自动更新社保部拷来数据盘中的应填写数据' font '宋体',13
@ 10,3 say '第四步：退出系统后打开上报社保部数据盘文件（“基数申报”电子表[.XLS]）审核维护“备注”等' font '宋体',12
@ 12,3 say '更新成功后拔出上报盘.....嘿嘿！报盘就是这么简单！！' font '宋体',13
yn = MESSAGEBOX("自动填写系统外缴费工资申报表吗？",4+32,"提示")
if yn=7
close all  
use zr1&ly
COPY TO &bf.\基数申报 FIELDS erprybh,scbh,ryxm,xb1,缴费工资,jfjs TYPE XL5 
********************人工维护好社保部拷来的数据盘************************************************
ELSE
close all
=MESSAGEBOX("请选择社保部拷来的“基数申报”电子表",48,"自动填写申报数据")
old_dirs='&xp'+'\基数申报.xls'
dirs=GETFILE("xls","基数申报电子表：")
if empty(dirs)
   dirs=old_dirs
endif
bakfile=allt(dirs)
if upper(right(bakfile,3))='XLS'
   use 基数申报 excl
   zap 
   APPEND FROM &bakfile TYPE xl8 
   brow titl '导入社保部拷来的数据盘中养老保险原始数据！'  
else   
   =MESSAGEBOX("数据导入失败！",48,"提示")
   close all
   clear 
   retu
endif        
*****************正常维护：默认使用人员状况库中的“个人编号”***************
close all
sele 2
use ryzk
inde on rybh to ryzk
sele 1
use zr1&ly
repl scbh with space(10) all
inde on rybh to zr1&ly
upda on rybh from B repl scbh with b.scbh
go top
fd = MESSAGEBOX("使用上年社会平均工资封顶保底处理吗？",4+32,"提示")
******************自动更新基数**************
close all
sele 2
use 基数申报 excl
*******“缴费基数申报”电子表***A序号B个人编号C姓名D公民身份证号码E月工资F备注*********20151123***************               
repl e with '' all
*****上报基数电子表清空重新生成**********
i=1
 scan 
      sele 1
      use zr1&ly
      loca for val(scbh)=val(b.b) 
       wait window '正在用个人编号关联填写第'+allt(str(i))+'/'+allt(str(recc()))+'人〔'+ALLTRIM(Str(i/recc()*100))+'%〕缴费基数,请稍候......' nowait
      if fd=6
      repl b.e with ALLTRIM(str(jfjs))         
         *********上报基数*******
      else
      repl b.e with ALLTRIM(str(pj))        
        *********申报基数*******    
      endif
       i=i+1 
  endscan
close all
use 基数申报 excl
count for empty(e) and len(allt(b))=10 to rs
if rs>0
  delete FOR empty(e) and len(allt(b))=10
   go top
   brow FOR empty(e) and len(allt(b))=10 fiel b:h='个人编号',c:h='姓名',d:h='身份证号',e:h='月缴费工资' titl '请维护下列人员'+nd1+'年缴费基数,多余人员将被删除！'
   yn3 = MESSAGEBOX("删除无编号记录吗？",4+32,"提示")
   if yn3=6
      pack
   else
      recall all
   endif   
endi
go top 
brow fiel  a:h='序号',b:h='个人编号',c:h='姓名',d:h='身份证号',e:h='月缴费工资',f:h='备注' titl '请认真审核“养老保险缴费基数”更新情况(特别增减人员)，将生成电子表：'+bf+'\基数申报 ·XLS'
*********************审核基数*****************
COPY TO &bf.\基数申报 TYPE XL5 
********************人工维护好社保部拷来的数据盘************************************************
close all  
use zr1&ly
************************通用功能：重号检查***********
repl a with 1 all
sort on scbh to temp
use temp excl
go top
do while !eof()
   bh=scbh
   skip
   if bh=scbh
      repl a with 0
      skip-1
      repl a with 0
      skip
   endi
enddo
coun for a=0 to aaa
if aaa>0
   copy to jfjstemp for a=0
   temp='jfjstemp'
   pd='Y'
   use jfjstemp excl
   inde on bmbh+rybh to jfjstemp
   go top
   brow fiel scbh:h='养老保险个人编号',cjmc:h='车间',bmmc:h='部门',ryxm:h='姓名',sfz:h='身份证号' titl '发现'+allt(str(aaa,4))+'人个人编号重号或空号！请修改正确（参照'+bf+"\编号重号·XLS）"   
   COPY TO &bf.\编号重号  TYPE XL5 FIELDS cjmc,bmmc,ryxm,rybh,sfz,scbh
   ********************人工维护好人员状况再操作一次（不能再导入社保部拷来的数据盘中的“个人编号”）************************************************
endif
close all
ENDIF
clear
yn2 = MESSAGEBOX("缴费工资填写或导出成功，现在使用电子表编辑打印吗？",4+32,"提示")
IF yn2 = 6
  if yn=6
     myexcel=CREATEOBJECT("excel.application")
     myexcel.workbooks.open("&bakfile")
     *********随机文件名***通用表达式*****
     myexcel.visible=.t.
     myexcel.caption="使用电子表编辑打印报表" 
  endif
myexcel=CREATEOBJECT("excel.application")
myexcel.workbooks.open("&bf"+"\基数申报.xls")
myexcel.visible=.t.
myexcel.caption="使用电子表编辑打印报表" 
ENDI   
close all
clear
use ccrq
go 8
I=val(图片)
repl 图片 with str(I+1,2)
IF I>=54
   repl 图片 with '0'
ENDIF
pic=ALLTRIM(图片)
**********用户清新杂志封面封底界面*********
pict_fd='fd'+pic+'.jpg'
_SCREEN.PICTURE = '&pict_fd'


