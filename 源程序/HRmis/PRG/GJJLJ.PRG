*************************
*gjjlj.prg(Build20050809)
*************************
close all
use dmk
maxlx=公积金利息
jzy0=对帐基准月
jzy01='0'+str(val(jzy0)-1,1)
sygjj=dm111+right(nd,2)+jzy01
close all
if month(date())>=6
   use dmk
   jzy=对帐基准月
else  
   jzy=' 1'
endif 
jzy1=str(val(yf),2)
rnd=right(nd,2)
if !file('kg'+dm111+rnd+yf+'.dbf')
   jzy1=str(val(jzy1)-1,2)
endif
set colo to 7+/2
on key label F1 do grcx
@ 8,5 say "校对住房公积金累计额基准月份([银行对帐月]或[1月]): "+nd+"年" font '宋体',16 get jzy font '宋体',16
@ 8,col() say "月" font '宋体',16 
read
if val(jzy)=1
   if val(jzy1)>val(jzy0)
      jzy1=jzy0
   endif
endif      
@ 10,10 say "公积金个人结存累计余额截止月份: "+nd+"年" font '宋体',16 get jzy1 font '宋体',16
@ 10,col() say "月" font '宋体',16 
read
jzy=allt(jzy)
close all
if val(jzy)>1  
@ 12,10 say '上年'+allt(jzy0)+'月至今年'+allt(jzy0)+'月银行利息最大值:' font '宋体',16 get maxlx font '宋体',16
@ 12,col() say "元" font '宋体',16 
read
use dmk
repl 对帐基准月 with jzy,公积金利息 with maxlx  
endif
jzy1=allt(jzy1)
kkk=val(jzy1)
k=val(jzy)
nd=str(year(date()),4)
nd1=str(year(date())-1,4)
kk=iif(k>9,str(k,2),'0'+str(k,1))
gjjNY=dm111+right(ND,2)+kk
gjjNY2=dm111+right(ND1,2)+'12'  
if k=1 and file('kg'+gjjny2+'.dbf')
    close all
    sele 2
    use kg&gjjNY2
    inde on rybh to kg&gjjNY2
    sele 1
    use kg&gjjNY
    inde on rybh to kg&gjjNY
    upda on rybh from B repl bnjn with b.bnjn+hj,gjjz with b.gjjz+hj
endif
*******************滚存12月份累计数[gjjz]*******************
close all
sele 4
use kg&sygjj
inde on rybh to kg&sygjj
***特指基准月上月******
if file('kg'+nd1+kk+'.dbf')
   sele 3
   use kg&ND1.&KK
   inde on rybh to kg&ND1.&KK
endif
sele 2
use ryzk
inde on rybh to ryzk
if file('kg'+nd1+kk+'.dbf') 
   upda on rybh from C repl ydhj1 with c.ydhj
endif
sele 1
use kg&gjjNY
clear
jzy1='A'
@ 8,30 say "请选择人工校对公积金排序方案：" font '宋体',16
@ 10,20 say "A.工资表顺序 B.按人员编号排序 C.按姓名排序" font '宋体',16 get jzy1 font '宋体',16
@ 15,20 say "提示：请认真审核新缴纳职工累计（首月累计＝当月合计）" font '宋体',16
read
jzy2='N'
@ 12,35 say "需要逐月分析核对吗？（Y/N）：" font '宋体',16 get jzy2 font '宋体',16
read
use kg&gjjNY
inde on bmbh+rybh to kg&gjjNY
do case
   case upper(jzy1)='B'
        inde on rybh to kg&gjjNY
   case upper(jzy1)='C'
        inde on ryxm to kg&gjjNY
endcase
if val(jzy)=1
   go top
   brow fiel cjmc:r:h='车间',bmmc:r:h='部门',gjjzh:r:h='个人帐号',ryxm:r:h='姓名',jfjs:r:h='缴纳基数',gjj:r:h='月个人金额',gjj1:r:h='月单位资助',hj:r:h='当月合计',bnjn:r:h='自上年'+allt(jzy0)+'月起至今缴交合计',gjjz:r:h='总累计数'  titl "请认真核对"+nd+"年"+jzy+"月累计数  [按 F1 查个人]"
*******************审核滚存12月份累计数[gjjz]*******************
else
   set relation to rybh into 2
   go top
   brow fiel cjmc:r:h='车间',bmmc:r:h='部门',gjjzh:r:h='个人帐号',ryxm:r:h='姓名',b.ydhj1:h='银行期初余额(校正)',ydhj:h='银行期末余额(校正)',gjjz:h='息前总累计数(校正)' titl "请认真核对"+nd+"年"+jzy+"月银行对帐清单及计算机自动累计个人本息总额[按 F1 查个人]"
   jlh=recn()
   repl b.ydhj with ydhj,bnlx with ydhj-gjjz all  
   **********保存银行余额******结算利息********* 
   count for snjn<>ydhj-b.ydhj1-bnlx or bnlx<0 or bnlx>maxlx to rs
   if rs>0
      go top
      brow for snjn<>ydhj-b.ydhj1-bnlx or bnlx<0 or bnlx>maxlx fiel cjmc:r:h='车间',bmmc:r:h='部门',gjjzh:r:h='个人帐号',ryxm:r:h='姓名',b.ydhj1:h='银行期初余额(校正)',snjn:h='自上年'+allt(jzy0)+'月起至今缴交合计(校正)',bnlx:r:h='本期利息',ydhj:h='银行期末余额(校正)',gjjz:r:h='息前总累计数',gjjz=gjjz+bnlx:r:h='息后总累计数' titl "发现下列"+allt(str(rs))+"人有问题，请认真检查核对! [按 F1 查个人]"
      jlh=recn()
      repl b.ydhj with ydhj,gjjz with b.ydhj1+snjn,bnlx with ydhj-gjjz all 
   endif
   go jlh
   bh=rybh
   brow fiel cjmc:r:h='车间',bmmc:r:h='部门',gjjzh:r:h='个人帐号',ryxm:r:h='姓名',b.ydhj1:h='银行期初余额(校正)',snjn:h='自上年'+allt(jzy0)+'月起至今缴交合计(校正)',bnlx:r:h='本期利息',ydhj:h='银行期末余额(校正)',gjjz:r:h='息前总累计数',gjjz=gjjz+bnlx:r:h='息后总累计数' titl "请认真核对结算利息[按 F1 查个人]"
repl bnjn with 0,b.ydhj with ydhj,gjjz with b.ydhj1+snjn,bnlx with ydhj-gjjz all
********起始月缴纳清零并保存审核校对数据********
   if file('kg'+nd1+kk+'.dbf')
      sele 3
      upda on rybh from B repl ydhj with b.ydhj1
   endif  
sele 4
upda on rybh from A repl snjn with a.snjn            
*******************个人结存利息结算[每年4月]及保存银行提供数据*******************
endif
sele 1
copy to kg&ND.&KK
*******************保存基准月个人结存数据****开始逐月分析***************
for k1=k+1 to kkk
    kk1=iif(k1>9,str(k1,2),'0'+str(k1,1))
    kk=iif(k>9,str(k,2),'0'+str(k,1))
    gjjNY=dm111+right(ND,2)+kk 
    gjjNY1=dm111+right(ND,2)+kk1
    close all
    sele 2
    use kg&gjjNY
    inde on rybh to kg&gjjNY
    sele 1
    use kg&gjjNY1
    set relation to rybh into 2
    if val(jzy)=1
       repl snjn with b.snjn+hj all 
    else
       repl snjn with b.snjn all 
    endif
    inde on rybh to kg&gjjNY1
    upda on rybh from B repl bnjn with b.bnjn+hj , bnlx with b.bnlx , gjjz with b.gjjz+hj
    k=k1
    inde on bmbh+rybh to kg&gjjNY1
if upper(jzy2)='Y'
do case
   case upper(jzy1)='B'
        inde on rybh to kg&gjjNY1
   case upper(jzy1)='C'
        inde on ryxm to kg&gjjNY1
endcase
  loca for rybh=bh
  if eof()
     go top
  endif   
  if val(jzy)=1
     brow fiel cjmc:r:h='车间',bmmc:r:h='部门',ryxm:r:h='姓名',jfjs:r:h='缴纳基数',gjj:r:h='月个人金额',gjj1:r:h='月单位资助',hj:r:h='本月缴交合计',snjn:r:h='自上年'+allt(jzy0)+'月起缴交合计',bnlx:r:h='本期利息',gjjz=gjjz+bnlx:r:h='总累计数' titl "请认真核对"+nd+"年"+allt(str(k1,2))+"月累计数[按 F1 查个人]"
  else
     brow fiel cjmc:r:h='车间',bmmc:r:h='部门',ryxm:r:h='姓名',jfjs:r:h='缴纳基数',gjj:r:h='月个人金额',gjj1:r:h='月单位资助',hj:r:h='本月缴交合计',bnjn=snjn+bnjn:r:h='自上年'+allt(jzy0)+'月起缴交合计',bnlx:r:h='本期利息',gjjz=gjjz+bnlx:r:h='总累计数' titl "请认真核对"+nd+"年"+allt(str(k1,2))+"月累计数[按 F1 查个人]"
  endif
  bh=rybh
endif
endfor
bcyf=iif(k1-1>9,str(k1-1,2),'0'+str(k1-1,1))
bcny=nd+bcyf
if file('kg'+bcny+'.dbf')
yn = MESSAGEBOX("需要恢复"+allt(str(k1-1,2))+"月份审核校对过的个人结存数据吗？",4+32,"提示")
if yn=6
   use kg&bcny
   copy to kg&gjjNY1
   use kg&gjjNY1
endif 
endif 
loca for rybh=bh
  if eof()
     go top
  endif   
if val(jzy)=1
   brow part 15 fiel cjmc:r:h='车间',bmmc:r:h='部门',gjjzh:r:h='个人帐号',ryxm:r:h='姓名',jfjs:r:h='缴纳基数',hj:r:h='本月缴交合计',snjn:r:h='自上年'+allt(jzy0)+'月起缴交合计',bnlx:r:h='本年利息',gjjz:r:h='息前总累计数',gjjz=gjjz+bnlx:r:h='息后总累计数' titl "请认真核对"+nd+"年"+allt(str(k1-1,2))+"月累计数[按 F1 查个人]"
else
   sele 3
   use ryzk
   inde on rybh to ryzk
   sele 1
   set relation to rybh into 3
   brow part 15 fiel cjmc:r:h='车间',bmmc:r:h='部门',gjjzh:r:h='个人帐号',ryxm:r:h='姓名',jfjs:r:h='缴纳基数',hj:r:h='本月缴交合计',c.ydhj1:r:h='银行期初余额',snjn:h='上年'+allt(jzy0)+'月至今年'+allt(jzy0)+'月缴纳(校正)',bnjn:h='自'+allt(jzy0)+'月起缴纳(校正)',bnjn=snjn+bnjn:r:h='自上年'+allt(jzy0)+'月起至今缴交合计',bnlx:h='本年利息(校正)',gjjz:h='息前总累计数(校正)',gjjz=gjjz+bnlx:r:h='息后总累计数' titl "请认真审核校对"+nd+"年"+allt(str(k1-1,2))+"月统计情况[按 F1 查个人]"
   yn = MESSAGEBOX("需要保存"+allt(str(k1-1,2))+"月份审核校对过的个人结存数据吗？(修改过才保存)",4+32,"提示")
   if yn=6
      copy to kg&bcny type fox2x
   endif   
endif
close all
clear
retu


                                 