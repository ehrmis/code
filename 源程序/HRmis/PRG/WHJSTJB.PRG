****************
* whjstjb.prg
****************
close all
set talk off
set safety off
wait wind "正在保存数据... ..." nowa
use zgjy
repl 人数 with 1 all
inde on whcd to zgjy
count for '党员'$zzmm1 to dyrs
count for xb='0' to nu
sum 人数 to zrs
tota on whcd to WHCD field 人数
tota on whcd to czry_wh field 人数 for !(val(ryfl)>=3 and val(ryfl)<=7)
tota on whcd to glry_wh field 人数 for val(ryfl)>=3 and val(ryfl)<=7
inde on zcdj to zgjy
tota on zcdj to JSZC field 人数
copy to czry_js  for !(val(ryfl)>=3 and val(ryfl)<=7)
copy to glry_js  for val(ryfl)>=3 and val(ryfl)<=7
close all
sele 1
use WHCD
sele 2
use JSZC
sele 3
use czry_wh
sele 4
use czry_js
count for '党员'$zzmm1 to dyrs1
count for xb='0' to nu1
sum 人数 to zrs1
sele 5
use glry_wh
sele 6
use glry_js
count for '党员'$zzmm1 to dyrs2
count for xb='0' to nu2
sum 人数 to zrs2
sele 7
use dmk
sele 8
use 文化技术 excl
zap
appe blank
repl 项目 with '操作人员',职工人数 with zrs1,女职工人数 with nu1,党员人数 with dyrs1
sele 4
sum 人数 to xt for ryfl='02' and ryfl='01' 
sum 人数 to cj for gl<8 and ryfl='01' 
sum 人数 to zj for gl>=8 and gl<15 and ryfl='01'
sum 人数 to gj for gl>=15 and ryfl='01'
sum 人数 to cj1 for zcdj='08'
sum 人数 to zj1 for zcdj='07'
sum 人数 to gj1 for zcdj='06'
sum 人数 to js for zcdj='09'
sele 8
repl 学徒 with xt,初级在岗数 with cj,初级取证数 with cj1,中级在岗数 with zj,中级取证数 with zj1,高级在岗数 with gj,高级取证数 with gj1,技师 with js,其它操作人员 with zrs1-xt-cj-zj-gj-js
sele 3
loca for whcd='02'
if !eof()
   sele 8
   repl 研究生 with c.人数
endif
sele 3
loca for whcd='03'
if !eof()
   sele 8
   repl 大学本科 with c.人数
endif
sele 3
loca for whcd='04'
if !eof()
   sele 8
   repl 专科 with c.人数
endif
sele 3
loca for whcd='05'
if !eof()
   sele 8
   repl 中专 with c.人数
endif
sele 3
loca for whcd='06'
if !eof()
   sele 8
   repl 技校 with c.人数
endif
sele 3
loca for whcd='07'
if !eof()
   sele 8
   repl 高中 with c.人数
endif
sele 3
loca for whcd='08'
if !eof()
   sele 8
   repl 初中 with c.人数
endif
sele 3
loca for whcd='09'
if !eof()
   sele 8
   repl 小学 with c.人数
endif
sele 3
loca for whcd='10'
if !eof()
   sele 8
   repl 文盲 with c.人数
endif
sele 8
appe blank
repl 项目 with '管理人员',职工人数 with zrs2,女职工人数 with nu2,党员人数 with dyrs2
sele 5
loca for whcd='02'
if !eof()
   sele 8
   repl 研究生 with e.人数
endif
sele 5
loca for whcd='03'
if !eof()
   sele 8
   repl 大学本科 with e.人数
endif
sele 5
loca for whcd='04'
if !eof()
   sele 8
   repl 专科 with e.人数
endif
sele 5
loca for whcd='05'
if !eof()
   sele 8
   repl 中专 with e.人数
endif
sele 5
loca for whcd='06'
if !eof()
   sele 8
   repl 技校 with e.人数
endif
sele 5
loca for whcd='07'
if !eof()
   sele 8
   repl 高中 with e.人数
endif
sele 5
loca for whcd='08'
if !eof()
   sele 8
   repl 初中 with e.人数
endif
sele 5
loca for whcd='09'
if !eof()
   sele 8
   repl 小学 with e.人数
endif
sele 5
loca for whcd='10'
if !eof()
   sele 8
   repl 文盲 with e.人数
endif
sele 6
sum 人数 to cjzc for zcdj='04' or zcdj='05'
sum 人数 to zjzc for zcdj='03'
sum 人数 to fgjzc for zcdj='02'
sum 人数 to gjzc for zcdj='01'
sele 8
repl 初级职称 with cjzc,中级职称 with zjzc,副高级 with fgjzc,正高级 with gjzc,其它管理人员 with zrs2-cjzc-zjzc-fgjzc-gjzc
appe blank
repl 项目 with '总    计',职工人数 with zrs,女职工人数 with nu,党员人数 with dyrs
sele 1
loca for whcd='02'
if !eof()
   sele 8
   repl 研究生 with a.人数
endif
sele 1
loca for whcd='03'
if !eof()
   sele 8
   repl 大学本科 with a.人数
endif
sele 1
loca for whcd='04'
if !eof()
   sele 8
   repl 专科 with a.人数
endif
sele 1
loca for whcd='05'
if !eof()
   sele 8
   repl 中专 with a.人数
endif
sele 1
loca for whcd='06'
if !eof()
   sele 8
   repl 技校 with a.人数
endif
sele 1
loca for whcd='07'
if !eof()
   sele 8
   repl 高中 with a.人数
endif
sele 1
loca for whcd='08'
if !eof()
   sele 8
   repl 初中 with a.人数
endif
sele 1
loca for whcd='09'
if !eof()
   sele 8
   repl 小学 with a.人数
endif
sele 1
loca for whcd='10'
if !eof()
   sele 8
   repl 文盲 with a.人数
endif
sele 8
repl 学徒 with xt,初级在岗数 with cj,初级取证数 with cj1,中级在岗数 with zj,中级取证数 with zj1,高级在岗数 with gj,高级取证数 with gj1,技师 with js,其它操作人员 with zrs1-xt-cj-zj-gj-js,初级职称 with cjzc,中级职称 with zjzc,副高级 with fgjzc,正高级 with gjzc,其它管理人员 with zrs2-cjzc-zjzc-fgjzc-gjzc
repl mc with g.mc,zg with g.zg,rs with g.rs,zb with g.zb all
close all
use whcd
copy to 文化程度 field whcd1,人数
use jszc
copy to 技术职称 field zc,人数
close all
delete file whcd.dbf
delete file jszc.dbf
retu
                               