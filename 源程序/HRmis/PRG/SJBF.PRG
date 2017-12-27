*************
*backup.prg
*************
close all
ok = MESSAGEBOX("插入磁盘备份当前数据！？",1+32,"提示")
if ok=1
wait window '正在备份当前员工培训情况... ...' nowait
*!copy zgjy.dbc a:\
!copy zgjy????.dbf a:\
*!copy zgjy.dct a:\
*!copy dmk.dbc a:\
!copy dmk.dbf a:\
*!copy dmk.dct a:\
*if file('a:\zgjy.dbf')
 *  OPEN DATABASE a:\zgjy 
  * use a:\zgjy excl
  * zap
   *OPEN DATABASE zgjy 
   *appe from zgjy
*else
 *  OPEN DATABASE zgjy
  * use zgjy
   *copy to A:\zgjy
*endif
*if file('a:\dmk.dbf')
 *  OPEN DATABASE a:\dmk
  * use a:\dmk excl
   *zap
   *OPEN DATABASE dmk
   *appe from dmk
*else
 *  OPEN DATABASE dmk
  * use dmk
  * copy to A:\dmk
*endif
endif
*close all
retu