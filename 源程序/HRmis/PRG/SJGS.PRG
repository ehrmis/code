**************************
*SJGS.prg(数据更新)
**************************
close all
yn=messagebox('需要审核或考评当年员工培训情况吗？！...请插入磁盘！',4+32+0,'提示')
if yn=6
        wait window '正在初始化，请稍候... ...' nowait        
       * sele 2
       * use A:\dmk 
       * sele 1
       * OPEN DATABASE dmk
        *use dmk
       * repl dw with allt(b.dw),mc with allt(b.mc),zg with allt(b.zg),rs with allt(b.rs),zb with allt(b.zb),路径 with allt(b.路径)
       * close all
       !copy a:\dmk.dbf
        use ygjy excl
        zap
        appe from zgjy
       * OPEN DATABASE zgjy
        *use zgjy excl
        *zap
        *OPEN DATABASE A:\zgjy
        *appe from A:\zgjy
        *go top
        !copy a:\zgjy????.dbf
        use zgjy
        go top
        dm=allt(dwdm)
        do zjwh
       * OPEN DATABASE zgjy
        use zjzk
        brow noedit title allt(mc)+'员工年度培训统计'
        do khqk
        close all
        wait window '正在更新当前数据，请稍候... ...' nowait    
        *OPEN DATABASE zgjy
       * use zgjy excl
        *zap
        *appe from ygjy
        *delete for dwdm='&dm'
        *pack
        *OPEN DATABASE A:\zgjy
        *appe from a:\zgjy
        close all      
endif
retu