do while .t.
      do srxm.spr
      sele 2
      go top
      locate for allt(RYXM)=allt(XM) or allt(RYBH)=allt(XM)
      JLH=recn()
      if eof()
        do dhkjg.spx
    M = readkey()
    if M=12 or M=268
      return
    endif
      else
      wait wind '已找到了！请先按[PgDn]、[PgUp]两个键刷新一下' nowa 
      exit    
     endif
enddo



proc aaaaa

select 1
    go top
    KKK1 = 5
    LLL1 = 0
do while  not eof()
      RYHDS = allt(FIELD_NAME)
      TM123 = allt(HHH1)
      select 2
      go JLH
      if upper(alltrim(RYHDS))='RYBH' or upper(alltrim(RYHDS))='RYXM' or upper(alltrim(RYHDS))='BMMC';
 or upper(alltrim(RYHDS))='CJMC'
        RYHDS=&RYHDS
        @ KKK1 , 40+LLL1 say RYHDS
      else
        @kkk1,40+lll1 get &ryhds PICT &TM123
      endif
      KKK1 = KKK1+2
      if KKK1>20
        KKK1 = 5
        LLL1 = LLL1+25
      endif
      select 1
      skip
enddo   
select 2
