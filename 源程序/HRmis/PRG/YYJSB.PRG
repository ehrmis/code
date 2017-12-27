****************************
* .\YYJSB.PRG(Build20050727)
****************************
define window MY from 5 , 25 to 20 , 75 color scheme 5 shadow title 'ÏÖĞĞÍËĞİ´ıÓöÉóÅú°ì·¨'
activate window MY
@ 1 , 5 say '1.±¨ÅúÊ±¼ä:Ò»°ãÒÔ·¨¶¨ÍËĞİµ±ÔÂÎª×¼;'
@ 2 , 5 say '2.°´·¨¶¨ÍËĞİÊ±¼äÕı³£°ìÀíÕß:ÍËĞİÔÂ·İÓë³öÉúÔÂ'
@ 3 , 0 say '·İÏàÍ¬,´ÎÔÂÆğÍ£½ÉÑøÀÏ±£ÏÕ¡¢Ê§Òµ±£ÏÕµÈÉç»á±£ÏÕ(ÖÕ'
@ 4 , 0 say 'Ö¹Éç»á±£ÏÕ½É·Ñ¹ØÏµ);'
@ 5 , 5 say '3.Òò¹ÊÍÏÑÓÕı³£°ìÀíÊ±¼äÕß: 20ÈÕÇ°±¨Éç±£²¿ÕßÍËĞİ'
@ 6 , 2 say 'ÔÂ·İÒÔ±¨ÅúÔÂ·İÎª×¼,´ÎÔÂÆğÍ£½ÉÉç»á±£ÏÕ,·´Ö®Ë³ÑÓ'
@ 11 , 10 say '×ÉÑ¯µç»°:8603146  8636663'
@ 12 , 26 say 'À¥¸ÖÉç±£²¿'
@ 13 , 20 say '¶ş£°£°ÈıÄêËÄÔÂ¶şÊ®ÎåÈÕ'
= inkey(60,'M')
deactivate window MY
close all
set talk off
@ 3 , 28 to 24 , 78
set color to gr+/b , bg+/rb
@ 4 , 29 clear to 23 , 77
PD = ''
YX = '1'
XM = ''
RQ = ''
DZZ = 'N'
JS = 'N'
JSZ = 'N'
Á¬Ğø¹¤Áä = space(6)
¹¤ÖÖÄêÏŞ = space(5)
¹¤ÁäÏµÊı = space(4)
store space(2) to ¼Æ·¢±ÈÀı , ¼Æ·¢±ÈÀı1 , ¼Æ·¢±ÈÀı2 , ¼Æ·¢±ÈÀı3 , ¼Æ·¢±ÈÀı4;
 , STJFN , STJFN1
store 1 to Á¬Ğø¹¤Áä , ½É·ÑÄêÏŞ , È«²¿½É·ÑÄê , JLH
store 0 to ¹¤ÁäÏµÊı , ¹¤ÖÖÄêÏŞ , ¸öÈËÖ¸Êı , ¸Ú¼¼ , ¾ÅÆß , Æ½¾ùÉçÆ½ , ZSHGZ;
 , DZF , JT , BT , JFN
use sbdmk
¾«¶ÈÖµ = Ö¸Êı»¯¾«¶È
DM111 = alltrim(DW)
MC111 = alltrim(MC)
snspn=ÉÏÄêÉçÆ½
use ryzk
do while .t.
  do srxm.spr
  JBXM = left(XM,8)
  locate for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)
  GZRQ = alltrim(CJGZRQ)
  if val(GZRQ)<19951001
     ½É·ÑÆğÊ¼Äê='1995'
  else
     ½É·ÑÆğÊ¼Äê=left(GZRQ,4)
  endif
  if eof()
    do jg01.spx
    M = readkey()
    if M=12 or M=268
      return
    endif
  else
    if file('zr1'+LY1+'.dbf')
    use zr1&LY1
    go top
    locate for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)
    if PJ=0
      @ 12 , 33 say 'ÇëÏÈÎ¬»¤'+LY1+'Äê¹¤×Ê×Ü¶î!... ...'
      = inkey(16,'M')
      return
    endif
    else
      @ 12 , 33 say 'ÇëÏÈ¼ÆËã'+LY1+'Äê¹¤×Ê×Ü¶î!... ...'
      = inkey(16,'M')
      return
    endif
    use yl&LY1
    go top
    locate for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)
    if ÀúÄêµ¥Î»½É+ÀúÄê¸öÈË½É=0
      @ 12 , 33 say 'ÇëÏÈÉóºË'+LY1+'Äê¸öÈËÕÊ»§,È·±£ÀúÄêÊı¾İ²»¿Õ!... ...'
      = inkey(16,'M')
      return
    endif
    do srtxrq.spx
    use yl&LY
    do while val(GZRQ)<19951001
      locate for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)
      @ 22 , 43 say 'ÌáÊ¾:²Î¼û¸öÈËÕÊ»§µÇ¼Ç¿¨'
      @ 11 , 30 say 'ÇëÊäÈë'+alltrim(RYXM)+'ÊÓÍ¬½É·ÑÄê:' get STJFN picture;
 '99'
      read
      @ 11 , 56 say 'Äê' get STJFN1 picture '99'
      read
      JFN = round(val(STJFN)+val(STJFN1)/12,2)
      @ 11 , 61 say '¸öÔÂ='
      PD = 'Y'
      @ 11 , 66 color w+/b say str(JFN,5,2)+'Äê'
      @ 13 , 46 color r+/b say 'ÕıÈ·Âğ?(Y/N)' get PD
      read
      replace ÊÓÍ¬½É·ÑÄê with JFN for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)
      if JFN>0 and upper(PD)='Y'
        exit
      endif
    enddo
    exit
  endif
enddo
ÈË¿É = 'ÍËĞİ'
do YLBXJS
NY = left(RQ,4)+substr(RQ,6,2)
close all
use lnk&LY
replace ½É·ÑÔÂÊı with 6 for Äê·İ='1995'
replace ÔÂ½É·Ñ»ùÊı with 0 for Äê·İ='2001'
replace ½É·ÑÔÂÊı with 12 for Äê·İ<>'1995'
replace ËµÃ÷ with str(val(Äê·İ)-1,4)+'ÄêÔÂÆ½¾ù¹¤×Ê' all
define window TS_WIN from 4 , 23 to 18 , 75 color scheme 10 shadow title;
 'ÑøÀÏ±£ÏÕÊÖ²á'
@ 20 , 50 say 'ÊäÈëÌáÊ¾'
@ 21 , 30 colo 4+/1 say '(1).2001ÄêÔÂ½É·Ñ»ùÊıÖ»ĞèÊäÈë2000ÄêÔÂÆ½¾ù¹¤×Ê'
@ 22 , 30 say '(2).ÒªÇó´Ó±¾ÈË½É·ÑÆğÊ¼ÄêÊäÈë'
@ 23 , 45 say '°´ [ Ctrl+End ] ´æÅÌ'
activate window TS_WIN
go top
browse for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5) field;
 RYBH :R : 5 :H = '±à ºÅ' , RYXM :R : 8 :H = 'ĞÕ  Ãû' , Äê·İ :R ;
,½É·ÑÔÂÊı , ÔÂ½É·Ñ»ùÊı , ËµÃ÷:R
deactivate window TS_WIN
go top
loca for Äê·İ='2002' and (left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5))
if !eof()
   Æ½¾ù2=ÔÂ½É·Ñ»ùÊı
endif
go top
loca for Äê·İ='2001' and (left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5))
if !eof()
   Æ½¾ù1=ÔÂ½É·Ñ»ùÊı
endif
replace ÔÂ½É·Ñ»ùÊı with round((Æ½¾ù1*(½É·ÑÔÂÊı-3)+Æ½¾ù2*3)/½É·ÑÔÂÊı,1) for Äê·İ='2001' and ½É·ÑÔÂÊı>3 and (left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5))
replace ÔÂ½É·Ñ»ùÊı with Æ½¾ù2 for Äê·İ='2001' and ½É·ÑÔÂÊı<=3 and (left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5))
*************ÔÆÕş·¢[2004]45ºÅ*****
sum ÔÂ½É·Ñ»ùÊı to yjshj for (left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5))
count for ÔÂ½É·Ñ»ùÊı>0 and (left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)) to yjscc
Êµ½ÉÆ½¾ù=round(yjshj/yjscc,2)
**********************************
close all
select 2
use »ù´¡Êı¾İ
delete for val(Äê·İ)>year(date())
pack
if snspn=1996
   sum to Æ½¾ùÉçÆ½ ÉÏÄêÉçÆ½£¤ for val(Äê·İ)>val(½É·ÑÆğÊ¼Äê)
   count to ÄêÊı for val(Äê·İ)>val(½É·ÑÆğÊ¼Äê)
else
   sum to Æ½¾ùÉçÆ½ ÉÏÄêÉçÆ½£¤ for val(Äê·İ)>=val(½É·ÑÆğÊ¼Äê)
   count to ÄêÊı for val(Äê·İ)>=val(½É·ÑÆğÊ¼Äê)
endif
½É·ÑÄêÊı = val(LY)-val(½É·ÑÆğÊ¼Äê)+1
Æ½¾ùÉçÆ½ = round(Æ½¾ùÉçÆ½/ÄêÊı,1)
index on Äê·İ to jcsj
*************ÔÆÕş·¢[2004]45ºÅ*****
go bott
ÉÏÄêÉçÆ½=ÉÏÄêÉçÆ½£¤
**********************************
go top
select 1
use lnk&ly
go bottom
set filter to left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)
go top
inde on Äê·İ to lnk&ly
update on Äê·İ from B replace Ö¸Êı with round(ÔÂ½É·Ñ»ùÊı/b.ÉÏÄêÉçÆ½£¤,4)
sum for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5) to ¸öÈËÖ¸Êı;
 Ö¸Êı
@ 3 , 28 to 24 , 78
set color to gr+/b , bg+/rb
@ 4 , 29 clear to 23 , 77
@ 20 , 50 say '¼ÆËãÌáÊ¾'
@ 21 , 30 colo 4+/1 say '(1).Çë³­ÓÚ[À¥¸ÖÖ°¹¤ÍËĞİ´ıÓö¼ÆËã±í]ÖĞÏàÓ¦¿Õ¸ñ'
@ 22 , 30 say '(2).ÒªÇó´Ó±¾ÈË½É·ÑÆğÊ¼ÄêÌîĞ´'
@ 23 , 45 say '°´ [ Ctrl+End ] ´æÅÌ'
activate window TS_WIN
brow fiel ryxm:h='ĞÕÃû',ÔÂ½É·Ñ»ùÊı,Ö¸Êı noedit titl 'Ö¸ÊıºÏ¼Æ='+str(¸öÈËÖ¸Êı,7,4)+';'+'Æ½¾ùÔÂ½É·Ñ»ùÊı='+str(Êµ½ÉÆ½¾ù,7,4)
deactivate window TS_WIN
set filter to
close all
select 3
use sbdmk
JBT = JTF
SXF = SBF+XLF
select 2
use ryzk
go top
locate for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)
¸Ú¼¼ = ·â¶¨¹¤×Ê
select 1
use yl&ly
go top
locate for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)
set color to w+/g
if left(alltrim(MC111),4)='À¥¸Ö'
  MC111 = alltrim('À¥Ã÷¸ÖÌú¼¯ÍÅÓĞÏŞÔğÈÎ¹«Ë¾'+substr(alltrim(MC111),5,16))
else
  MC111 = alltrim('À¥Ã÷¸ÖÌú¼¯ÍÅÓĞÏŞÔğÈÎ¹«Ë¾'+alltrim(MC111))
endif
TT1 = '  µ¥Î»Ãû³Æ:&MC111'+'                    ½ğ¶îµ¥Î»:Ôª(±£Áô½Ç)'
TT2 = '©°©¤©¤©Ğ©¤©¤©¤©¤©Ğ©¤©¤©Ğ©¤©Ğ©¤©¤©Ğ©¤©¤©¤©¤©¤©Ğ©¤©¤©¤©Ğ©¤©¤©¤©¤©¤©Ğ©¤©¤©¤©Ğ©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©´'
TT4 = '©¦    ©¦        ©¦    ©¦  ©¦ÄêÔÂ©¦          ©¦×÷ÄêÔÂ©¦          ©¦(¹¤ÖÖ)©¦                    ©¦'
TT5 = '©À©¤©¤©Ø©Ğ©¤©¤©¤©Ø©¤©¤©à©¤©Ø©¤©Ğ©Ø©¤©¤©¤©¤©Ğ©Ø©¤©¤©¤©Ø©¤©Ğ©¤©¤©¤©Ø©¤©Ğ©¤©Ø©¤©¤©¤©¤©Ğ©¤©¤©¤©¤©¤©È'
TT6 = '©¦ÊÓÍ¬½É©¦            ©¦È«²¿½É©¦          ©¦ÀÛ¼Æ¸öÈËÕÊ»§©¦          ©¦ÀÛ¼Æ¸öÈË½É·Ñ©¦          ©¦'
TT7 = '©À©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©Ø©¤©¤©¤©¤©¤©Ø©¤©Ğ©¤©¤©¤©¤©Ø©¤©Ğ©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©È'
TT8 = '©À©¤©¤©Ğ©¤©Ğ©¤©¤©¤©¤©¤©à©¤©¤©¤©Ğ©¤©¤©Ğ©¤©¤©¤©¤©à©¤©¤©Ğ©¤©Ğ©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ğ©¤©¤©¤©¤©¤©¤©È'
TT9 = '©¦    ©¦Ğò©¦  Ïî  Ä¿  ©¦ »ùÊı ©¦±ÈÀı©¦ ½ğ  ¶î ©¦    ©¦Ğò©¦                      ©¦   ½ğ  ¶î   ©¦'
TT10 = '©¦ ·â ©¦ºÅ©¦          ©¦      ©¦ %  ©¦        ©¦    ©¦ºÅ©¦    Ïî          Ä¿    ©¦ ( ±È  Àı ) ©¦'
TT11 = '©¦    ©À©¤©à©¤©¤©¤©¤©¤©à©¤©¤©¤©à©¤©¤©à©¤©¤©¤©¤©È    ©À©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©È'
TT12 = '©¦    ©¦  ©¦  104ºÅ   ©¦      ©¦    ©¦        ©¦    ©¦  ©¦                      ©¦            ©¦'
TT13 = '©¦ Ê± ©À©¤©à©¤©¤©¤©¤©¤©à©¤©¤©¤©à©¤©¤©à©¤©¤©¤©¤©È °ì ©À©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©È'
TT14 = '©¦ Ô­ ©¦  ©¦ [86] 7ºÅ ©¦      ©¦    ©¦        ©¦ ·¨ ©¦  ©¦                      ©¦            ©¦'
TT15 = '©¦    ©À©¤©à©¤©¤©¤©¤©¤©à©¤©¤©¤©Ø©¤©¤©à©¤©¤©¤©¤©È    ©À©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©È'
TT16 = '©¦ °ì ©¦3 ©¦  ÔÆÀÍÏÕ  ©¦   Êé±¨·Ñ¡¢ ©¦        ©¦ ¼Æ ©¦12©¦ÔÂ¹ı¶ÉĞÔÑøÀÏ½ğ        ©¦            ©¦'
TT17 = '©¦ ·¨ ©À©¤©à©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©È Ëã ©À©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©È'
TT18 = '©¦    ©¦4 ©¦  ÔÆÕş·¢  ©¦            ©¦        ©¦    ©¦13©¦ÔÂ¹ı¶ÉĞÔµ÷½Ú½ğ        ©¦            ©¦'
TT19 = '©¦    ©À©¤©à©¤©¤©¤©¤©¤©à©¤©¤©¤©Ğ©¤©¤©à©¤©¤©¤©¤©È    ©À©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©È'
TT20 = '©¦    ©¦  ©¦   ÌõÁî   ©¦      ©¦    ©¦        ©¦    ©¦  ©¦                      ©¦            ©¦'
TT21 = '©¦ ÔÂ ©À©¤©à©¤©¤©¤©¤©¤©à©¤©¤©¤©à©¤©¤©à©¤©¤©¤©¤©È ±¾ ©À©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©È'
TT22 = '©¦ »ù ©¦  ©¦ [88]14ºÅ ©¦      ©¦    ©¦        ©¦ Ñø ©¦  ©¦                      ©¦            ©¦'
TT23 = '©¦    ©À©¤©à©¤©¤©¤©¤©¤©à©¤©¤©¤©à©¤©¤©à©¤©¤©¤©¤©È    ©À©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©È'
TT24 = '©¦    ©¦  ©¦ [93]91ºÅ ©¦      ©¦    ©¦        ©¦    ©¦  ©¦                      ©¦            ©¦'
TT25 = '©¦ Ñø ©À©¤©à©¤©¤©¤©¤©¤©Ø©¤©¤©¤©Ø©¤©¤©Ø©¤©¤©¤©¤©È ½ğ ©À©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©È'
TT26 = '©¦ ÀÏ ©À©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ğ©¤©¤©¤©¤©È    ©À©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©È'
TT27 = '©¦    ©¦9 ©¦     Ğ¡        ¼Æ       ©¦        ©¦    ©¦18©¦Éç»á±£ÏÕ»ú¹¹Ö§¸¶ÔÂ»ù±¾©¦            ©¦'
TT28 = '©¦ ½ğ ©¦  ©¦                        ©¦        ©¦    ©¦  ©¦ÑøÀÏ½ğ×Ü¶î(10+11+12+13©¦            ©¦'
if val(GZRQ)<19951001
   ½É·ÑÄêÏŞ = val(left(RQ,4))-1996+round((val(substr(RQ,6,2))+3)/12,2)
                                           ******12-10+1=3(xs)*****
else
   xs=12-val(subs(gzrq,5,2))+1
   ½É·ÑÄêÏŞ = val(left(RQ,4))-(val(½É·ÑÆğÊ¼Äê)+1)+round((val(substr(RQ,6,2))+xs)/12,2)
endif
È«²¿½É·ÑÄê = round(½É·ÑÄêÏŞ+ÊÓÍ¬½É·ÑÄê,2)
PD='N'
do while uppe(PD)='N'
@ 3 , 28 to 24 , 78
set color to gr+/b , bg+/rb
@ 4 , 29 clear to 23 , 77
@ 5 , 33 color gr+/b say '--ÔÂ¹ı¶ÉĞÔÑøÀÏ½ğºÍÔÂ¹ı¶ÉĞÔµ÷½Ú½ğ¼ÆËãÌáÊ¾--'
@ 7 , 30 colo 7+/1 say 'ÇëÈ·ÈÏ:ª±.Ö¸ÊıºÏ¼Æ(¡Æ):' get ¸öÈËÖ¸Êı
read
@ 9 , 30 colo 7+/1 say 'ª².½É·ÑÄêÊı(ĞéÄê¼õÖĞ¶Ï½É·ÑÄê):' get ½É·ÑÄêÊı
read
@ 11 , 30 colo 7+/1 say 'ª³.Æ½¾ùÉçÆ½¹¤×Ê:' get Æ½¾ùÉçÆ½
read
½É·ÑÖ¸Êı=round(¸öÈËÖ¸Êı/½É·ÑÄêÊı,4)
@ 13 , 30 colo 7+/1 say 'ª´.ÔÂÆ½¾ù½É·ÑÖ¸Êı:' get ½É·ÑÖ¸Êı
read
*ZSHGZ1 = round(¸öÈËÖ¸Êı/(½É·ÑÄêÊı*12),4)
*@ 13 , 30 clear to 13 , 70
*@ 13 , 30 colo 7+/1 say 'ª´.ÔÂÆ½¾ù½É·ÑÖ¸Êı:' get ZSHGZ1
*read
ZSHGZ = round(round(¸öÈËÖ¸Êı/½É·ÑÄêÊı,4)*Æ½¾ùÉçÆ½,¾«¶ÈÖµ)
@ 13 , 30 clear to 13 , 70
@ 13 , 30 colo 7+/1 say '*.Ö¸Êı»¯ÔÂÆ½¾ù½É·Ñ¹¤×Ê:' get ZSHGZ
read
Ö¸Êı»¯ÏµÊı=1.4
@ 13 , 30 clear to 13 , 70
@ 13 , 30 colo 7+/1 say '*.Ö¸Êı»¯¹¤×ÊÏµÊı(ª¥):' get Ö¸Êı»¯ÏµÊı
read
@ 15 , 30 colo 7+/1 say 'ªµ.Êµ¼Ê½É·ÑÄêÏŞ(ÔÂ¶ÔÔÂ):' get ½É·ÑÄêÏŞ
read
@ 17 , 30 colo 7+/1 say 'ª¶.Æ½¾ùÔÂ½É·Ñ»ùÊı:' get Êµ½ÉÆ½¾ù
read
PD='Y'
@ 20 , 46 color r+/b say 'ÕıÈ·Âğ?(Y/N)' get PD
read
if upper(PD)='Y'
   exit
endif
enddo
PD='N'
do while uppe(PD)='N'
@ 3 , 28 to 24 , 78
set color to gr+/b , bg+/rb
@ 4 , 29 clear to 23 , 77
@ 5 , 35 color r+/b say '--À¥¸Ö¼¯ÍÅ¹«Ë¾Ö°¹¤ÍËĞİ´ıÓö¼ÆËãÌáÊ¾--'
@ 7 , 33 say 'ÇëÈ·ÈÏ:ª·.'+allt(ryxm)+'ÕÛËã¹¤ÁäÏµÊı(Ğ¡ÓÚ1):' get ¹¤ÁäÏµÊı picture '9.99'
read
@ 8 , 33 say 'ª¸.'+allt(ryxm)+'´ÓÊÂÌØÊâ¹¤ÖÖÄêÏŞ(Äê):' get ¹¤ÖÖÄêÏŞ picture '99.99'
read
do while .t.
  @ 10 , 33 say 'ª¹.'+allt(ryxm)+'·â¶¨¹¤×Ê(Ôª):' get ¸Ú¼¼ picture '9999'
  read
  if ¸Ú¼¼>0
    exit
  endif
enddo
if val(GZRQ)<19980401
  do while .t.
    @ 12 , 33 say 'ª±ª°.'+allt(ryxm)+'¸öÈËÕÊ»§¾ÅÆßÄê×Ü¼Æ(Ôª):' get ¾ÅÆß picture;
 '9999.9'
    read
    if ¾ÅÆß>0
      exit
    endif
  enddo
endif
@ 14,33 clear to 14,70
@ 14 , 33 color w+/b say 'ª±ª±.'+allt(ryxm)+'µµ°¸¼ÇÔØ¹¤Áä¹¤×Ê(Ôª):' get b.GLGZ picture;
 '999.9'
read
@ 14,33 clear to 14,70
@ 14 , 33 color w+/b say 'ª±ª².'+allt(ryxm)+'µµ°¸¼ÇÔØÒ½Ò©·Ñ(Ôª):' get b.YLF picture;
 '999.9'
read
Á¬Ğø¹¤Áä = round(½É·ÑÄêÏŞ+ÊÓÍ¬½É·ÑÄê+¹¤ÖÖÄêÏŞ*¹¤ÁäÏµÊı,2)
@ 14 , 33 color w+/b say 'ÇëÈ·ÈÏ'+allt(ryxm)+'µ±Ç°Á¬Ğø¹¤Áä(Äê):' get Á¬Ğø¹¤Áä picture;
 '999.99'
read
do case
case Á¬Ğø¹¤Áä>=10 and Á¬Ğø¹¤Áä<15
  ¼Æ·¢±ÈÀı1 = '65'
case Á¬Ğø¹¤Áä>=15 and Á¬Ğø¹¤Áä<20
  ¼Æ·¢±ÈÀı1 = '70'
  ¼Æ·¢±ÈÀı3 = ' 5'
case Á¬Ğø¹¤Áä>=20 and Á¬Ğø¹¤Áä<25
  ¼Æ·¢±ÈÀı1 = '75'
  ¼Æ·¢±ÈÀı3 = ' 5'
case Á¬Ğø¹¤Áä>=25 and Á¬Ğø¹¤Áä<30 and b.XB='1'
  ¼Æ·¢±ÈÀı1 = '75'
  ¼Æ·¢±ÈÀı2 = '10'
case Á¬Ğø¹¤Áä>=25 and Á¬Ğø¹¤Áä<30 and b.XB='0'
  ¼Æ·¢±ÈÀı1 = '75'
  ¼Æ·¢±ÈÀı2 = '10'
  ¼Æ·¢±ÈÀı4 = ' 5'
case Á¬Ğø¹¤Áä>=30 and Á¬Ğø¹¤Áä<35 and b.XB='1'
  ¼Æ·¢±ÈÀı1 = '75'
  ¼Æ·¢±ÈÀı2 = '15'
case Á¬Ğø¹¤Áä>=30 and Á¬Ğø¹¤Áä<35 and b.XB='0'
  ¼Æ·¢±ÈÀı1 = '75'
  ¼Æ·¢±ÈÀı2 = '15'
  ¼Æ·¢±ÈÀı4 = ' 5'
case Á¬Ğø¹¤Áä>=35
  ¼Æ·¢±ÈÀı1 = '75'
  ¼Æ·¢±ÈÀı2 = '20'
endcase
¼Æ·¢±ÈÀı = str(val(¼Æ·¢±ÈÀı1)+val(¼Æ·¢±ÈÀı2)+val(¼Æ·¢±ÈÀı3)+val(¼Æ·¢±ÈÀı4);
,2)
@ 16 , 33 color g+/b say 'ª±ª³.'+allt(ryxm)+'¸Ú¼¼¹¤×Ê¼Æ·¢±ÈÀı(£¥):' get ¼Æ·¢±ÈÀı
read
@ 18 , 33 color r+/b say 'ª±ª´.'+allt(ryxm)+'ÊÇ·ñ½ÌÊ¦(Y/N):' get JS
read
if upper(JS)='Y'
  JSS = b.JSJT
  @ 20 , 33 say 'ÏÖĞĞ½ÌÊ¦½òÌù(Ôª):' get JSS
  read
  ¸Ú¼¼ = ¸Ú¼¼+JSS
endif
@ 18 , 33 color w+/b say 'ª±ªµ.'+allt(ryxm)+'ÊÇ·ñ¼¼Ê¦(Y/N):' get JSZ
read
if upper(JSZ)='Y'
  JT = round(b.ZCJT*val(¼Æ·¢±ÈÀı)/100,1)
  @ 20 , 33 say 'ÏÖĞĞ¼¼Ê¦½òÌù(Ôª):' get JT
  read
endif
@ 18 , 33 color gr+/b say 'ª±ª¶.'+allt(ryxm)+'ÊÇ·ñ³ÖÓĞ¶ÀÉú×ÓÅ®Ö¤(Y/N):' get DZZ
read
@ 20 , 33 say 'ª±ª·.ÍËĞİ½»Í¨²¹Ìù(Ôª):' get JBT
read
@ 20 , 33 say 'ª±ª¸.ÍËĞİÊé±¨¡¢Ï´Àí·Ñ(Ôª):' get SXF
read
PD='Y'
@ 22 , 46 color r+/b say 'ÕıÈ·Âğ?(Y/N)' get PD
read
if upper(PD)='Y'
   exit
endif
enddo
BT = JBT+SXF
if upper(DZZ)='Y'
  DZF = round(¸Ú¼¼*5/100,1)
endif
replace ¹ı¶ÉÑøÀÏ½ğ with round(ZSHGZ*Ö¸Êı»¯ÏµÊı/100*round(ÊÓÍ¬½É·ÑÄê,2),1)
*if val(GZRQ)<19951001
*   ½É·Ñµ÷½Ú½ğ = round(round((È«²¿½É·ÑÄê-15)*0.4/100*ÉÏÄêÉçÆ½,2)+round(½É·ÑÄêÏŞ*Êµ½ÉÆ½¾ù*0.6/100,2),1)
*else
*   ½É·Ñµ÷½Ú½ğ = round(½É·ÑÄêÏŞ*Êµ½ÉÆ½¾ù*0.6/100,1)
*endif
ÀÏÍËĞİ½ğ = round(¸Ú¼¼*(val(left(¼Æ·¢±ÈÀı,2))/100),1)+round(BT+JT+DZF,2)+round(¾ÅÆß/120,1)
µ÷½Ú½ğ = round(ÀÏÍËĞİ½ğ-»ù´¡ÑøÀÏ½ğ-ÕË»§ÑøÀÏ½ğ-¹ı¶ÉÑøÀÏ½ğ,1)
if µ÷½Ú½ğ<0
   µ÷½Ú½ğ = 0
endif
replace ÑøÀÏ´ıÓö¢ñ with round(»ù´¡ÑøÀÏ½ğ+ÕË»§ÑøÀÏ½ğ+µ÷½Ú½ğ,1) , ÑøÀÏ´ıÓö¢ò with »ù´¡ÑøÀÏ½ğ+ÕË»§ÑøÀÏ½ğ+¹ı¶ÉÑøÀÏ½ğ+µ÷½Ú½ğ , ÑøÀÏ´ıÓö¢ó;
 with round(»ù´¡ÑøÀÏ½ğ+ÕË»§ÑøÀÏ½ğ+ZSHGZ*Ö¸Êı»¯ÏµÊı/100*ÊÓÍ¬½É·ÑÄê,1)
@ 3 , 28 to 24 , 78
set color to gr+/b , bg+/rb
@ 4 , 29 clear to 23 , 77
D1 = 0
@ 9 , 47 prompt '1.×Ô¶¯´òÓ¡ÉóÅú±í(ÍÆ¼ö)' message '²»ĞèÒªÈË¹¤±à¼­Ö±½Ó¿ìËÙ´òÓ¡²İ¸å±í(Ç¿ÁÒÍÆ¼öÒ»°ãÓÃ»§Ê¹ÓÃ)'
@ 10 , 47 prompt '2.ÈË¹¤±à¼­ºó´òÓ¡ÉóÅú±íÕıÃæ' message 'ÏÈä¯ÀÀÑù±í²¢ÊäÈëÓĞ¹ØÄÚÈİ;
,°´[ F1 ]¼ü´æÅÌ´òÓ¡,°´[ Esc ]¼ü°ïÖú'
@ 11 , 47 prompt '3.ÈË¹¤±à¼­ºó´òÓ¡ÉóÅú±í±³Ãæ' message 'ÏÈä¯ÀÀÑù±í²¢ÊäÈëÓĞ¹ØÄÚÈİ;
,°´[ F1 ]¼ü´æÅÌ´òÓ¡,°´[ Esc ]¼ü°ïÖú'
@ 14 , 30 color r+/b say 'ËµÃ÷:1.´Ë±íÎª²İ±í,ÇëÈÏÕæ¶ÔÕÕÕıÊ½±íĞŞ¸ÄÍêÉÆ,'
@ 15 , 35 color r+/b say '×óÀ¸°´ĞòºÅÌîĞ´,ÖĞ¼ä²»Áô¿Õ°×;¹æ·¶ÌîĞ´ÄÚ'
@ 16 , 35 color r+/b say 'Èİ,Çë×ñÕÕÉÏ¼¶Òâ¼û;ÍËĞİÔ­ÒòºÍµ¥Î»Òâ¼û½ö'
@ 17 , 35 color r+/b say '¹©²Î¿¼,Çë¸ù¾İÊµ¼ÊÇé¿öÌîĞ´¡£'
@ 18 , 35 color r+/b say '2.²»Òª¶Ô±êÌâ×÷ÈÎºÎĞŞ¸Ä(Èç:"@L"µÈ·ûºÅºÍÎ»ÖÃ)!'
menu to D1
set message to 24
@ 3 , 28 to 24 , 78
set color to gr+/b , bg+/rb
@ 4 , 29 clear to 23 , 77
do case
case D1=1
  do YYJSB1
case D1=2
  set console off
  set printer on
  set device to printer
  set printer to YYJSB.TXT
  go top
  locate for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)
  @ 0 , 0 say '`@L`'+'      ÔÆÄÏÊ¡Ö°¹¤ÍËĞİ»ù±¾ÑøÀÏ±£ÏÕ´ıÓö±¨Åú±í'
  @ 1 , 0 say '`@A`'+TT1
  @ 2 , 0 say TT2
  @ 3 , 0 say '©¦ĞÕÃû©¦'+b.RYXM+'©¦ĞÔ±ğ©¦'+b.XB1+'©¦³öÉú©¦'+left(b.CSRQ,4)+'Äê'+str(val(substr(b.CSRQ;
,5,2)),2)+'ÔÂ'
  @ prow() , pcol() say '©¦½É·ÑÆğ©¦'+left(b.CJGZRQ,4)+'Äê'+str(val(substr(b.CJGZRQ;
,5,2)),2)+'ÔÂ'
  @ prow() , pcol() say '©¦ Ö°¹¤ ©¦'+b.GZ1+'        ©¦'
  @ prow()+1 , 0 say TT4
  @ prow()+1 , 0 say TT5
  @ prow()+1 , 0 say TT6
  @ prow()+1 , 0 say '©¦·ÑÄêÏŞ©¦'+left(str(ÊÓÍ¬½É·ÑÄê,5,2),2)+'ÄêÁã'+str(val(right(str(ÊÓÍ¬½É·ÑÄê;
,5,2),2))/100*12,2)+'¸öÔÂ'
  @ prow() , pcol() say '©¦·ÑÄêÏŞ©¦'+left(str(È«²¿½É·ÑÄê,5,2),2)+'Äê'+str(val(right(str(È«²¿½É·ÑÄê;
,5,2),2))/100*12,2)+'¸öÔÂ'
  @ prow() , pcol() say '©¦ ±¾Ï¢´¢´æ¶î ©¦'
  @ prow() , pcol() say ×ÜÀÛ¼Æ picture '9999999.99'
  @ prow() , pcol() say '©¦ ±¾Ï¢´¢´æ¶î ©¦'
  @ prow() , pcol() say ÀúÄê¸öÈË½É+ÀúÄêÀûÏ¢2+¸öÈË½É·Ñ+µ±ÄêÀûÏ¢ picture '9999999.99'
  @ prow() , pcol() say '©¦'
  @ prow()+1 , 0 say TT7
  @ prow()+1 , 0 say '©¦Ö°¹¤ÕÊºÅ(Éí·İÖ¤ºÅÂë)©¦'+space(3)+b.SFZ+space(1)+'©¦;
 ÍË Ğİ ÈÕ ÆÚ©¦'+space(6)+left(RQ,4)+space(1)+'Äê'+space(2)+str(val(substr(RQ;
,6,2)),2)+space(1)+'ÔÂ'+space(5)+'       ©¦'
  @ prow()+1 , 0 say TT8
  @ prow()+1 , 0 say TT9
  @ prow()+1 , 0 say TT10
  @ prow()+1 , 0 say TT11
  @ prow()+1 , 0 say '©¦ ¶¨ ©¦1 ©¦¹ú·¢[ 78 ]©¦'+str(¸Ú¼¼,6)+'©¦ '+¼Æ·¢±ÈÀı1+';
 ©¦'
  @ prow() , pcol() say round(val(¼Æ·¢±ÈÀı1)/100*¸Ú¼¼,1) picture '99999.99'
  @ prow() , pcol() say '©¦ ĞÂ ©¦10©¦ÔÂ»ù´¡ÑøÀÏ½ğ          ©¦'
  @ prow() , pcol() say »ù´¡ÑøÀÏ½ğ picture '999999999.99'
  @ prow() , pcol() say '©¦'
  @ prow()+1 , 0 say TT12
  @ prow()+1 , 0 say TT13
  if  not empty(¼Æ·¢±ÈÀı2)
    @ prow()+1 , 0 say '©¦    ©¦2 ©¦ ÔÆÀÍÈËÏÕ ©¦'+str(¸Ú¼¼,6)+'©¦ '+¼Æ·¢±ÈÀı2+';
 ©¦'
    @ prow() , pcol() say round(val(¼Æ·¢±ÈÀı2)/100*¸Ú¼¼,1) picture '99999.99'
    @ prow() , pcol() say '©¦    ©¦11©¦ÔÂ¸öÈËÕÊ»§ÑøÀÏ½ğ      ©¦'
    @ prow() , pcol() say ÕË»§ÑøÀÏ½ğ picture '999999999.99'
    @ prow() , pcol() say '©¦'
    @ prow()+1 , 0 say TT14
  else
    @ prow()+1 , 0 say '©¦    ©¦2 ©¦          ©¦      ©¦    ©¦        ©¦;
    ©¦11©¦ÔÂ¸öÈËÕË»§ÑøÀÏ½ğ      ©¦'
    @ prow() , pcol() say ÕË»§ÑøÀÏ½ğ picture '999999999.99'
    @ prow() , pcol() say '©¦'
    @ prow()+1 , 0 say '©¦ Ô­ ©¦  ©¦          ©¦      ©¦    ©¦        ©¦;
 ·¨ ©¦  ©¦                      ©¦            ©¦'
  endif
  @ prow()+1 , 0 say TT15
  @ prow()+1 , 0 say TT16
  @ prow()+1 , 0 say '©¦    ©¦  ©¦ [91] 3ºÅ ©¦   Ï´Àí·Ñ   ©¦'+str(SXF,8,2)
    @ prow() , pcol() say '©¦    ©¦  ©¦                      ©¦'
    @ prow() , pcol() say ¹ı¶ÉÑøÀÏ½ğ picture '999999999.99'
    @ prow() , pcol() say '©¦'
  @ prow()+1 , 0 say TT17
  @ prow()+1 , 0 say TT18
  @ prow()+1 , 0 say '©¦ ¼Æ ©¦  ©¦ [93]91ºÅ ©¦    ½»Í¨·Ñ  ©¦'+str(JBT,8,2)+'©¦;
 ÔÂ ©¦  ©¦                      ©¦'
    @ prow() , pcol() say µ÷½Ú½ğ picture '999999999.99'
    @ prow() , pcol() say '©¦'
  @ prow()+1 , 0 say TT19
  if DZZ$'Yy'
    @ prow()+1 , 0 say '©¦ Ëã ©¦5 ©¦ ¼Æ»®ÉúÓı ©¦'+str(¸Ú¼¼,6)+'©¦  5 ©¦'
    @ prow() , pcol() say DZF picture '99999.99'
    @ prow() , pcol() say '©¦ »ù ©¦14©¦                      ©¦         ;
   ©¦'
    @ prow()+1 , 0 say TT20
  else
    @ prow()+1 , 0 say '©¦ Ëã ©¦5 ©¦          ©¦      ©¦    ©¦        ©¦;
 »ù ©¦14©¦                      ©¦            ©¦'
    @ prow()+1 , 0 say '©¦    ©¦  ©¦          ©¦      ©¦    ©¦        ©¦;
    ©¦  ©¦                      ©¦            ©¦'
  endif
  @ prow()+1 , 0 say TT21
  if  not empty(¼Æ·¢±ÈÀı3)
    @ prow()+1 , 0 say '©¦    ©¦6 ©¦ ÔÆÀÍÈËÏÕ ©¦'+str(¸Ú¼¼,6)+'©¦ '+¼Æ·¢±ÈÀı3+' ©¦'
    @ prow() , pcol() say round(val(¼Æ·¢±ÈÀı3)/100*¸Ú¼¼,1) picture '99999.99'
    @ prow() , pcol() say '©¦    ©¦15©¦                      ©¦            ©¦'
    @ prow()+1 , 0 say TT22
  else
    @ prow()+1 , 0 say '©¦    ©¦6 ©¦          ©¦      ©¦    ©¦        '
    @ prow() , pcol() say '©¦    ©¦15©¦                      ©¦            ©¦'
    @ prow()+1 , 0 say '©¦ »ù ©¦  ©¦          ©¦      ©¦    ©¦        ©¦ Ñø ©¦  ©¦                      ©¦            ©¦'
  endif
  @ prow()+1 , 0 say TT23
  if  not empty(¼Æ·¢±ÈÀı4)
    @ prow()+1 , 0 say '©¦ ±¾ ©¦7 ©¦  ÔÆÕş·¢  ©¦'+str(¸Ú¼¼,6)+'©¦ '+¼Æ·¢±ÈÀı4+';
 ©¦'
    @ prow() , pcol() say round(val(¼Æ·¢±ÈÀı4)/100*¸Ú¼¼,1) picture '99999.99'
    @ prow() , pcol() say '©¦ ÀÏ ©¦16©¦                      ©¦            ©¦'
    @ prow()+1 , 0 say TT24
  else
    @ prow()+1 , 0 say '©¦ ±¾ ©¦7 ©¦          ©¦      ©¦    ©¦        ©¦;
 ÀÏ ©¦16©¦                      ©¦            ©¦'
    @ prow()+1 , 0 say '©¦    ©¦  ©¦          ©¦      ©¦    ©¦        ©¦;
    ©¦  ©¦                      ©¦            ©¦'
  endif
  @ prow()+1 , 0 say TT25
  @ prow()+1 , 0 say '©¦    ©¦8 ©¦1998Äê·â¶¨Ç°¸öÈËÕË»§Óà¶î/120='
  @ prow() , pcol() say round(¾ÅÆß/120,1) picture '999.9'
  @ prow() , pcol() say '©¦    ©¦'
  if upper(JSZ)='Y'
    @ prow() , pcol() say '17©¦¼¼Ê¦½òÌù(ÌîÈë×ó¿Õ°×À¸)©¦'
    @ prow() , pcol() say JT picture '999999999.99'
    @ prow() , pcol() say '©¦'
  else
    @ prow() , pcol() say '17©¦                      ©¦            ©¦'
  endif
  @ prow()+1 , 0 say TT26
  @ prow()+1 , 0 say TT27
  @ prow()+1 , 0 say TT28
  @ prow()+1 , 0 say '©¦    ©¦  ©¦  (1+2+3+4+5+6+7+8)     ©¦'
  @ prow() , pcol() say ÀÏÍËĞİ½ğ picture '99999.99'
  @ prow() , pcol() say '©¦    ©¦  ©¦+14+15+16+17)         ©¦'
  @ prow() , pcol() say round(ÑøÀÏ´ıÓö¢ò,1) picture '999999999.99'
  @ prow() , pcol() say '©¦'
  if (b.NL<50 and b.XB='0') or (b.NL<60 and b.XB='1')
    @ prow()+1 , 0 say '©À©¤©¤©à©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©Ø©¤©¤©Ø©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©È'
    @ prow()+1 , 0 say '©¦    ©¦                                                                                      ©¦'
    @ prow()+1 , 0 say '©¦ ÍË ©¦                                        ;
                                              ©¦'
    @ prow()+1 , 0 say '©¦ Ğİ ©¦    Äê  ÔÂÖÁ  Äê  ÔÂÔÚ           ´ÓÊÂ   ;
      ,ÀÛ¼ÆÂú  Äê ¸öÔÂ,·ûºÏ[1978]104ºÅÎÄ¹æ¶¨µÄ©¦'
    @ prow()+1 , 0 say '©¦ Ô­ ©¦                                        ;
                                              ©¦'
    @ prow()+1 , 0 say '©¦ Òò ©¦           ÌØÊâ¹¤ÖÖ¡£ÕÛËã   ¹¤Áä     Äê ;
   ¸öÔÂ¡£ÉêÇë°´ÌØÊâ¹¤ÖÖÍËĞİ¡£                 ©¦'
    @ prow()+1 , 0 say '©¦ ¼° ©¦                                        ;
                                              ©¦'
    @ prow()+1 , 0 say '©¦ Òª ©¦                                        ;
                                              ©¦'
    @ prow()+1 , 0 say '©¦ Çó ©¦                                                                                      ©¦'
    @ prow()+1 , 0 say '©¦    ©¦                                                                                      ©¦'
    @ prow()+1 , 0 say '©¸©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¼'
  else
    @ prow()+1 , 0 say '©À©¤©¤©à©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©Ø©¤©¤©Ø©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©È'
    @ prow()+1 , 0 say '©¦    ©¦                                                                                      ©¦'
    @ prow()+1 , 0 say '©¦ ÍË ©¦       ÄêÁäÂú    Ëê   ¸öÔÂ,ÒÑ´ïµ½Õı³£ÍËĞİÄêÁä,ÉêÇë°´Õı³£ÍËĞİ¡£                        ©¦'
    @ prow()+1 , 0 say '©¦ Ğİ ©¦                                        ;
                                              ©¦'
    @ prow()+1 , 0 say '©¦ Ô­ ©¦                                        ;
                                              ©¦'
    @ prow()+1 , 0 say '©¦ Òò ©¦                                        ;
                                              ©¦'
    @ prow()+1 , 0 say '©¦ ¼° ©¦                                        ;
                                              ©¦'
    @ prow()+1 , 0 say '©¦ Òª ©¦                                        ;
                                              ©¦'
    @ prow()+1 , 0 say '©¦ Çó ©¦                                                                                      ©¦'
    @ prow()+1 , 0 say '©¦    ©¦                                                                                      ©¦'
    @ prow()+1 , 0 say '©¸©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¼'
  endif
  @ prow()+1 , 0 say ' '
  @ 3 , 28 to 24 , 78
  set color to gr+/b , bg+/rb
  @ 4 , 29 clear to 23 , 77
  set printer off
  set device to screen
  run cced YYJSB.txt
  do srfs.spx
  I = 1
  do while I<=val(YX)
    @ 12 , 36 say 'ÕıÔÚ´òÓ¡ÑøÀÏ±£ÏÕ´ıÓöÉóÅú±í,ÇëÉÔºò...'
    run copy YYJSB.txt prn>NUL
    @ 13 , 36 say 'Çë»»Ö½ÕÅºó°´ÈÎÒâ¼ü¼ÌĞø... ...'
    =inkey(0)
    set console off
    I = I+1
  enddo
case D1=3
  do BMD2
endcase
set printer off
set device to screen
close all
ÈË¿É = '00'
delete file jcsj.idx
delete file lnk.idx
return

PROCEDURE BMD2
@ 13 , 36 say 'Çë·­¹ıÖ½ÕÅ±³Ãæ¼ÌĞø´òÓ¡... ...'
= inkey(0,'M')
set printer on
set device to printer
set printer to BMD2.TXT
if (b.NL<50 and b.XB='0') or (b.NL<60 and b.XB='1')
  @ 0 , 0 say '©°©¤©¤©¤©¤©Ğ©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©´'
  @ prow()+1 , 0 say '©¦   µ¥   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦       ¾­ÉóºË,·ûºÏ      ÌØÊâ¹¤ÖÖÍËĞİÌõ¼ş;
,Í¬ÒâÍËĞİ¡£                               ©¦'
  @ prow()+1 , 0 say '©¦   Î»   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦   Òâ   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦   ¼û   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                Äê        ÔÂ      ÈÕ        ©¦'
else
  @ 0 , 0 say '©°©¤©¤©¤©¤©Ğ©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©´'
  @ prow()+1 , 0 say '©¦   µ¥   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦       ¾­ÉóºË,ÒÑ´ïµ½Õı³£ÍËĞİÄêÁä,Í¬ÒâÍËĞİ¡£;
                                       ©¦'
  @ prow()+1 , 0 say '©¦   Î»   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦   Òâ   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦   ¼û   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                Äê        ÔÂ      ÈÕ        ©¦'
endif
@ prow()+1 , 0 say '©À©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©È'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ Ö÷  Òâ ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ¹Ü     ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ²¿     ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ÃÅ  ¼û ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
              Äê        ÔÂ      ÈÕ        ©¦'
@ prow()+1 , 0 say '©À©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©È'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ Éç  »ú ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ »á  ¹¹ ©¦                 Í¬Òâ´Ó    Äê   ÔÂ¿ªÊ¼ÏíÊÜ;
                                        ©¦'
@ prow()+1 , 0 say '©¦ ±£  Éó ©¦                 ÔÂ»ù±¾ÑøÀÏ½ğ       Ôª¡£;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ÏÕ  ºË ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
              Äê        ÔÂ      ÈÕ        ©¦'
@ prow()+1 , 0 say '©À©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©È'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ÀÍ  ²¿ ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ¶¯  ÃÅ ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ±£  Åú ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ÕÏ  ×¼ ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
              Äê        ÔÂ      ÈÕ        ©¦'
@ prow()+1 , 0 say '©À©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©È'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦      ¸ù¾İÀ¥Ã÷¸ÖÌú×Ü¹«Ë¾ÀÍÈË×Ö[99]036ºÅÎÄ¹æ¶¨;
,ÒÔÏÂ½òÌù²»Êô»ù±¾ÑøÀÏ±£ÏÕÍ³³ïÏîÄ¿,   ©¦'
@ prow()+1 , 0 say '©¦   ±¸   ©¦  ÊÇÓÉÆóÒµ¸ù¾İĞ§ÒæÇé¿ö,Ö§¸¶ÄÜÁ¦×ÔĞĞ·¢·ÅµÄ´ıÓö¡£;
                                   ©¦'
@ prow()+1 , 0 say '©¦        ©¦  ¹¤Áä¹¤×Ê'+str(b.GLGZ+b.ylf,5,1)+'Ôª¡¢ÃºÌù'+str(c.MT;
,2)+'Ôª¡¢Êé±¨½»Í¨Ï´Àí·Ñ'+str(c.JTXLF,2)+'Ôª¡¢Ë®µç²¹Ìù'+str(c.SDBT,2)+'Ôª;
,³§ÄÚ'+str(c.chn,3)+'Ôª,ºÏ¼Æ'+str(b.GLGZ+b.ylf+c.MT+c.JTXLF+c.SDBT+c.chn,5,1)+'Ôª©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦   ×¢   ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦  ×¢:Ò½ÁÆ±£ÏÕ½É·Ñ»ùÊı='+str(ÑøÀÏ´ıÓö¢ò+b.GLGZ+c.MT+c.JTXLF+c.SDBT+c.chn,6,1)+'Ôª'+space(10)+';
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¸©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¼'
@ prow()+1 , 0 say '×¢:´Ë±íÒ»Ê½Èı·İ,¾­Åú×¼ºó,Ò»·İ´æÔ­µ¥Î»,'
@ prow()+1 , 3 say 'Ò»·İÉç»á±£ÏÕ»ú¹¹Áô´æ,Ò»·İÀÍ¶¯ºÍÉç»á±£ÕÏ²¿ÃÅÁô´æ¡£'
@ prow()+1 , 0 say ' '
@ 3 , 28 to 24 , 78
set color to gr+/b , bg+/rb
@ 4 , 29 clear to 23 , 77
set printer off
set device to screen
run cced BMD2.txt
do srfs.spx
I = 1
pd='Y'
@ 13 , 36 say '½¨Òé´òÓ¡¿Õ±í¸ñ,Í¬ÒâÂğ(Y/N)?' get pd
read
do while I<=val(YX)
  @ 12 , 31 say 'Õı ÔÚ ´ò Ó¡ ×é Ö¯ Éó Åú Òâ ¼û Ñù ±í... ...'
  if uppe(pd)='Y'
     run copy YYJSB1.txt prn>NUL
  else
     run copy BMD2.txt prn>NUL
  endi
  @ 14 , 36 say 'Çë»»Ö½ÕÅºó°´ÈÎÒâ¼ü¼ÌĞø... ...'
    =inkey(0)
  set console off
  I = I+1
enddo
return

PROCEDURE YYJSB1
on error do error
DH1 = '2'
do DHK1
if DH1=1
  set printer on
  set device to printer
  set printer to yyjsb.TXT
else
  do DHK162
  @ 12 , 38 say 'Õı  ÔÚ  ´ò  Ó¡,Çë  ÉÔ  ºò!... ...'
  set device to printer
  set printer on
endif
go top
locate for left(JBXM,8)=left(RYXM,8) or left(JBXM,5)=left(RYBH,5)
@ 0 , 0 say '`@L`'+'ÔÆÄÏÊ¡Ö°¹¤ÍËĞİ»ù±¾ÑøÀÏ±£ÏÕ´ıÓö±¨Åú±í'+'`@B`'+'(ÔÆÑøÏÕËÄ±í)'
@ 1 , 0 say '`@a`'+'(ËµÃ÷:´Ë±íÎª²İ¸å±í,ÇëÖØ³­ÓÚÕıÊ½±í,ÒªÇó×óÀ¸°´ĞòºÅ³­Ìî;
,ÖĞ¼ä²»Áô¿Õ°×;¹æ·¶ÌîĞ´ÄÚÈİ,Çë×ñÕÕÉÏ¼¶Òâ¼û;ÍËĞİÔ­ÒòºÍµ¥Î»Òâ¼û½ö¹©²Î¿¼,Çë¸ù¾İÊµ¼ÊÇé¿öÌîĞ´)'
@ 2 , 0 say '`@A`'+TT1
@ 3 , 0 say TT2
@ 4 , 0 say '©¦ĞÕÃû©¦'+b.RYXM+'©¦ĞÔ±ğ©¦'+b.XB1+'©¦³öÉú©¦'+left(b.CSRQ,4)+'Äê'+str(val(substr(b.CSRQ;
,5,2)),2)+'ÔÂ'
@ prow() , pcol() say '©¦²Î¼Ó¹¤©¦'+left(b.CJGZRQ,4)+'Äê'+str(val(substr(b.CJGZRQ;
,5,2)),2)+'ÔÂ'
@ prow() , pcol() say '©¦ Ö°¹¤ ©¦'+b.GZ1+'          ©¦'
@ prow()+1 , 0 say TT4
@ prow()+1 , 0 say TT5
@ prow()+1 , 0 say TT6
@ prow()+1 , 0 say '©¦·ÑÄêÏŞ©¦'+left(str(ÊÓÍ¬½É·ÑÄê,5,2),2)+'ÄêÁã'+str(val(right(str(ÊÓÍ¬½É·ÑÄê;
,5,2),2))/100*12,2)+'¸öÔÂ'
@ prow() , pcol() say '©¦·ÑÄêÏŞ©¦'+left(str(È«²¿½É·ÑÄê,5,2),2)+'Äê'+str(val(right(str(È«²¿½É·ÑÄê;
,5,2),2))/100*12,2)+'¸öÔÂ'
@ prow() , pcol() say '©¦ ±¾Ï¢´¢´æ¶î ©¦'
@ prow() , pcol() say ×ÜÀÛ¼Æ picture '9999999.99'
@ prow() , pcol() say '©¦ ±¾Ï¢´¢´æ¶î ©¦'
@ prow() , pcol() say ÀúÄê¸öÈË½É+ÀúÄêÀûÏ¢2+¸öÈË½É·Ñ+µ±ÄêÀûÏ¢ picture '9999999.99'
@ prow() , pcol() say '©¦'
@ prow()+1 , 0 say TT7
@ prow()+1 , 0 say '©¦Ö°¹¤ÕÊºÅ(Éí·İÖ¤ºÅÂë)©¦'+space(3)+b.SFZ+space(1)+'©¦;
 ÍË Ğİ ÈÕ ÆÚ©¦'+space(6)+left(RQ,4)+space(1)+'Äê'+space(2)+str(val(substr(RQ;
,6,2)),2)+space(1)+'ÔÂ'+space(5)+'       ©¦'
@ prow()+1 , 0 say TT8
@ prow()+1 , 0 say TT9
@ prow()+1 , 0 say TT10
@ prow()+1 , 0 say TT11
@ prow()+1 , 0 say '©¦ ¶¨ ©¦1 ©¦¹ú·¢[ 78 ]©¦'+str(¸Ú¼¼,6)+'©¦ '+¼Æ·¢±ÈÀı1+';
 ©¦'
@ prow() , pcol() say round(val(¼Æ·¢±ÈÀı1)/100*¸Ú¼¼,1) picture '99999.99'
@ prow() , pcol() say '©¦ ĞÂ ©¦10©¦ÔÂ»ù´¡ÑøÀÏ½ğ          ©¦'
@ prow() , pcol() say »ù´¡ÑøÀÏ½ğ picture '999999999.99'
@ prow() , pcol() say '©¦'
@ prow()+1 , 0 say TT12
@ prow()+1 , 0 say TT13
if  not empty(¼Æ·¢±ÈÀı2)
  @ prow()+1 , 0 say '©¦    ©¦2 ©¦ ÔÆÀÍÈËÏÕ ©¦'+str(¸Ú¼¼,6)+'©¦ '+¼Æ·¢±ÈÀı2+';
 ©¦'
  @ prow() , pcol() say round(val(¼Æ·¢±ÈÀı2)/100*¸Ú¼¼,1) picture '99999.99'
  @ prow() , pcol() say '©¦    ©¦11©¦ÔÂ¸öÈËÕÊ»§ÑøÀÏ½ğ      ©¦'
  @ prow() , pcol() say ÕË»§ÑøÀÏ½ğ picture '999999999.99'
  @ prow() , pcol() say '©¦'
  @ prow()+1 , 0 say TT14
else
  @ prow()+1 , 0 say '©¦    ©¦2 ©¦          ©¦      ©¦    ©¦        ©¦  ;
  ©¦11©¦ÔÂ¸öÈËÕË»§ÑøÀÏ½ğ      ©¦'
  @ prow() , pcol() say ÕË»§ÑøÀÏ½ğ picture '999999999.99'
  @ prow()+1 , 0 say '©¦ Ô­ ©¦  ©¦          ©¦      ©¦    ©¦        ©¦ ·¨;
 ©¦  ©¦                      ©¦            ©¦'
endif
  @ prow()+1 , 0 say TT15
  @ prow()+1 , 0 say TT16
  @ prow()+1 , 0 say '©¦    ©¦  ©¦ [91] 3ºÅ ©¦   Ï´Àí·Ñ   ©¦'+str(SXF,8,2)
    @ prow() , pcol() say '©¦    ©¦  ©¦                      ©¦'
    @ prow() , pcol() say ¹ı¶ÉÑøÀÏ½ğ picture '999999999.99'
    @ prow() , pcol() say '©¦'
  @ prow()+1 , 0 say TT17
  @ prow()+1 , 0 say TT18
  @ prow()+1 , 0 say '©¦ ¼Æ ©¦  ©¦ [93]91ºÅ ©¦    ½»Í¨·Ñ  ©¦'+str(JBT,8,2)+'©¦;
 ÔÂ ©¦  ©¦                      ©¦'
    @ prow() , pcol() say µ÷½Ú½ğ picture '999999999.99'
    @ prow() , pcol() say '©¦'
  @ prow()+1 , 0 say TT19
if DZZ$'Yy'
  @ prow()+1 , 0 say '©¦ Ëã ©¦5 ©¦ ¼Æ»®ÉúÓı ©¦'+str(¸Ú¼¼,6)+'©¦  5 ©¦'
  @ prow() , pcol() say DZF picture '99999.99'
  @ prow() , pcol() say '©¦ »ù ©¦14©¦                      ©¦           ;
 ©¦'
  @ prow()+1 , 0 say TT20
else
  @ prow()+1 , 0 say '©¦ Ëã ©¦5 ©¦          ©¦      ©¦    ©¦        ©¦ »ù;
 ©¦14©¦                      ©¦            ©¦'
  @ prow()+1 , 0 say '©¦    ©¦  ©¦          ©¦      ©¦    ©¦        ©¦  ;
  ©¦  ©¦                      ©¦            ©¦'
endif
  @ prow()+1 , 0 say TT21
  if  not empty(¼Æ·¢±ÈÀı3)
    @ prow()+1 , 0 say '©¦    ©¦6 ©¦ ÔÆÀÍÈËÏÕ ©¦'+str(¸Ú¼¼,6)+'©¦ '+¼Æ·¢±ÈÀı3+' ©¦'
    @ prow() , pcol() say round(val(¼Æ·¢±ÈÀı3)/100*¸Ú¼¼,1) picture '99999.99'
    @ prow() , pcol() say '©¦    ©¦15©¦                      ©¦            ©¦'
    @ prow()+1 , 0 say TT22
  else
    @ prow()+1 , 0 say '©¦    ©¦6 ©¦          ©¦      ©¦    ©¦        '
    @ prow() , pcol() say '©¦    ©¦15©¦                      ©¦            ©¦'
    @ prow()+1 , 0 say '©¦ »ù ©¦  ©¦          ©¦      ©¦    ©¦        ©¦ Ñø ©¦  ©¦                      ©¦            ©¦'
  endif
  @ prow()+1 , 0 say TT23
  if  not empty(¼Æ·¢±ÈÀı4)
    @ prow()+1 , 0 say '©¦ ±¾ ©¦7 ©¦  ÔÆÕş·¢  ©¦'+str(¸Ú¼¼,6)+'©¦ '+¼Æ·¢±ÈÀı4+';
 ©¦'
    @ prow() , pcol() say round(val(¼Æ·¢±ÈÀı4)/100*¸Ú¼¼,1) picture '99999.99'
    @ prow() , pcol() say '©¦ ÀÏ ©¦16©¦                      ©¦            ©¦'
    @ prow()+1 , 0 say TT24
  else
    @ prow()+1 , 0 say '©¦ ±¾ ©¦7 ©¦          ©¦      ©¦    ©¦        ©¦;
 ÀÏ ©¦16©¦                      ©¦            ©¦'
    @ prow()+1 , 0 say '©¦    ©¦  ©¦          ©¦      ©¦    ©¦        ©¦;
    ©¦  ©¦                      ©¦            ©¦'
  endif
  @ prow()+1 , 0 say TT25
@ prow()+1 , 0 say '©¦    ©¦8 ©¦1998Äê·â¶¨Ç°¸öÈËÕË»§Óà¶î/120='
@ prow() , pcol() say round(¾ÅÆß/120,1) picture '999.9'
@ prow() , pcol() say '©¦    ©¦'
if upper(JSZ)='Y'
  @ prow() , pcol() say '17©¦¼¼Ê¦½òÌù(ÌîÈë×ó¿Õ°×À¸)©¦'
  @ prow() , pcol() say JT picture '999999999.99'
  @ prow() , pcol() say '©¦'
else
  @ prow() , pcol() say '17©¦                      ©¦            ©¦'
endif
@ prow()+1 , 0 say TT26
@ prow()+1 , 0 say TT27
@ prow()+1 , 0 say TT28
@ prow()+1 , 0 say '©¦    ©¦  ©¦  (1+2+3+4+5+6+7+8)     ©¦'
@ prow() , pcol() say ÀÏÍËĞİ½ğ picture '99999.99'
@ prow() , pcol() say '©¦    ©¦  ©¦+14+15+16+17)         ©¦'
@ prow() , pcol() say round(ÑøÀÏ´ıÓö¢ò,1) picture '999999999.99'
@ prow() , pcol() say '©¦'
if (b.NL<50 and b.XB='0') or (b.NL<60 and b.XB='1')
  @ prow()+1 , 0 say '©À©¤©¤©à©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©Ø©¤©¤©Ø©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©È'
    @ prow()+1 , 0 say '©¦    ©¦                                                                                      ©¦'
  @ prow()+1 , 0 say '©¦ ÍË ©¦                                          ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦ Ğİ ©¦    Äê  ÔÂÖÁ  Äê  ÔÂÔÚ           ´ÓÊÂ     ;
    ,ÀÛ¼ÆÂú  Äê ¸öÔÂ,·ûºÏ[1978]104ºÅÎÄ¹æ¶¨µÄ©¦'
  @ prow()+1 , 0 say '©¦ Ô­ ©¦                                          ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦ Òò ©¦           ÌØÊâ¹¤ÖÖ¡£ÕÛËã   ¹¤Áä     Äê   ;
 ¸öÔÂ¡£ÉêÇë°´ÌØÊâ¹¤ÖÖÍËĞİ¡£                 ©¦'
  @ prow()+1 , 0 say '©¦ ¼° ©¦                                          ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦ Òª ©¦                                          ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦ Çó ©¦                                                                                      ©¦'
    @ prow()+1 , 0 say '©¦    ©¦                                                                                      ©¦'
  @ prow()+1 , 0 say '©¸©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¼'
else
  @ prow()+1 , 0 say '©À©¤©¤©à©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©Ø©¤©¤©Ø©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©È'
    @ prow()+1 , 0 say '©¦    ©¦                                                                                      ©¦'
  @ prow()+1 , 0 say '©¦ ÍË ©¦       ÄêÁäÂú    Ëê   ¸öÔÂ,ÒÑ´ïµ½Õı³£ÍËĞİÄêÁä;
,ÉêÇë°´Õı³£ÍËĞİ¡£                        ©¦'
  @ prow()+1 , 0 say '©¦ Ğİ ©¦                                          ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦ Ô­ ©¦                                          ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦ Òò ©¦                                          ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦ ¼° ©¦                                          ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦ Òª ©¦                                          ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦ Çó ©¦                                          ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦    ©¦                                                                                      ©¦'
  @ prow()+1 , 0 say '©¸©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¼'
endif
@ prow()+1 , 0 say ' '
if D1=1 and DH1=2
  set printer off
  set device to screen
  eject
  @ 12 , 38 say '    Çë·­¹ıÖ½ÕÅ±³Ãæ¼ÌĞø´òÓ¡'
  = inkey(0,'M')
  @ 12 , 31 say 'Õı ÔÚ ´ò Ó¡ ×é Ö¯ Éó Åú Òâ ¼û Ñù ±í... ...'
  set device to printer
  set printer on
endif
if (b.NL<50 and b.XB='0') or (b.NL<60 and b.XB='1')
  @ 0 , 0 say '©°©¤©¤©¤©¤©Ğ©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©´'
  @ prow()+1 , 0 say '©¦   µ¥   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦       ¾­ÉóºË,·ûºÏ      ÌØÊâ¹¤ÖÖÍËĞİÌõ¼ş;
,Í¬ÒâÍËĞİ¡£                               ©¦'
  @ prow()+1 , 0 say '©¦   Î»   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦   Òâ   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦   ¼û   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                Äê        ÔÂ      ÈÕ        ©¦'
else
  @ 0 , 0 say '©°©¤©¤©¤©¤©Ğ©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©´'
  @ prow()+1 , 0 say '©¦   µ¥   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦       ¾­ÉóºË,ÒÑ´ïµ½Õı³£ÍËĞİÄêÁä,Í¬ÒâÍËĞİ¡£;
                                       ©¦'
  @ prow()+1 , 0 say '©¦   Î»   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦   Òâ   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦   ¼û   ©¦                                      ;
                                            ©¦'
  @ prow()+1 , 0 say '©¦        ©¦                                      ;
                Äê        ÔÂ      ÈÕ        ©¦'
endif
@ prow()+1 , 0 say '©À©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©È'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ Ö÷  Òâ ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ¹Ü     ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ²¿     ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ÃÅ  ¼û ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
              Äê        ÔÂ      ÈÕ        ©¦'
@ prow()+1 , 0 say '©À©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©È'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ Éç  »ú ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ »á  ¹¹ ©¦                 Í¬Òâ´Ó    Äê   ÔÂ¿ªÊ¼ÏíÊÜ;
                                        ©¦'
@ prow()+1 , 0 say '©¦ ±£  Éó ©¦                 ÔÂ»ù±¾ÑøÀÏ½ğ       Ôª¡£;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ÏÕ  ºË ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
              Äê        ÔÂ      ÈÕ        ©¦'
@ prow()+1 , 0 say '©À©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©È'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ÀÍ  ²¿ ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ¶¯  ÃÅ ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ±£  Åú ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦ ÕÏ  ×¼ ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
              Äê        ÔÂ      ÈÕ        ©¦'
@ prow()+1 , 0 say '©À©¤©¤©¤©¤©à©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©È'
@ prow()+1 , 0 say '©¦        ©¦                                                                                  ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                                                                  ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                                                                  ©¦'
@ prow()+1 , 0 say '©¦        ©¦      ¸ù¾İÀ¥Ã÷¸ÖÌú×Ü¹«Ë¾ÀÍÈË×Ö[99]036ºÅÎÄ¹æ¶¨;
,ÒÔÏÂ½òÌù²»Êô»ù±¾ÑøÀÏ±£ÏÕÍ³³ïÏîÄ¿,   ©¦'
@ prow()+1 , 0 say '©¦   ±¸   ©¦  ÊÇÓÉÆóÒµ¸ù¾İĞ§ÒæÇé¿ö,Ö§¸¶ÄÜÁ¦×ÔĞĞ·¢·ÅµÄ´ıÓö¡£;
                                   ©¦'
@ prow()+1 , 0 say '©¦        ©¦  ¹¤Áä¹¤×Ê'+alltrim(str(b.GLGZ+b.YLF;
,10,1))+'Ôª¡¢ÃºÌù'+alltrim(str(c.MT))+'Ôª¡¢Êé±¨½»Í¨Ï´Àí·Ñ'+alltrim(str(c.JTXLF))+'Ôª¡¢Ë®µç²¹Ìù'+alltrim(str(c.SDBT))+'Ôª;
,³§ÄÚ'+str(c.chn,3)+'Ôª,ºÏ¼Æ'+alltrim(str(b.GLGZ+b.YLF+c.MT+c.JTXLF+c.SDBT+c.chn,10,1))+'Ôª©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦   ×¢   ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦  ×¢:Ò½ÁÆ±£ÏÕ½É·Ñ»ùÊı='+str(ÑøÀÏ´ıÓö¢ò+b.GLGZ+c.MT+c.JTXLF+c.SDBT+c.chn,6,1)+'Ôª'+space(10)+';
                                          ©¦'
@ prow()+1 , 0 say '©¦        ©¦                                        ;
                                          ©¦'
@ prow()+1 , 0 say '©¸©¤©¤©¤©¤©Ø©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¤©¼'
@ prow()+1 , 0 say '×¢:´Ë±íÒ»Ê½Èı·İ,¾­Åú×¼ºó,Ò»·İ´æÔ­µ¥Î»,'
@ prow()+1 , 3 say 'Ò»·İÉç»á±£ÏÕ»ú¹¹Áô´æ,Ò»·İÀÍ¶¯ºÍÉç»á±£ÕÏ²¿ÃÅÁô´æ¡£'
@ prow()+1 , 0 say ' '
close all
ÈË¿É = '00'
@ 3 , 28 to 24 , 78
set color to gr+/b , bg+/rb
@ 4 , 29 clear to 23 , 77
delete file jcsj.idx
delete file lnk.idx
set printer off
set device to screen
if DH1=1 and D1=1
  run cced yyjsb.txt
endif
return
ÿ Zÿ   L    ) 3 = G Q [ e o y ƒ  — ¡ « µ ¿ É Ó İ ç ñ û #-   á¹á                                     