***************************
* .\LY_SCSJ.PRG
***************************
if file('kq'+LY+'.dbf')
yn = MESSAGEBOX('ɾ��'+alltrim(LY)+'��ʧҵ�����ļ�(SY'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
IF yn = 6
    delete file sy&LY..dbf
    delete file sy&LY..idx     
=MESSAGEBOX("����ɾ���ɹ���",48,"��ʼ��")
endif
*xn=str(val(LY)+1,4)
*if file('zr1'+LY+'.dbf')
*yn = MESSAGEBOX('ɾ��'+alltrim(xn)+'��ɷѻ����ļ�(ZR1'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
*IF yn = 6
 *   delete file zr1&LY..dbf
  *  delete file zr1&LY..idx
   * delete file gzze&LY..dbf
    *delete file gzze&LY..idx
*=MESSAGEBOX("����ɾ���ɹ���",48,"��ʼ��")
*endif
*endif
else
=MESSAGEBOX("���ݲ����ڣ�",48,"��ʾ")
endif