***************************
* .\LY_SCSJ.PRG
***************************
if file('zr'+LY+'.dbf')
yn = MESSAGEBOX('ɾ��'+alltrim(LY)+'�깤���ܶ��ļ�(ZR'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
IF yn = 6
    delete file zr&LY..dbf
    delete file zr&LY..idx     
=MESSAGEBOX("����ɾ���ɹ���",48,"��ϲ")
endif
xn=str(val(LY)+1,4)
if file('zr1'+LY+'.dbf')
yn = MESSAGEBOX('ɾ��'+alltrim(xn)+'��ɷѻ����ļ�(ZR1'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
IF yn = 6
    delete file zr1&LY..dbf
    delete file zr1&LY..idx
    delete file gzze&LY..dbf
    delete file gzze&LY..idx
=MESSAGEBOX("����ɾ���ɹ���",48,"��ϲ")
endif
endif
if file('xzjs'+LY+'.dbf')
yn = MESSAGEBOX('ɾ��'+alltrim(LY)+'��������Ա�ɷѻ����ļ�(XZJS'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
IF yn = 6
    delete file xzjs&LY..dbf
    delete file xzjs&LY..idx     
=MESSAGEBOX("����ɾ���ɹ���",48,"��ϲ")
endif
endif
else
=MESSAGEBOX("���ݲ����ڣ�",48,"��ʾ")
endif