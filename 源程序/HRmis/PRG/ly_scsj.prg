***************************
* .\LY_SCSJ.PRG
***************************
if file('yl'+LY+'.dbf') or file('sy'+LY+'.dbf')
yn = MESSAGEBOX('ɾ��'+alltrim(LY)+'�����ϱ����ļ�(YL'+alltrim(LY)+')��',4+32,"��ʾ")
IF yn = 6
    delete file yl&LY..dbf
    delete file yl&LY..idx
    delete file ylbx&LY..dbf
    delete file ylbx&LY..idx
    delete file lnk&LY..dbf
    delete file lnk&LY..idx
=MESSAGEBOX("����ɾ���ɹ���",48,"��ϲ")
endif
if file('sy'+LY+'.dbf')
yn = MESSAGEBOX('ɾ��'+alltrim(LY)+'��ʧҵ�����ļ�(SY'+alltrim(LY)+'.DBF)��',4+32,"��ʾ")
IF yn = 6
    delete file sy&LY..dbf
    delete file sy&LY..idx     
=MESSAGEBOX("����ɾ���ɹ���",48,"��ϲ")
endif
endif
else
=MESSAGEBOX("���ݲ����ڣ�",48,"��ʾ")
endif