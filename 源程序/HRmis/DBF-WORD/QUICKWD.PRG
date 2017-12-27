*!******************************************************************************
*!
*! 过程 FORCEEXT
*!
*!******************************************************************************
FUNCTION ForceExt
PARAMETERS filname,ext
PRIVATE ALL
IF SUBSTR(m.ext,1,1) = "."
	m.ext = SUBSTR(m.ext,2,3)
ENDIF
m.pname = justpath(m.filname)
m.filname = justfname(UPPER(ALLTRIM(m.filname)))
IF AT('.',m.filname) > 0
	m.filname = SUBSTR(m.filname,1,AT('.',m.filname)-1) + '.' + m.ext
ELSE
	m.filname = m.filname + '.' + m.ext
ENDIF
RETURN addbs(m.pname) + m.filname
*!******************************************************************************
*!
*! 过程 DEFAULTEXT
*!
*!******************************************************************************
FUNCTION defaultext
PARAMETERS filname,ext
PRIVATE ALL
IF EMPTY(justext(m.filname))
	IF SUBSTR(m.ext,1,1) = "."
		m.ext = SUBSTR(m.ext,2,3)
	ENDIF
	RETURN m.filname + '.' + m.ext
ELSE
	RETURN filname
ENDIF
*!******************************************************************************
*!
*! 过程 JUSTFNAME
*!
*!******************************************************************************
FUNCTION justfname
PARAMETERS filname
PRIVATE ALL
IF RAT('\',m.filname) > 0
	m.filname = SUBSTR(m.filname,RAT('\',m.filname)+1,255)
ENDIF
IF AT(':',m.filname) > 0
	m.filname = SUBSTR(m.filname,AT(':',m.filname)+1,255)
ENDIF
RETURN ALLTRIM(UPPER(m.filname))
*!******************************************************************************
*!
*! 过程 JUSTSTEM
*!
*!******************************************************************************
FUNCTION juststem
PARAMETERS filname
PRIVATE ALL
IF RAT('\',m.filname) > 0
	m.filname = SUBSTR(m.filname,RAT('\',m.filname)+1,255)
ENDIF
IF AT(':',m.filname) > 0
	m.filname = SUBSTR(m.filname,AT(':',m.filname)+1,255)
ENDIF
IF AT('.',m.filname) > 0
	m.filname = SUBSTR(m.filname,1,AT('.',m.filname)-1)
ENDIF
RETURN ALLTRIM(UPPER(m.filname))
*!******************************************************************************
*!
*! 过程 JUSTEXT
*!
*!******************************************************************************
FUNCTION justext
PARAMETERS filname
PRIVATE ALL
filname = justfname(m.filname)
m.ext = ""
IF AT('.',m.filname) > 0
	m.ext = SUBSTR(m.filname,AT('.',m.filname)+1,3)
ENDIF
RETURN UPPER(m.ext)
*!******************************************************************************
*!
*! 过程 JUSTPATH
*!
*!******************************************************************************
FUNCTION justpath
PARAMETERS filname
PRIVATE ALL
m.filname = ALLTRIM(UPPER(m.filname))
IF '\' $ m.filname
	m.filname = SUBSTR(m.filname,1,RAT('\',m.filname))
	IF RIGHT(m.filname,1) = '\' AND LEN(m.filname) > 1 ;
			AND SUBSTR(m.filname,LEN(m.filname)-1,1) <> ':'
		m.filname = SUBSTR(m.filname,1,LEN(m.filname)-1)
	ENDIF
	RETURN m.filname
ELSE
	RETURN ''
ENDIF
*!******************************************************************************
*!
*! 过程 ADDBS
*!
*!******************************************************************************
FUNCTION addbs
PARAMETER pathname
PRIVATE ALL
pathname = ALLTRIM(UPPER(m.pathname))
IF !(RIGHT(m.pathname,1) $ '\:') AND !EMPTY(m.pathname)
	pathname = m.pathname + '\'
ENDIF
RETURN m.pathname
*!******************************************************************************
*!
*! 过程 VFPXTAB
*!
*!******************************************************************************
PROCEDURE vfpxtab
LPARAMETER p1,p2,p3,p4,p5,p6,p7,p8,p9,p10
IF PARAMETERS() < 3
	p3 = .T.
ENDIF
IF PARAMETERS() < 4
	p4 = .T.
ENDIF
oNewXtab=CREATE("genxtab",m.p1,m.p2,m.p3,m.p4,m.p5,m.p6,m.p7,m.p8,m.p9,m.p10)
IF TYPE("oNewXtab")="O"
	oNewXtab.makextab()
ENDIF
*!******************************************************************************
*!
*! 过程 ACTTHERM
*!
*!******************************************************************************
PROCEDURE acttherm
PARAMETERS lcPrompt
PUBLIC poThermRef
poThermRef = CREATEOBJECT("thermometer",m.lcPrompt)
poThermRef.SHOW()
*!******************************************************************************
*!
*! 过程 UPDTHERM
*!
*!******************************************************************************
PROCEDURE updtherm
LPARAMETER lcPercent,lcTask
IF m.lcPercent = 100
	poThermRef.COMPLETE()
ELSE
	poThermRef.UPDATE(m.lcPercent,lcTask)
ENDIF
*!******************************************************************************
*!
*! 过程 DEACTTHERMO
*!
*!******************************************************************************
PROCEDURE deactthermo
IF TYPE("oThermRef") = "O"
	poThermRef.RELEASE()
ENDIF
RELEASE poThermRef
