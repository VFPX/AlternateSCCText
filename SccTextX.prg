****************************************************************************
*	File:		SccTextX.PRG
****************************************************************************
* Enhanced version for VFPX, works with VFP9.x only
****************************************************************************
*
*	Copyright:	(c) 1995, Microsoft Corporation.
*				All Rights Reserved.
*
*	Contents:	Routines for creating text representations of .SCX, .VCX,
*				.MNX, .FRX, and .LBX files for the purpose of supporting
*				merge capabilities in source control systems.
*
*   Authors:	Sherri Kennamer
*				Mark Wilden (Mark@mWilden.com)
*				Markus Winhard (mw@bingo-ev.de)
*				Frank Camp (frank@prosysplus.nl)
*				Beth Massi (bmassi@goamerica.net)
*				Jürgen Wondzinski (wOOdy@ProLib.de)
*
*	Parameters:	cTableName	C	Fully-qualified name of the SCX/VCX/MNX/FRX/LBX
*				cType		C	Code indicating the file type
*								(See PRJTYPE_ constants, defined below)
*				cTextName	C	Fully-qualified name of the text file
*				lGenText	L	.T. Create a text file from the table
*								.F. Create a table from the text file
*
*	Returns:	0		File or table was successfully generated
*				-1		An error occurred
*
*	History:	17-Aug-95	sherrike	written
*				20-Nov-95	sherrike	use smart defaults for single filename
*				02-Dec-95	sherrike	return values for merge support
*               18-Jan-01   M.Winhard	implemented modifications of Mark Wilden as published
*										in FoxPro Advisor Oct.97, further improved it by:
*										+ finally fixed sort order for method names in class
*										+ dramatically improved performance with large VCX's
*										  (see SccTextEngine::SortMethods())
*										+ fixed sort order for classes in VCX
*										+ changed all PARAMETERS() to PCOUNT()
*										+ made sure SET COMPATIBLE is OFF as needed by FSIZE()
*										  to work the intended way
*				01-Feb-01	M.Winhard	auto-localizing for german users
*										(other localized VFP versions get english strings,
*										feel free to extend ;-) )
*				13-Aug-01	F.Camp		so I did. :-)
*										Fixed/Extend it to sort all the methods, objects
*										and properties.
*										Fixed second parameter in FGETS() calls.
*										In VFP7 it uses the Case-Insensitive Sort Order
*				15-Aug-01	M.Winhard	sort SCX like VCX
*				09-Aug-02	M.Winhard	allow creation of SCA files for FP2x screens
*				29-Oct-02	M.Winhard	VFP 7 and above: case insensitive sorting of methods
*										and properties is now optional; thus by default
*										VFP's standard methods and properties will be first,
*										not mixed with custom methods and properties
*				30-Oct-02	BethM		Fixed a Problem in Routine SortMethods()
*				31-Oct-02	wOOdy		changed all this.JUST*, this.ADDBS, this.FORCEEXT
*										functions to the VFP internal routines; revised
*										Thermometer code to use two container instead of nine
*										shapes; Reworked the HexStr2BinStr function for faster
*										and cleaner code. Overall Code cleanups (removed
*										unnecessary "m.", sorted and indented code,  etc),
*										fixed SCC_DBC_MEMO define error.
*				01-Nov-02	wOOdy		Spotted the cause for the "BugBug" workarounds in
*										proc "CreateTable"; Killed bug, cleaned code and got
*										rid of proc "GetReportStructure"
*				12-Nov-02	M.Winhard	Changed #defines for DBC file extensions to DC*;
*										added #defines for DBF file extensions as DB*
*				31-Nov-02	F.Camp		Fixed the sort order of contained objects in VCX and SCX
*										files so that they are created straight after the parent.
*										Header files will be created first, then the other controls
*										in a column of a grid
*				01-May-08	wOOdy		Changed the name from PL_SCCTEXT to SccTextX and
*										added it to codeplex.com/vfpx
*
*************************************************************************


#include "foxpro.h"

#DEFINE C_DEBUG .F.

*--------------------------------------------
* Set this constant to .T. if you want to have
* methods and properties sorted case insensitive.
* This means custom methods and properties will
* mixed with VFP's core methods and properties.
#DEFINE SORT_CASE_INSENSITIVE .F.

* If merge support is 1 and C_WRITECHECKSUMS is .T., write a checksum (sys(2007)) instead of
* converting binary to ascii. This drastically improves performance because OLE controls can
* be large and time-consuming to convert.
#DEFINE C_WRITECHECKSUMS .T.

#DEFINE SCCTEXTVER_LOC "SccTextX Version 1.0.0.1"	
#DEFINE ALERTTITLE_LOC "Microsoft Visual FoxPro"


* localizing for french, german, english strings for the rest 
#IF VERSION(3) == "33"	&& French
	#DEFINE ERRORMESSAGE_LOC "Erreur n°" + TRANSFORM(m.nError) + " dans " + m.cMethod + " (" + TRANSFORM(m.nLine) + "): " + m.cMessage
	#DEFINE ERRORTITLE_LOC "Erreur"
	#DEFINE ERR_ALERTCONTINUE_LOC "Continuer ?"
	#DEFINE ERR_BADVERSION_LOC "Mauvaise version de SCCTEXT."
	#DEFINE ERR_BIN2TEXTNOTSUPPORTED_LOC "La génération de fichier Texte n'est pas supporté pour le type de fichier '&cType'."
	#DEFINE ERR_FCREATE_LOC "FCREATE() erreur : "
	#DEFINE ERR_FIELDLISTTOOLONG_LOC "Liste de champs trop longue."
	#DEFINE ERR_FILENOTFOUND_LOC "Fichier non trouvé : "
	#DEFINE ERR_FOPEN_LOC "FOPEN() Erreur: "
	#DEFINE ERR_FOXERROR_11_LOC "Une valeur, un type ou le nombre des paramètres de l'appel de cette fonction est incorrect."
	#DEFINE ERR_INVALIDREVERSE_LOC "Paramètre REVERSE non valide."
	#DEFINE ERR_INVALIDTEXTNAME_LOC "Paramètre TEXTNAME non valide."
	#DEFINE ERR_LINENOACTION_LOC "Pas d'action effectuée sur cette ligne : "
	#DEFINE ERR_MAXBINLEN_LOC "MAXBINLEN doit être un multiple de 8. Programme arreté."
	#DEFINE ERR_NOTABLE_LOC "Un nom de table est requit."
	#DEFINE ERR_NOTEXTFILE_LOC "Le nom du fichier texte est requit pour créer une table."
	#DEFINE ERR_OVERWRITEREADONLY_LOC "Le fichier &cParameter1 est en lecture seule. L'écraser ?"
	#DEFINE ERR_TEXT2BINNOTSUPPORTED_LOC "La génération du fichier binaire n'est pas supporté pour le type '&cType' files."
	#DEFINE ERR_UNSUPPORTEDFIELDTYPE_LOC "Type de champ non supporté : "
	#DEFINE ERR_UNSUPPORTEDFILETYPE_LOC "Type de fichier non supporté : "
	* Used by the thermometer
	#DEFINE C_BINARYCONVERSION_LOC	"Conversion des données binaires : &cBinaryProgress.%"
	#DEFINE C_THERMCOMPLETE_LOC		"Génération de &cThermLabel terminée !"
	#DEFINE C_THERMLABEL_LOC		"Génération de &cThermLabel"
#ELIF VERSION(3) == "49"	&& german
	#DEFINE ERRORMESSAGE_LOC "Fehler Nr. " + TRANSFORM(m.nError) + " in " + m.cMethod +	" (" + TRANSFORM(m.nLine) + "): " + m.cMessage
	#DEFINE ERRORTITLE_LOC "Programmfehler"
	#DEFINE ERR_ALERTCONTINUE_LOC "Weiter?"
	#DEFINE ERR_BADVERSION_LOC "Falsche SCCTEXT-Version."
	#DEFINE ERR_BIN2TEXTNOTSUPPORTED_LOC "Das Generieren von Textdateien wird für Dateien vom Typ '&cType' nicht unterstützt."
	#DEFINE ERR_FCREATE_LOC "FCREATE()-Fehler: "
	#DEFINE ERR_FIELDLISTTOOLONG_LOC "Feldliste zu lang."
	#DEFINE ERR_FILENOTFOUND_LOC "Datei nicht gefunden: "
	#DEFINE ERR_FOPEN_LOC "FOPEN()-Fehler: "
	#DEFINE ERR_FOXERROR_11_LOC "Funktionsargumentwert, -typ oder -anzahl ungültig."
	#DEFINE ERR_INVALIDREVERSE_LOC "Ungültiger REVERSE-Parameter."
	#DEFINE ERR_INVALIDTEXTNAME_LOC "Ungültiger TEXTNAME-Parameter."
	#DEFINE ERR_LINENOACTION_LOC "Keine Maßnahme erfolgte in Zeile: "
	#DEFINE ERR_MAXBINLEN_LOC "Der Wert für MAXBINLEN muß ein Vielfaches von 8  sein. Programm abgebrochen."
	#DEFINE ERR_NOTABLE_LOC "Ein Tabellenname muß angegeben werden."
	#DEFINE ERR_NOTEXTFILE_LOC "Zum Erstellen einer Tabelle ist ein Textdateiname erforderlich."
	#DEFINE ERR_OVERWRITEREADONLY_LOC "Die Datei &cParameter1 ist schreibgeschützt. Überschreiben?"
	#DEFINE ERR_TEXT2BINNOTSUPPORTED_LOC "Das Generieren von Binärdateien wird für Dateien vom Typ '&cType' nicht unterstützt."
	#DEFINE ERR_UNSUPPORTEDFIELDTYPE_LOC "Feldtyp nicht unterstützt: "
	#DEFINE ERR_UNSUPPORTEDFILETYPE_LOC "Dateityp nicht unterstützt: "
	* Used by the thermometer
	#DEFINE C_BINARYCONVERSION_LOC	"Binärdaten werden konvertiert: &cBinaryProgress.%"
	#DEFINE C_THERMCOMPLETE_LOC		"&cThermLabel generiert."
	#DEFINE C_THERMLABEL_LOC		"&cThermLabel wird generiert..."
#ELSE
	#DEFINE ERRORMESSAGE_LOC "Error #" + TRANSFORM(m.nError) + " in " + m.cMethod + " (" + TRANSFORM(m.nLine) + "): " + m.cMessage
	#DEFINE ERRORTITLE_LOC "Program Error"
	#DEFINE ERR_ALERTCONTINUE_LOC "Continue?"
	#DEFINE ERR_BADVERSION_LOC "Bad SCCTEXT version."
	#DEFINE ERR_BIN2TEXTNOTSUPPORTED_LOC "Text file generation not supported for type '&cType' files."
	#DEFINE ERR_FCREATE_LOC "FCREATE() error: "
	#DEFINE ERR_FIELDLISTTOOLONG_LOC "Field list is too long."
	#DEFINE ERR_FILENOTFOUND_LOC "File not found: "
	#DEFINE ERR_FOPEN_LOC "FOPEN() error: "
	#DEFINE ERR_FOXERROR_11_LOC "Function argument value, type, or count is invalid."
	#DEFINE ERR_INVALIDREVERSE_LOC "Invalid REVERSE parameter."
	#DEFINE ERR_INVALIDTEXTNAME_LOC "Invalid TEXTNAME parameter."
	#DEFINE ERR_LINENOACTION_LOC "No action was taken on line: "
	#DEFINE ERR_MAXBINLEN_LOC "MAXBINLEN value must be a multiple of 8. Program aborted."
	#DEFINE ERR_NOTABLE_LOC "A table name is required."
	#DEFINE ERR_NOTEXTFILE_LOC "Text file name is required to create a table."
	#DEFINE ERR_OVERWRITEREADONLY_LOC "File &cParameter1 is read-only. Overwrite it?"
	#DEFINE ERR_TEXT2BINNOTSUPPORTED_LOC "Binary file generation not supported for type '&cType' files."
	#DEFINE ERR_UNSUPPORTEDFIELDTYPE_LOC "Field type not supported: "
	#DEFINE ERR_UNSUPPORTEDFILETYPE_LOC "File type not supported: "
	* Used by the thermometer
	#DEFINE C_BINARYCONVERSION_LOC	"Converting binary data: &cBinaryProgress.%"
	#DEFINE C_THERMCOMPLETE_LOC		"Generate &cThermLabel complete!"
	#DEFINE C_THERMLABEL_LOC		"Generating &cThermLabel"
#ENDIF

#DEFINE CRLF		CHR(13) + CHR(10)
#DEFINE MAXBINLEN	96		&& this value must be a multiple of 8!!!

#DEFINE FILE_ATTRIBUTE_NORMAL	128

* Text file support for each file type
*	0 indicates no text file support
*	1 indicates one-way support (to text)
*	2 indicates two-way support (for merging)
#DEFINE SCC_DBC_SUPPORT		0
#DEFINE SCC_DBF_SUPPORT		0
#DEFINE SCC_FORM_SUPPORT	1
#DEFINE SCC_LABEL_SUPPORT	1
#DEFINE SCC_MENU_SUPPORT	1
#DEFINE SCC_REPORT_SUPPORT	1
#DEFINE SCC_VCX_SUPPORT		1

* These are the extensions used for the text file
#DEFINE SCC_ASCII_DBC_EXT		"DCA"
#DEFINE SCC_ASCII_DBF_EXT		"DBA"
#DEFINE SCC_ASCII_FORM_EXT		"SCA"
#DEFINE SCC_ASCII_LABEL_EXT		"LBA"
#DEFINE SCC_ASCII_MENU_EXT		"MNA"
#DEFINE SCC_ASCII_REPORT_EXT	"FRA"
#DEFINE SCC_ASCII_VCX_EXT		"VCA"

* These are the extensions used for the binary file
#DEFINE SCC_DBC_EXT			"DBC"
#DEFINE SCC_DBF_EXT			"DBF"
#DEFINE SCC_FORM_EXT		"SCX"
#DEFINE SCC_LABEL_EXT		"LBX"
#DEFINE SCC_MENU_EXT		"MNX"
#DEFINE SCC_REPORT_EXT		"FRX"
#DEFINE SCC_VCX_EXT			"VCX"

* These are the extensions used for the binary file
#DEFINE SCC_DBC_MEMO		"DBT"
#DEFINE SCC_DBF_MEMO		"FPT"
#DEFINE SCC_FORM_MEMO		"SCT"
#DEFINE SCC_LABEL_MEMO		"LBT"
#DEFINE SCC_MENU_MEMO		"MNT"
#DEFINE SCC_REPORT_MEMO		"FRT"
#DEFINE SCC_VCX_MEMO		"VCT"

* These are the project type identifiers for the files
#DEFINE PRJTYPE_DBC			"d"
#DEFINE PRJTYPE_DBF			"D"
#DEFINE PRJTYPE_FORM		"K"
#DEFINE PRJTYPE_LABEL		"B"
#DEFINE PRJTYPE_MENU		"M"
#DEFINE PRJTYPE_REPORT		"R"
#DEFINE PRJTYPE_VCX			"V"

* These are the extensions used for table backups
#DEFINE SCC_DBC_INDEX_BAK		"DC3"
#DEFINE SCC_DBC_MEMO_BAK		"DC2"
#DEFINE SCC_DBC_TABLE_BAK		"DC1"
#DEFINE SCC_DBF_INDEX_BAK		"DB3"
#DEFINE SCC_DBF_MEMO_BAK		"DB2"
#DEFINE SCC_DBF_TABLE_BAK		"DB1"
#DEFINE SCC_FORM_MEMO_BAK		"SC2"
#DEFINE SCC_FORM_TABLE_BAK		"SC1"
#DEFINE SCC_LABEL_MEMO_BAK		"LB2"
#DEFINE SCC_LABEL_TABLE_BAK		"LB1"
#DEFINE SCC_MENU_MEMO_BAK		"MN2"
#DEFINE SCC_MENU_TABLE_BAK		"MN1"
#DEFINE SCC_REPORT_MEMO_BAK		"FR2"
#DEFINE SCC_REPORT_TABLE_BAK	"FR1"
#DEFINE SCC_VCX_MEMO_BAK		"VC2"
#DEFINE SCC_VCX_TABLE_BAK		"VC1"

* These are the extensions used for text file backups
#DEFINE SCC_DBC_TEXT_BAK		"DCB"
#DEFINE SCC_DBF_TEXT_BAK		"DBB"
#DEFINE SCC_FORM_TEXT_BAK		"SCB"
#DEFINE SCC_LABEL_TEXT_BAK		"LBB"
#DEFINE SCC_MENU_TEXT_BAK		"MNB"
#DEFINE SCC_REPORT_TEXT_BAK		"FRB"
#DEFINE SCC_VCX_TEXT_BAK		"VCB"

* These are used for building markers used to parse the text back into a table
#DEFINE MARKBINENDWORD		"[BINEND "
#DEFINE MARKBINENDWORD2		"]"
#DEFINE MARKBINSTARTWORD	"[BINSTART "
#DEFINE MARKBINSTARTWORD2	"]"
#DEFINE MARKCHECKSUM		"CHECKSUM="
#DEFINE MARKEOF				"[EOF]"
#DEFINE MARKFIELDEND		"] "
#DEFINE MARKFIELDSTART		"["
#DEFINE MARKMEMOENDWORD		"[END "
#DEFINE MARKMEMOENDWORD2	"]"
#DEFINE MARKMEMOSTARTWORD	"[START "
#DEFINE MARKMEMOSTARTWORD2	"]"
#DEFINE MARKRECORDEND		" RECORD]"
#DEFINE MARKRECORDSTART		"["

#DEFINE SKIPEMPTYFIELD		.T.

* These are used to override default behavior for specific fields
#DEFINE VCX_CHARASBIN_LIST		""
#DEFINE VCX_EXCLUDE_LIST		" OBJCODE TIMESTAMP "
#DEFINE VCX_MEMOASBIN_LIST		" OLE OLE2 "
#DEFINE VCX_MEMOASCHAR_LIST		" CLASS CLASSLOC BASECLASS OBJNAME PARENT "
#DEFINE VCX_MEMOVARIES_LIST		" RESERVED4 RESERVED5 "

#DEFINE FRX_CHARASBIN_LIST		""
#DEFINE FRX_EXCLUDE_LIST		" TIMESTAMP "
#DEFINE FRX_MEMOASBIN_LIST		" TAG TAG2 "
#DEFINE FRX_MEMOASCHAR_LIST		" NAME STYLE PICTURE ORDER FONTFACE "
#DEFINE FRX_MEMOVARIES_LIST		""

#DEFINE MNX_CHARASBIN_LIST		" MARK "
#DEFINE MNX_EXCLUDE_LIST		" TIMESTAMP "
#DEFINE MNX_MEMOASBIN_LIST		""
#DEFINE MNX_MEMOASCHAR_LIST		" NAME PROMPT COMMAND MESSAGE KEYNAME KEYLABEL "
#DEFINE MNX_MEMOVARIES_LIST		""

#DEFINE DBC_CHARASBIN_LIST		""
#DEFINE DBC_EXCLUDE_LIST		""
#DEFINE DBC_MEMOASBIN_LIST		""
#DEFINE DBC_MEMOASCHAR_LIST		""
#DEFINE DBC_MEMOVARIES_LIST		" PROPERTY CODE USER "

* Used by the thermometer
#DEFINE WIN32FONT				"MS Sans Serif"
#DEFINE WIN95FONT				"Arial"

*---------------------------------------------------------------------------
PARAMETERS cTableName, cType, cTextName, lGenText


LOCAL iParmCount, cCollate, cCompatible
iParmCount = PCOUNT()
cCollate = SET("COLLATE")
cCompatible = SET("COMPATIBLE")
* we need a constant sort order with ASORT()
SET COLLATE TO "MACHINE"
* FSIZE() needs SET COMPATIBLE OFF to work the way it's used in this program
SET COMPATIBLE OFF


LOCAL  obj, iResult
iResult = -1
IF m.iParmCount = 1 .AND. TYPE('m.cTableName') = 'C'
	* Check to see if we've been passed only a PRJTYPE value. If so, return a
	* value to indicate text support for the file type.
	*	0 indicates no text file support
	*	1 indicates one-way support (to text)
	*	2 indicates two-way support (for merging)
	*  -1 indicates m.cTableName is not a recognized file type
	iResult = TextSupport(m.cTableName)
ENDIF

IF m.iResult = -1 && .and. file(m.cTableName)
	obj = CREATEOBJ("SccTextEngine", m.cTableName, m.cType, m.cTextName, m.lGenText, m.iParmCount)
	IF TYPE("m.obj") = "O" .AND. .NOT. ISNULL(m.obj)
		obj.PROCESS()
		IF TYPE("m.obj") = "O" .AND. .NOT. ISNULL(m.obj)
			iResult = obj.iResult
		ENDIF
	ENDIF
	RELEASE obj
ENDIF


IF NOT m.cCollate == "MACHINE"
	* set back changed collate setting to user's choice
	SET COLLATE TO (m.cCollate)
ENDIF
IF NOT m.cCompatible == "OFF"
	* set back changed compatible setting to user's choice
	SET COMPATIBLE &cCompatible.
ENDIF


RETURN (m.iResult)

******************************************************************
PROCEDURE TextSupport
	PARAMETERS cFileType
	DO CASE
			* Check to see if we've been passed only a PRJTYPE value. If so, return a
			* value to indicate text support for the file type.
			*	0 indicates no text file support
			*	1 indicates one-way support (to text)
			*	2 indicates two-way support (for merging)
		CASE m.cFileType == PRJTYPE_FORM
			RETURN SCC_FORM_SUPPORT
		CASE m.cFileType == PRJTYPE_LABEL
			RETURN SCC_LABEL_SUPPORT
		CASE m.cFileType == PRJTYPE_MENU
			RETURN SCC_MENU_SUPPORT
		CASE m.cFileType == PRJTYPE_REPORT
			RETURN SCC_REPORT_SUPPORT
		CASE m.cFileType == PRJTYPE_VCX
			RETURN SCC_VCX_SUPPORT
		CASE m.cFileType == PRJTYPE_DBC
			RETURN SCC_DBC_SUPPORT
		OTHERWISE
			RETURN -1
	ENDCASE
ENDPROC

******************************************************************
DEFINE CLASS SccTextEngine AS CUSTOM
	cIndexBakName = ""
	cIndexName    = ""
	cMemoBakName  = ""
	cMemoName     = ""
	cMessage      = ""
	cTableBakName = ""
	cTableName    = ""
	cTextBakName  = ""
	cTextName     = ""
	cType         = ""
	cVCXCursor    = ""	&& If we're generating text for a .VCX, we create a temporary file with the classes sorted.
	HadError      = .F.
	iError        = 0
	iHandle       = -1
	iResult       = -1 && Fail
	lGenText      = .T.
	lMadeBackup   = .F.
	oThermRef     = ""
	SetErrorOff   = .F.
	DIMENSION aEnvironment[1]

	PROCEDURE INIT(cTableName, cType, cTextName, lGenText, iParmCount)
		LOCAL iAction

		IF m.iParmCount = 1 .AND. TYPE('m.cTableName') = 'C'
			* Interpret the single parameter as a filename and be smart about defaults
			IF THIS.ISBINARY(m.cTableName)
				cType     = THIS.GetPrjType(m.cTableName)
				cTextName = FORCEEXT(m.cTableName, THIS.GetAsciiExt(m.cType))
				lGenText  = .T.
			ELSE
				IF THIS.IsAscii(m.cTableName)
					cType      = THIS.GetPrjType(m.cTableName)
					cTextName  = m.cTableName
					cTableName = FORCEEXT(m.cTextName, THIS.GetBinaryExt(m.cType))
					lGenText   = .F.
				ENDIF
			ENDIF
		ENDIF

		THIS.cTableName = m.cTableName
		THIS.cTextName  = m.cTextName
		THIS.cType      = m.cType
		THIS.lGenText   = m.lGenText

		* Verify that we've got valid parameters
		IF TYPE('this.cTableName') <> 'C' .OR. TYPE('this.cType') <> 'C' ;
				.OR. TYPE('this.cTextName') <> 'C' .OR. TYPE('this.lGenText') <> 'L'
			THIS.Alert(ERR_FOXERROR_11_LOC)
			RETURN .F.
		ENDIF

		* Verify parameters before calling ForceExt
		THIS.cMemoName = FORCEEXT(THIS.cTableName, THIS.GetBinaryMemo(THIS.cType))

		* Verify that we support the requested action
		iAction = IIF(m.lGenText, 1, 2)
		DO CASE
			CASE m.cType == PRJTYPE_FORM .AND. SCC_FORM_SUPPORT < m.iAction
				iAction = m.iAction * -1
			CASE m.cType == PRJTYPE_LABEL .AND. SCC_LABEL_SUPPORT < m.iAction
				iAction = m.iAction * -1
			CASE m.cType == PRJTYPE_MENU .AND. SCC_MENU_SUPPORT < m.iAction
				iAction = m.iAction * -1
			CASE m.cType == PRJTYPE_REPORT .AND. SCC_REPORT_SUPPORT < m.iAction
				iAction = m.iAction * -1
			CASE m.cType == PRJTYPE_VCX .AND. SCC_VCX_SUPPORT < m.iAction
				iAction = m.iAction * -1
			CASE m.cType == PRJTYPE_DBC .AND. SCC_DBC_SUPPORT < m.iAction
				iAction = m.iAction * -1
		ENDCASE

		IF m.iAction = -1
			THIS.Alert(ERR_BIN2TEXTNOTSUPPORTED_LOC)
			RETURN .F.
		ENDIF
		IF m.iAction = -2
			THIS.Alert(ERR_TEXT2BINNOTSUPPORTED_LOC)
			RETURN .F.
		ENDIF

		IF .NOT. THIS.SETUP()
			RETURN .F.
		ENDIF

		IF (MAXBINLEN % 8 <> 0)
			THIS.Alert(ERR_MAXBINLEN_LOC)
			RETURN .F.
		ENDIF
	ENDPROC

	PROCEDURE ERASE
		PARAMETERS cFilename
		IF !EMPTY(m.cFilename) .AND. FILE(m.cFilename)
			= SetFileAttributes(m.cFilename, FILE_ATTRIBUTE_NORMAL)
			ERASE (m.cFilename)
		ENDIF
	ENDPROC

	PROCEDURE MakeBackup
		* Fill in the names of the backup files
		DO CASE
			CASE THIS.cType == PRJTYPE_FORM
				THIS.cTextBakName  = FORCEEXT(THIS.cTextName, SCC_FORM_TEXT_BAK)
				THIS.cTableBakName = FORCEEXT(THIS.cTableName, SCC_FORM_TABLE_BAK)
				THIS.cMemoBakName  = FORCEEXT(THIS.cMemoName, SCC_FORM_MEMO_BAK)
			CASE THIS.cType == PRJTYPE_REPORT
				THIS.cTextBakName  = FORCEEXT(THIS.cTextName, SCC_REPORT_TEXT_BAK)
				THIS.cTableBakName = FORCEEXT(THIS.cTableName, SCC_REPORT_TABLE_BAK)
				THIS.cMemoBakName  = FORCEEXT(THIS.cMemoName, SCC_REPORT_MEMO_BAK)
			CASE THIS.cType == PRJTYPE_VCX
				THIS.cTextBakName  = FORCEEXT(THIS.cTextName, SCC_VCX_TEXT_BAK)
				THIS.cTableBakName = FORCEEXT(THIS.cTableName, SCC_VCX_TABLE_BAK)
				THIS.cMemoBakName  = FORCEEXT(THIS.cMemoName, SCC_VCX_MEMO_BAK)
			CASE THIS.cType == PRJTYPE_MENU
				THIS.cTextBakName  = FORCEEXT(THIS.cTextName, SCC_MENU_TEXT_BAK)
				THIS.cTableBakName = FORCEEXT(THIS.cTableName, SCC_MENU_TABLE_BAK)
				THIS.cMemoBakName  = FORCEEXT(THIS.cMemoName, SCC_MENU_MEMO_BAK)
			CASE THIS.cType == PRJTYPE_LABEL
				THIS.cTextBakName  = FORCEEXT(THIS.cTextName, SCC_LABEL_TEXT_BAK)
				THIS.cTableBakName = FORCEEXT(THIS.cTableName, SCC_LABEL_TABLE_BAK)
				THIS.cMemoBakName  = FORCEEXT(THIS.cMemoName, SCC_LABEL_MEMO_BAK)
			CASE THIS.cType == PRJTYPE_DBC
				THIS.cTextBakName  = FORCEEXT(THIS.cTextName, SCC_DBC_TEXT_BAK)
				THIS.cTableBakName = FORCEEXT(THIS.cTableName, SCC_DBC_TABLE_BAK)
				THIS.cMemoBakName  = FORCEEXT(THIS.cMemoName, SCC_DBC_MEMO_BAK)
				THIS.cIndexBakName = FORCEEXT(THIS.cIndexName, SCC_DBC_INDEX_BAK)
		ENDCASE

		* Delete any existing backup
		THIS.DeleteBackup()

		* Create new backup files
		IF THIS.lGenText
			IF FILE(THIS.cTextName)
				COPY FILE (THIS.cTextName) TO (THIS.cTextBakName)
			ENDIF
		ELSE
			IF FILE(THIS.cTableName) .AND. FILE(THIS.cMemoName)
				COPY FILE (THIS.cTableName) TO (THIS.cTableBakName)
				COPY FILE (THIS.cMemoName) TO (THIS.cMemoBakName)
				IF !EMPTY(THIS.cIndexName) .AND. FILE(THIS.cIndexName)
					COPY FILE (THIS.cIndexName) TO (THIS.cIndexBakName)
				ENDIF
			ENDIF
		ENDIF

		THIS.lMadeBackup = .T.
	ENDPROC

	PROCEDURE RestoreBackup
		IF THIS.lGenText
			THIS.ERASE(THIS.cTextName)
		ELSE
			THIS.ERASE(THIS.cTableName)
			THIS.ERASE(THIS.cMemoName)
			IF .NOT. EMPTY(THIS.cIndexName)
				THIS.ERASE(THIS.cIndexName)
			ENDIF
		ENDIF

		IF THIS.lGenText
			IF FILE(THIS.cTextBakName)
				COPY FILE (THIS.cTextBakName) TO (THIS.cTextName)
			ENDIF
		ELSE
			IF FILE(THIS.cTableBakName) .AND. FILE(THIS.cMemoBakName)
				COPY FILE (THIS.cTableBakName) TO (THIS.cTableName)
				COPY FILE (THIS.cMemoBakName) TO (THIS.cMemoName)
				IF !EMPTY(THIS.cIndexBakName) .AND. FILE(THIS.cIndexBakName)
					COPY FILE (THIS.cIndexBakName) TO (THIS.cIndexName)
				ENDIF
			ENDIF
		ENDIF
	ENDPROC

	PROCEDURE DeleteBackup
		IF THIS.lGenText
			THIS.ERASE(THIS.cTextBakName)
		ELSE
			THIS.ERASE(THIS.cTableBakName)
			THIS.ERASE(THIS.cMemoBakName)
			IF !EMPTY(THIS.cIndexBakName)
				THIS.ERASE(THIS.cIndexBakName)
			ENDIF
		ENDIF
	ENDPROC

	PROCEDURE GetAsciiExt
		PARAMETERS cType
		DO CASE
			CASE m.cType == PRJTYPE_FORM
				RETURN SCC_ASCII_FORM_EXT
			CASE m.cType == PRJTYPE_REPORT
				RETURN SCC_ASCII_REPORT_EXT
			CASE m.cType == PRJTYPE_VCX
				RETURN SCC_ASCII_VCX_EXT
			CASE m.cType == PRJTYPE_MENU
				RETURN SCC_ASCII_MENU_EXT
			CASE m.cType == PRJTYPE_LABEL
				RETURN SCC_ASCII_LABEL_EXT
			CASE m.cType == PRJTYPE_DBC
				RETURN SCC_ASCII_DBC_EXT
		ENDCASE
	ENDPROC

	PROCEDURE GetBinaryExt
		PARAMETERS cType
		DO CASE
			CASE m.cType == PRJTYPE_FORM
				RETURN SCC_FORM_EXT
			CASE m.cType == PRJTYPE_REPORT
				RETURN SCC_REPORT_EXT
			CASE m.cType == PRJTYPE_VCX
				RETURN SCC_VCX_EXT
			CASE m.cType == PRJTYPE_MENU
				RETURN SCC_MENU_EXT
			CASE m.cType == PRJTYPE_LABEL
				RETURN SCC_LABEL_EXT
			CASE m.cType == PRJTYPE_DBC
				RETURN SCC_DBC_EXT
		ENDCASE
	ENDPROC

	PROCEDURE GetBinaryMemo
		PARAMETERS cType
		DO CASE
			CASE m.cType == PRJTYPE_FORM
				RETURN SCC_FORM_MEMO
			CASE m.cType == PRJTYPE_REPORT
				RETURN SCC_REPORT_MEMO
			CASE m.cType == PRJTYPE_VCX
				RETURN SCC_VCX_MEMO
			CASE m.cType == PRJTYPE_MENU
				RETURN SCC_MENU_MEMO
			CASE m.cType == PRJTYPE_LABEL
				RETURN SCC_LABEL_MEMO
			CASE m.cType == PRJTYPE_DBC
				RETURN SCC_DBC_MEMO
		ENDCASE
	ENDPROC

	PROCEDURE GetPrjType
		PARAMETERS cFilename
		LOCAL m.cExt
		m.cExt = UPPER(JUSTEXT(m.cFilename))
		DO CASE
			CASE INLIST(m.cExt, SCC_ASCII_FORM_EXT, SCC_FORM_EXT)
				RETURN PRJTYPE_FORM
			CASE INLIST(m.cExt, SCC_ASCII_REPORT_EXT, SCC_REPORT_EXT)
				RETURN PRJTYPE_REPORT
			CASE INLIST(m.cExt, SCC_ASCII_VCX_EXT, SCC_VCX_EXT)
				RETURN PRJTYPE_VCX
			CASE INLIST(m.cExt, SCC_ASCII_MENU_EXT, SCC_MENU_EXT)
				RETURN PRJTYPE_MENU
			CASE INLIST(m.cExt, SCC_ASCII_LABEL_EXT, SCC_LABEL_EXT)
				RETURN PRJTYPE_LABEL
			CASE INLIST(m.cExt, SCC_ASCII_DBC_EXT, SCC_DBC_EXT)
				RETURN PRJTYPE_DBC
			OTHERWISE
				RETURN ''
		ENDCASE
	ENDPROC

	PROCEDURE IsAscii
		PARAMETERS cFilename
		LOCAL m.cExt
		m.cExt = UPPER(JUSTEXT(m.cFilename))
		RETURN INLIST(m.cExt, SCC_ASCII_FORM_EXT, SCC_ASCII_REPORT_EXT, SCC_ASCII_VCX_EXT, SCC_ASCII_MENU_EXT, SCC_ASCII_LABEL_EXT, SCC_ASCII_DBC_EXT)
	ENDPROC

	PROCEDURE ISBINARY
		PARAMETERS cFilename
		LOCAL m.cExt
		m.cExt = UPPER(JUSTEXT(m.cFilename))
		RETURN INLIST(m.cExt, SCC_FORM_EXT, SCC_REPORT_EXT, SCC_VCX_EXT, SCC_MENU_EXT, SCC_LABEL_EXT, SCC_DBC_EXT)
	ENDPROC

	PROCEDURE SETUP

		DIMENSION THIS.aEnvironment[5]

		THIS.aEnvironment[1] = SET("deleted")
		THIS.aEnvironment[2] = SELECT()
		THIS.aEnvironment[3] = SET("safety")
		THIS.aEnvironment[4] = SET("talk")
		THIS.aEnvironment[5] = SET("asserts")

		SET TALK OFF

		DECLARE INTEGER SetFileAttributes IN win32api STRING lpFileName, INTEGER dwFileAttributes
		DECLARE INTEGER sprintf IN msvcrt40.DLL ;
			STRING @lpBuffer, STRING lpFormat, INTEGER iChar1, INTEGER iChar2, ;
			INTEGER iChar3, INTEGER iChar4, INTEGER iChar5, INTEGER iChar6, ;
			INTEGER iChar7, INTEGER iChar8

		SET SAFETY OFF
		SET DELETED OFF
		SELECT 0
		IF C_DEBUG
			SET ASSERTS ON
		ENDIF

	ENDPROC

	PROCEDURE Cleanup
		LOCAL ARRAY aEnvironment[alen(this.aEnvironment)]
		ACOPY(THIS.aEnvironment, aEnvironment)
		SET DELETED &aEnvironment[1]
		SET SAFETY &aEnvironment[3]
		USE
		SELECT (aEnvironment[2])
		IF THIS.iHandle <> -1
			FCLOSE(THIS.iHandle)
			THIS.iHandle = -1
		ENDIF
		SET TALK &aEnvironment[4]
		IF USED(THIS.cVCXCursor)
			USE IN (THIS.cVCXCursor)
			THIS.cVCXCursor = ""
		ENDIF
		SET ASSERTS &aEnvironment[5]
	ENDPROC

	PROCEDURE DESTROY
		IF TYPE("this.oThermRef") = "O"
			THIS.oThermRef.RELEASE()
		ENDIF

		THIS.Cleanup()

		IF THIS.lMadeBackup
			IF THIS.iResult <> 0
				THIS.RestoreBackup()
			ENDIF
			THIS.DeleteBackup()
		ENDIF
	ENDPROC

	PROCEDURE ERROR
		PARAMETERS nError, cMethod, nLine, oObject, cMessage

		LOCAL cAction

		THIS.HadError = .T.
		THIS.iError   = m.nError
		THIS.cMessage = IIF(EMPTY(m.cMessage), MESSAGE(), m.cMessage)

		IF THIS.SetErrorOff
			RETURN
		ENDIF

		cMessage = IIF(EMPTY(m.cMessage), MESSAGE(), m.cMessage)
		IF TYPE("m.oObject") = "O" .AND. .NOT. ISNULL(m.oObject) .AND. AT(".", m.cMethod) = 0
			cMethod = m.oObject.NAME + "." + m.cMethod
		ENDIF

		IF C_DEBUG
			cAction = THIS.Alert(ERRORMESSAGE_LOC, MB_ICONEXCLAMATION + MB_ABORTRETRYIGNORE, ERRORTITLE_LOC)
			DO CASE
				CASE m.cAction="RETRY"
					THIS.HadError = .F.
					CLEAR TYPEAHEAD
					SET STEP ON
					&cAction
				CASE m.cAction="IGNORE"
					THIS.HadError = .F.
					RETURN
			ENDCASE
		ELSE
			IF m.nError = 1098  && User-defined error
				cAction = THIS.Alert(MESSAGE(), MB_ICONEXCLAMATION + MB_OK, ERRORTITLE_LOC)
			ELSE
				cAction = THIS.Alert(ERRORMESSAGE_LOC, MB_ICONEXCLAMATION + MB_OK, ERRORTITLE_LOC)
			ENDIF
		ENDIF
		THIS.CANCEL()

	ENDPROC

	PROCEDURE CANCEL
		PARAMETERS cMessage
		IF !EMPTY(m.cMessage)
			cAction = THIS.Alert(m.cMessage)
		ENDIF
		RETURN TO PROCESS -1
	ENDPROC

	PROCEDURE Alert
		PARAMETERS cMessage, cOptions, cTitle, cParameter1, cParameter2

		PRIVATE cOptions, cResponse
		cOptions = IIF(EMPTY(m.cOptions), 0, m.cOptions)

		IF PCOUNT() > 3 && a parameter was passed
			m.cMessage = [&cMessage]
		ENDIF

		CLEAR TYPEAHEAD
		IF !EMPTY(m.cTitle)
			cResponse = MESSAGEBOX(m.cMessage, m.cOptions, m.cTitle)
		ELSE
			cResponse = MESSAGEBOX(m.cMessage, m.cOptions, ALERTTITLE_LOC)
		ENDIF

		DO CASE
				* The strings below are used internally and should not
				* be localized
			CASE m.cResponse = 1
				cResponse = "OK"
			CASE m.cResponse = 6
				cResponse = "YES"
			CASE m.cResponse = 7
				cResponse = "NO"
			CASE m.cResponse = 2
				cResponse = "CANCEL"
			CASE m.cResponse = 3
				cResponse = "ABORT"
			CASE m.cResponse = 4
				cResponse = "RETRY"
			CASE m.cResponse = 5
				cResponse = "IGNORE"
		ENDCASE
		RETURN m.cResponse

	ENDPROC

	PROCEDURE PROCESS
		LOCAL cThermLabel

		IF THIS.FilesAreWritable()
			* Backup the file(s)
			THIS.MakeBackup()

			* Create and show the thermometer
			m.cThermLabel = IIF(THIS.lGenText, THIS.cTextName, THIS.cTableName)
			THIS.oThermRef = CREATEOBJECT("thermometer", C_THERMLABEL_LOC)
			THIS.oThermRef.SHOW()

			IF THIS.lGenText
				THIS.iResult = THIS.WriteTextFile()
			ELSE
				THIS.iResult = THIS.WriteTableFile()
			ENDIF

			IF THIS.iResult = 0
				THIS.oThermRef.COMPLETE(C_THERMCOMPLETE_LOC)
			ENDIF
		ENDIF
	ENDPROC

	PROCEDURE FilesAreWritable
		PRIVATE aText
		IF THIS.lGenText
			* Verify we can write the text file
			IF (ADIR(aText, THIS.cTextName) = 1 .AND. 'R' $ aText[1, 5])
				IF THIS.Alert(ERR_OVERWRITEREADONLY_LOC, MB_YESNO, '', THIS.cTextName) = "NO"
					RETURN .F.
				ENDIF
			ENDIF
			SetFileAttributes(THIS.cTextName, FILE_ATTRIBUTE_NORMAL)
		ELSE
			* Verify we can write the table
			IF (ADIR(aText, THIS.cTableName) = 1 .AND. 'R' $ aText[1, 5])
				IF THIS.Alert(ERR_OVERWRITEREADONLY_LOC, MB_YESNO, '', THIS.cTableName) = "NO"
					RETURN .F.
				ENDIF
			ELSE
				IF (ADIR(aText, THIS.cMemoName) = 1 .AND. 'R' $ aText[1, 5])
					IF THIS.Alert(ERR_OVERWRITEREADONLY_LOC, MB_YESNO, '', THIS.cMemoName) = "NO"
						RETURN .F.
					ENDIF
				ENDIF
			ENDIF
			SetFileAttributes(THIS.cTableName, FILE_ATTRIBUTE_NORMAL)
			SetFileAttributes(THIS.cMemoName, FILE_ATTRIBUTE_NORMAL)
		ENDIF
		RETURN .T.
	ENDPROC

	PROCEDURE WriteTableFile
		THIS.iHandle = FOPEN(THIS.cTextName)
		IF THIS.iHandle = -1
			THIS.Alert(ERR_FOPEN_LOC + THIS.cTextName)
			RETURN -1
		ENDIF

		THIS.oThermRef.iBasis = FSEEK(THIS.iHandle, 0, 2)
		FSEEK(THIS.iHandle, 0, 0)

		THIS.ValidVersion(FGETS(THIS.iHandle, 8192))
		THIS.CreateTable(FGETS(THIS.iHandle, 8192), VAL(FGETS(THIS.iHandle, 8192)))
		DO CASE
			CASE INLIST(THIS.cType, PRJTYPE_FORM, PRJTYPE_VCX, PRJTYPE_MENU, ;
					PRJTYPE_REPORT, PRJTYPE_LABEL)
				THIS.WriteTable
			OTHERWISE
				THIS.CANCEL(ERR_UNSUPPORTEDFILETYPE_LOC + THIS.cType)
		ENDCASE

		FCLOSE(THIS.iHandle)
		THIS.iHandle = -1
		IF INLIST(THIS.cType, PRJTYPE_FORM, PRJTYPE_VCX)
			IF THIS.cType = PRJTYPE_VCX
				* Additional work may need to be performed on a VCX
				THIS.FixUpVCX
			ENDIF

			USE
			COMPILE FORM (THIS.cTableName)
		ENDIF
		USE
		RETURN 0 && Success
	ENDPROC

	PROCEDURE FixUpVCX
		PRIVATE aClassList, i
		SELECT objname, RECNO() FROM DBF() WHERE NOT DELETED() AND reserved1 == 'Class' ;
			INTO ARRAY aClassList
		IF TYPE('aClassList[1]') <> 'U'
			* If objects were added to or removed from a class during merge,
			* the record count will be out of sync.
			FOR m.i = 1 TO ALEN(aClassList, 1)
				GO (aClassList[m.i, 2])
				IF m.i = ALEN(aClassList, 1)
					REPLACE reserved2 WITH ALLTRIM(STR(RECCOUNT() - aClassList[m.i, 2]))
				ELSE
					REPLACE reserved2 WITH ALLTRIM(STR(aClassList[m.i + 1, 2] - aClassList[m.i, 2] - 1))
				ENDIF
			ENDFOR
		ENDIF
	ENDPROC

	PROCEDURE CreateTable
		PARAMETERS cFieldlist, iCodePage
		PRIVATE c1, c2, c3, c4, c5, c6, aStruct

		DO CASE
			CASE LEN(m.cFieldlist) < 251
				CREATE TABLE (THIS.cTableName) FREE (&cFieldlist)
			CASE LEN(m.cFieldlist) < 501
				m.c1 = SUBSTR(m.cFieldlist, 1, 250)
				m.c2 = SUBSTR(m.cFieldlist, 251)
				CREATE TABLE (THIS.cTableName) FREE (&c1.&c2)
			CASE LEN(m.cFieldlist) < 751
				m.c1 = SUBSTR(m.cFieldlist, 1, 250)
				m.c2 = SUBSTR(m.cFieldlist, 251, 250)
				m.c3 = SUBSTR(m.cFieldlist, 501)
				CREATE TABLE (THIS.cTableName) FREE (&c1.&c2.&c3)
			CASE LEN(m.cFieldlist) < 1001
				m.c1 = SUBSTR(m.cFieldlist, 1, 250)
				m.c2 = SUBSTR(m.cFieldlist, 251, 250)
				m.c3 = SUBSTR(m.cFieldlist, 501, 250)
				m.c4 = SUBSTR(m.cFieldlist, 751)
				CREATE TABLE (THIS.cTableName) FREE (&c1.&c2.&c3.&c4)
			CASE LEN(m.cFieldlist) < 1251
				m.c1 = SUBSTR(m.cFieldlist, 1, 250)
				m.c2 = SUBSTR(m.cFieldlist, 251, 250)
				m.c3 = SUBSTR(m.cFieldlist, 501, 250)
				m.c4 = SUBSTR(m.cFieldlist, 751, 250)
				m.c5 = SUBSTR(m.cFieldlist, 1001)
				CREATE TABLE (THIS.cTableName) FREE (&c1.&c2.&c3.&c4.&c5)
			CASE LEN(m.cFieldlist) < 1501
				m.c1 = SUBSTR(m.cFieldlist, 1, 250)
				m.c2 = SUBSTR(m.cFieldlist, 251, 250)
				m.c3 = SUBSTR(m.cFieldlist, 501, 250)
				m.c4 = SUBSTR(m.cFieldlist, 751, 250)
				m.c5 = SUBSTR(m.cFieldlist, 1001, 250)
				m.c6 = SUBSTR(m.cFieldlist, 1251)
				CREATE TABLE (THIS.cTableName) FREE (&c1.&c2.&c3.&c4.&c5.&c6)
			OTHERWISE
				* Not supported
				THIS.CANCEL(ERR_FIELDLISTTOOLONG_LOC)
		ENDCASE
		IF CPDBF() <> m.iCodePage
			USE
			THIS.SetCodePage(THIS.cTableName, m.iCodePage)
		ENDIF
		USE (THIS.cTableName) EXCLUSIVE
	ENDPROC

	PROCEDURE ValidVersion
		PARAMETERS cVersion
		IF .NOT. m.cVersion == SCCTEXTVER_LOC
			THIS.CANCEL(ERR_BADVERSION_LOC)
		ENDIF
	ENDPROC

	PROCEDURE Fieldlist
		* Returns a CREATE TABLE compatible field list for the current workarea.
		LOCAL cStruct, i
		LOCAL ARRAY aStruct[1]

		AFIELDS(aStruct)
		cStruct = ""
		FOR i = 1 TO ALEN(aStruct, 1)
			IF NOT EMPTY(m.cStruct)
				cStruct = m.cStruct + ","
			ENDIF
			cStruct = m.cStruct + aStruct[m.i, 1] + " " + aStruct[m.i, 2] + ;
				"(" + ALLTRIM(STR(aStruct[m.i, 3])) + "," + ;
				ALLTRIM(STR(aStruct[m.i, 4])) + ")"
		ENDFOR

		RETURN m.cStruct
	ENDPROC

	PROCEDURE CreateVcxCursor
		LOCAL lcDeleted
		lcDeleted = SET("DELETED")
		SET DELETED ON
		PRIVATE iSelect, aClassList, i, j, iCount, aRec, aStruct, cTempCursor, cTempCursor2
		THIS.cVCXCursor = "_" + SYS(3)
		DO WHILE USED(THIS.cVCXCursor)
			THIS.cVCXCursor = "_" + SYS(3)
		ENDDO

		SELECT PADR( uniqueid, FSIZE("uniqueid") ), ;
			RECNO(), ;
			PADR( objname, 50 ) ;
			FROM DBF() ;
			WHERE reserved1 == "Class" ;
			INTO ARRAY aClassList ;
			ORDER BY 3, 1

		iSelect = SELECT() && The original .VCX

		* Create the temporary cursor
		AFIELDS(aStruct)
		CREATE CURSOR (THIS.cVCXCursor) FROM ARRAY aStruct

		* Copy the header record
		SELECT (m.iSelect)
		GO TOP
		SCATTER MEMO TO aRec
		INSERT INTO (THIS.cVCXCursor) FROM ARRAY aRec
		* Scan through the class list and copy the classes over
		IF TYPE('aClassList[1]') <> 'U'
			FOR m.i = 1 TO ALEN(aClassList, 1)
				GO (aClassList[m.i, 2])
				m.iCount = VAL(reserved2) - 1
				SCATTER MEMO TO aRec
				INSERT INTO (THIS.cVCXCursor) FROM ARRAY aRec
				SKIP
				IF m.iCount>0
					cTempCursor = "_" + SYS(3)
					DO WHILE USED(cTempCursor)
						cTempCursor = "_" + SYS(3)
					ENDDO
					CREATE CURSOR (cTempCursor) FROM ARRAY aStruct
					SELECT (m.iSelect)
					FOR m.j = 1 TO m.iCount
						SCATTER MEMO TO aRec
						INSERT INTO (cTempCursor) FROM ARRAY aRec
						SKIP
					ENDFOR
					cTempCursor2 = "_" + SYS(3)
					DO WHILE USED(cTempCursor2)
						cTempCursor2 = "_" + SYS(3)
					ENDDO

					SELECT *, ;
						PADR( LOWER(TRIM(PARENT)+IIF(EMPTY(TRIM(PARENT)),"",".")+IIF(LOWER(BASECLASS)="header","0","")+LOWER(objname)), 254 ) AS SortOrder1, ;
						PADR( SUBSTR(LOWER(TRIM(PARENT))+IIF(EMPTY(TRIM(PARENT)),"",".")+IIF(LOWER(BASECLASS)="header","0","")+LOWER(objname),255), 254 ) AS SortOrder2, ;
						PADR( SUBSTR(LOWER(TRIM(PARENT))+IIF(EMPTY(TRIM(PARENT)),"",".")+IIF(LOWER(BASECLASS)="header","0","")+LOWER(objname),509), 254 ) AS SortOrder3 ;
						FROM (cTempCursor) ORDER BY SortOrder1,SortOrder2,SortOrder3 INTO CURSOR (cTempCursor2) NOFILTER

					SELECT (THIS.cVCXCursor)
					APPEND FROM DBF(cTempCursor2)
					SELECT (m.iSelect)
					USE IN SELECT(cTempCursor)
					USE IN SELECT(cTempCursor2)
				ENDIF
				SCATTER MEMO TO aRec
				INSERT INTO (THIS.cVCXCursor) FROM ARRAY aRec
				SKIP
			ENDFOR
		ENDIF

		* Close the original file and use the cursor we've created
		USE IN (m.iSelect)
		SELECT (THIS.cVCXCursor)

		IF m.lcDeleted == "OFF"
			SET DELETED OFF
		ENDIF
	ENDPROC

	PROCEDURE CreateScxCursor
		LOCAL lnSelect, lcDeleted, lnRecCount, lcThisForm, i, lcFormsetName
		LOCAL ARRAY laForm[1], laRec[1], laStruct[1]

		lnSelect = SELECT(0)			&& The original .SCX
		lcDeleted = SET("DELETED")
		lcFormsetName = ""
		SET DELETED ON
		LOCATE
		COUNT TO lnRecCount
		LOCATE

		THIS.cVCXCursor = "_" + SYS(3)
		DO WHILE USED(THIS.cVCXCursor)
			THIS.cVCXCursor = "_" + SYS(3)
		ENDDO

		* Create the temporary cursor
		AFIELDS(laStruct)
		CREATE CURSOR (THIS.cVCXCursor) FROM ARRAY laStruct

		IF VARTYPE( BASECLASS ) == "U"

			* This is a FP2x screen.
			SELECT (THIS.cVCXCursor)
			APPEND FROM (DBF( m.lnSelect ))
			GO TOP		&& Clear buffer.

		ELSE

			* Copy the header, dataenvironment, cursor, relation, and formset records
			SELECT (m.lnSelect)
			GO TOP
			SCAN WHILE NOT LOWER(BASECLASS) == "form"
				IF LOWER( BASECLASS ) == "formset"
					lcFormsetName = LOWER(objname) + "."
				ENDIF
				SCATTER MEMO TO laRec
				INSERT INTO (THIS.cVCXCursor) FROM ARRAY laRec
			ENDSCAN

			* Copy form by form
			SCAN FOR LOWER(BASECLASS) == "form"
				* Copy the record for the form itself
				SCATTER MEMO TO laRec
				INSERT INTO (THIS.cVCXCursor) FROM ARRAY laRec
				* Sort the records of this form by object path
				lcThisForm = m.lcFormsetName + LOWER(objname) + "."
				SELECT RECNO(), ;
					PADR( LOWER(TRIM(PARENT)+IIF(EMPTY(TRIM(PARENT)),"",".")+IIF(LOWER(BASECLASS)="header","0","")+LOWER(objname)), 254 ), ;
					PADR( SUBSTR(LOWER(TRIM(PARENT))+IIF(EMPTY(TRIM(PARENT)),"",".")+IIF(LOWER(BASECLASS)="header","0","")+LOWER(objname),255), 254 ), ;
					PADR( SUBSTR(LOWER(TRIM(PARENT))+IIF(EMPTY(TRIM(PARENT)),"",".")+IIF(LOWER(BASECLASS)="header","0","")+LOWER(objname),509), 254 ), ;
					PADR( uniqueid, FSIZE("uniqueid") ) ;
					FROM DBF() ;
					WHERE PADR( LOWER(PARENT), LEN(m.lcThisForm), "." ) == m.lcThisForm ;
					INTO ARRAY laForm ;
					ORDER BY 2,3,4,5
				* Copy the records of this form
				FOR i = 1 TO _TALLY
					GO (laForm[i,1])
					SCATTER MEMO TO laRec
					INSERT INTO (THIS.cVCXCursor) FROM ARRAY laRec
				ENDFOR
				SELECT (m.lnSelect)
			ENDSCAN

			* Copy the footer record
			GO BOTTOM
			SCATTER MEMO TO laRec
			INSERT INTO (THIS.cVCXCursor) FROM ARRAY laRec
			GO TOP

			* All records copied?
			ASSERT RECCOUNT(THIS.cVCXCursor) == m.lnRecCount MESSAGE ;
				"Error in CreateScxCursor: " + ;
				LTRIM(STR( m.lnRecCount - RECCOUNT(THIS.cVCXCursor) )) + ;
				" records of " + DBF(m.lnSelect) + " ignored."

		ENDIF

		* Close the original file and use the cursor we've created
		USE IN (m.lnSelect)

		IF m.lcDeleted == "OFF"
			SET DELETED OFF
		ENDIF

		SELECT (THIS.cVCXCursor)
	ENDPROC

	PROCEDURE WriteTextFile
		PRIVATE iCodePage, aText

		USE (THIS.cTableName) EXCLUSIVE

		THIS.oThermRef.iBasis = RECCOUNT()

		m.iCodePage = CPDBF()
		DO CASE
			CASE THIS.cType = PRJTYPE_VCX
				THIS.CreateVcxCursor()
			CASE THIS.cType = PRJTYPE_FORM
				THIS.CreateScxCursor()
		ENDCASE

		THIS.iHandle = FCREATE(THIS.cTextName)
		IF THIS.iHandle = -1
			THIS.Alert(ERR_FCREATE_LOC + THIS.cTextName)
			RETURN -1
		ENDIF

		* First line contains the SCCTEXT version string
		FPUTS(THIS.iHandle, SCCTEXTVER_LOC)

		* Second line contains the CREATE TABLE compatible field list
		FPUTS(THIS.iHandle, THIS.Fieldlist())
		* Third line contains the code page
		FPUTS(THIS.iHandle, ALLTRIM(STR(m.iCodePage)))

		DO CASE
			CASE INLIST(THIS.cType, PRJTYPE_FORM, PRJTYPE_VCX, PRJTYPE_LABEL, PRJTYPE_REPORT, PRJTYPE_MENU, PRJTYPE_DBC)
				THIS.WriteText
			OTHERWISE
				THIS.CANCEL(ERR_UNSUPPORTEDFILETYPE_LOC + m.cType)
		ENDCASE

		FCLOSE(THIS.iHandle)
		THIS.iHandle = -1
		USE
		RETURN 0 && Success
	ENDPROC

	PROCEDURE WriteTable
		PRIVATE cLine, bInMemo, cMemo, cEndMark, bBinary, cFieldname, cValue, iSeconds
		bBinary    = .F.
		bInMemo    = .F.
		cEndMark   = ""
		cFieldname = ""
		cLine      = ""
		cMemo      = ""
		cValue     = ""

		THIS.oThermRef.UPDATE(FSEEK(THIS.iHandle, 0, 1))
		m.iSeconds = SECONDS()

		DO WHILE NOT FEOF(THIS.iHandle)
			IF (SECONDS() - m.iSeconds > 1)
				THIS.oThermRef.UPDATE(FSEEK(THIS.iHandle, 0, 1))
				iSeconds = SECONDS()
			ENDIF

			cLine = FGETS(THIS.iHandle, 8192)

			IF m.bInMemo
				DO CASE
					CASE m.cEndMark == m.cLine
					CASE RAT(m.cEndMark, m.cLine) <> 0
						IF m.bBinary
							cMemo = m.cMemo + THIS.HexStr2BinStr(LEFT(m.cLine, RAT(m.cEndMark, m.cLine) - 1))
						ELSE
							cMemo = m.cMemo + LEFT(m.cLine, RAT(m.cEndMark, m.cLine) - 1)
						ENDIF
					OTHERWISE
						IF m.bBinary
							cMemo = m.cMemo + THIS.HexStr2BinStr(m.cLine)
						ELSE
							cMemo = m.cMemo + m.cLine + CRLF
						ENDIF
						LOOP
				ENDCASE

				* Drop out of if/endif to write the memo field
			ELSE
				DO CASE
					CASE EMPTY(m.cLine)
						LOOP
					CASE m.cLine == MARKEOF
						* Don't read anything past the [EOF] mark
						RETURN
					CASE m.bInMemo .AND. m.cEndMark == m.cLine
					CASE THIS.IsRecordMark(m.cLine)
						APPEND BLANK
						LOOP
					CASE THIS.IsMemoStartMark(m.cLine, @cFieldname)
						bInMemo = .T.
						bBinary = .F.
						cEndMark = THIS.SectionMark(m.cFieldname, .F., .F.)
						LOOP
					CASE THIS.IsBinStartMark(m.cLine, @cFieldname)
						bInMemo = .T.
						bBinary = .T.
						cEndMark = THIS.SectionMark(m.cFieldname, .F., .T.)
						LOOP
					CASE THIS.IsFieldMark(m.cLine, @cFieldname, @cValue)
						DO CASE
							CASE INLIST(TYPE(m.cFieldname), "C", "M")
								REPLACE (m.cFieldname) WITH m.cValue
							CASE TYPE(m.cFieldname) = "N"
								REPLACE (m.cFieldname) WITH VAL(m.cValue)
							CASE TYPE(m.cFieldname) = "L"
								REPLACE (m.cFieldname) WITH &cValue
							OTHERWISE
								THIS.CANCEL(ERR_UNSUPPORTEDFIELDTYPE_LOC + TYPE(m.cFieldname))
						ENDCASE
						LOOP
					OTHERWISE
						IF THIS.Alert(ERR_LINENOACTION_LOC + CHR(13) + CHR(13) + m.cLine + CHR(13) + CHR(13) + ;
								ERR_ALERTCONTINUE_LOC, MB_YESNO) = IDNO
							THIS.CANCEL
						ENDIF
				ENDCASE
			ENDIF

			* Write the memo field
			REPLACE (m.cFieldname) WITH m.cMemo
			bInMemo    = .F.
			cEndMark   = ""
			cFieldname = ""
			cMemo      = ""
		ENDDO
	ENDPROC

	PROCEDURE IsMemoStartMark
		PARAMETERS cLine, cFieldname
		PRIVATE cStartMark, cStartMark2
		IF AT(MARKMEMOSTARTWORD, m.cLine) = 1
			cFieldname = STRTRAN(m.cLine, MARKMEMOSTARTWORD, "", 1, 1)
			cFieldname = LEFT(m.cFieldname, RAT(MARKMEMOSTARTWORD2, m.cFieldname) - 1)
			RETURN .T.
		ENDIF
		RETURN .F.
	ENDPROC

	PROCEDURE IsBinStartMark
		PARAMETERS cLine, cFieldname
		PRIVATE cStartMark, cStartMark2
		IF AT(MARKBINSTARTWORD, m.cLine) = 1
			cFieldname = STRTRAN(m.cLine, MARKBINSTARTWORD, "", 1, 1)
			cFieldname = LEFT(m.cFieldname, RAT(MARKBINSTARTWORD2, m.cFieldname) - 1)
			RETURN .T.
		ENDIF
		RETURN .F.
	ENDPROC

	PROCEDURE IsFieldMark
		PARAMETERS cLine, cFieldname, cValue
		IF AT(MARKFIELDSTART, m.cLine) = 1
			cFieldname = STRTRAN(m.cLine, MARKFIELDSTART, "", 1, 1)
			cFieldname = LEFT(m.cFieldname, AT(MARKFIELDEND, m.cFieldname) - 1)
			cValue = SUBSTR(m.cLine, AT(MARKFIELDEND, m.cLine))
			cValue = STRTRAN(m.cValue, MARKFIELDEND, "", 1, 1)
			RETURN .T.
		ENDIF
		RETURN .F.
	ENDPROC

	PROCEDURE RECORDMARK
		PARAMETERS cUniqueId
		FPUTS(THIS.iHandle, "")
		FPUTS(THIS.iHandle, MARKRECORDSTART + MARKRECORDEND)
	ENDPROC

	PROCEDURE IsRecordMark
		PARAMETERS cLine
		IF LEFT(m.cLine, LEN(MARKRECORDSTART)) == MARKRECORDSTART .AND. ;
				RIGHT(m.cLine, LEN(MARKRECORDEND)) == MARKRECORDEND
			RETURN .T.
		ELSE
			RETURN .F.
		ENDIF
	ENDPROC

	PROCEDURE WriteText
		PRIVATE cExcludeList, cMemoAsCharList, cMemoAsBinList, cCharAsBinList
		cCharAsBinList  = ""
		cExcludeList    = ""
		cMemoAsBinList  = ""
		cMemoAsCharList = ""
		cMemoVariesList = ""

		DO CASE
			CASE INLIST(THIS.cType, PRJTYPE_FORM, PRJTYPE_VCX)
				cCharAsBinList  = VCX_CHARASBIN_LIST
				cExcludeFields  = VCX_EXCLUDE_LIST
				cMemoAsBinList  = VCX_MEMOASBIN_LIST
				cMemoAsCharList = VCX_MEMOASCHAR_LIST
				cMemoVariesList = VCX_MEMOVARIES_LIST
			CASE INLIST(THIS.cType, PRJTYPE_REPORT, PRJTYPE_LABEL)
				cCharAsBinList  = FRX_CHARASBIN_LIST
				cExcludeFields  = FRX_EXCLUDE_LIST
				cMemoAsBinList  = FRX_MEMOASBIN_LIST
				cMemoAsCharList = FRX_MEMOASCHAR_LIST
				cMemoVariesList = FRX_MEMOVARIES_LIST
			CASE THIS.cType == PRJTYPE_MENU
				cCharAsBinList  = MNX_CHARASBIN_LIST
				cExcludeFields  = MNX_EXCLUDE_LIST
				cMemoAsBinList  = MNX_MEMOASBIN_LIST
				cMemoAsCharList = MNX_MEMOASCHAR_LIST
				cMemoVariesList = MNX_MEMOVARIES_LIST
			CASE THIS.cType == PRJTYPE_DBC
				cCharAsBinList  = DBC_CHARASBIN_LIST
				cExcludeFields  = DBC_EXCLUDE_LIST
				cMemoAsBinList  = DBC_MEMOASBIN_LIST
				cMemoAsCharList = DBC_MEMOASCHAR_LIST
				cMemoVariesList = DBC_MEMOVARIES_LIST
			OTHERWISE
				THIS.CANCEL(ERR_UNSUPPORTEDFILETYPE_LOC + THIS.cType)
		ENDCASE

		SCAN
			THIS.oThermRef.UPDATE(RECNO())
			IF TYPE("UNIQUEID") <> 'U'
				THIS.RECORDMARK(uniqueid)
			ENDIF
			FOR m.i = 1 TO FCOUNT()
				IF SKIPEMPTYFIELD AND EMPTY(EVALUATE(FIELD(m.i)))
					LOOP
				ENDIF
				DO CASE
					CASE FIELD(m.i) == "RESERVED1" AND ;
							INLIST(THIS.cType, PRJTYPE_FORM, PRJTYPE_VCX) ;
							AND THIS.MemoIsBinary(FIELD(m.i))
						* do nothing
					CASE FIELD(m.i) == "METHODS" AND ;
							INLIST(THIS.cType, PRJTYPE_FORM, PRJTYPE_VCX)
						THIS.MethodsWrite(FIELD(m.i))
					CASE FIELD(m.i) == "PROPERTIES" AND ;
							INLIST(THIS.cType, PRJTYPE_FORM, PRJTYPE_VCX)
						THIS.PropertiesWrite(FIELD(m.i))
					CASE FIELD(m.i) == "PROTECTED" AND ;
							INLIST(THIS.cType, PRJTYPE_FORM, PRJTYPE_VCX)
						THIS.PropertiesWrite(FIELD(m.i))
					CASE FIELD(m.i) == "RESERVED3" AND ;
							INLIST(THIS.cType, PRJTYPE_FORM, PRJTYPE_VCX)
						THIS.PropertiesWrite(FIELD(m.i))
					CASE " " + FIELD(m.i) + " " $ m.cExcludeFields
						* skip this field
					CASE " " + FIELD(m.i) + " " $ m.cMemoAsCharList
						* memo fields treated as CHAR
						THIS.CharWrite(FIELD(m.i))
					CASE TYPE(FIELD(m.i)) = "C"
						IF " " + FIELD(m.i) + " " $ m.cCharAsBinList
							THIS.MemoWrite(FIELD(m.i), .T.)
						ELSE
							THIS.CharWrite(FIELD(m.i))
						ENDIF
					CASE TYPE(FIELD(m.i)) = "M"
						IF " " + FIELD(m.i) + " " $ m.cMemoVariesList
							* treat as text or binary based on contents of the memofield
							IF THIS.MemoIsBinary(FIELD(m.i))
								THIS.MemoWrite(FIELD(m.i), .T.)
							ELSE
								THIS.MemoWrite(FIELD(m.i), .F.)
							ENDIF
						ELSE
							IF " " + FIELD(m.i) + " " $ m.cMemoAsBinList
								* memo fields treated as BINARY
								THIS.MemoWrite(FIELD(m.i), .T.)
							ELSE
								THIS.MemoWrite(FIELD(m.i), .F.)
							ENDIF
						ENDIF
					CASE TYPE(FIELD(m.i)) = "N"
						THIS.NumWrite(FIELD(m.i))
					CASE TYPE(FIELD(m.i)) = "L"
						THIS.BoolWrite(FIELD(m.i))
					OTHERWISE
						THIS.Alert(ERR_UNSUPPORTEDFIELDTYPE_LOC + TYPE(FIELD(m.i)))
				ENDCASE
			ENDFOR
		ENDSCAN
		THIS.EOFMark()
	ENDPROC

	PROCEDURE MemoIsBinary
		* Scan the memo field to see if it contains binary characters
		PARAMETERS cFieldname
		PRIVATE i, bIsBinary, cMemo
		cMemo     = &cFieldname
		bIsBinary = .T.

		DO CASE
			CASE CHR(0) $ m.cMemo
			OTHERWISE
				bIsBinary = .F.
				IF LEN(m.cMemo) < 126
					FOR m.i = 1 TO LEN(m.cMemo)
						IF ASC(SUBSTR(m.cMemo, m.i, 1)) > 126
							m.bIsBinary = .T.
							EXIT
						ENDIF
					ENDFOR
				ELSE
					FOR m.i = 126 TO 255
						IF CHR(m.i) $ m.cMemo
							m.bIsBinary = .T.
							EXIT
						ENDIF
					ENDFOR
				ENDIF
		ENDCASE
		RETURN m.bIsBinary
	ENDPROC

	PROCEDURE EOFMark
		FPUTS(THIS.iHandle, MARKEOF)
	ENDPROC

	PROCEDURE CharWrite
		PARAMETERS cFieldname
		PRIVATE cTempfield
		cTempfield = &cFieldname
		FPUTS(THIS.iHandle, MARKFIELDSTART + m.cFieldname + MARKFIELDEND + m.cTempfield)
	ENDPROC

	PROCEDURE MemoWrite
		PARAMETERS cFieldname, bBinary
		PRIVATE i, iLen, iStart, cBuf, cBinary, cBinaryProgress, iSeconds
		FPUTS(THIS.iHandle, THIS.SectionMark(m.cFieldname, .T., m.bBinary))
		iLen = LEN(&cFieldname)
		IF m.bBinary
			* If we don't support merging, simply write the checksum
			IF C_WRITECHECKSUMS .AND. TextSupport(THIS.cType) == 1
				FPUTS(THIS.iHandle, MARKCHECKSUM + SYS(2007, &cFieldname))
			ELSE
				cBuf = REPL(CHR(0), 17)

				cBinaryProgress = "0"
				THIS.oThermRef.UpdateTaskMessage(C_BINARYCONVERSION_LOC)
				iSeconds = SECONDS()

				FOR m.i = 1 TO INT(m.iLen / MAXBINLEN) + IIF(m.iLen % MAXBINLEN = 0, 0, 1)
					IF SECONDS() - m.iSeconds > 1
						cBinaryProgress = ALLTRIM(STR(INT(((m.i * MAXBINLEN) / m.iLen) * 100)))
						THIS.oThermRef.UpdateTaskMessage(C_BINARYCONVERSION_LOC)
						iSeconds = SECONDS()
					ENDIF
					cBinary = SUBSTR(&cFieldname, ((m.i - 1) * MAXBINLEN) + 1, MAXBINLEN)
					FOR m.j = 1 TO INT(LEN(m.cBinary) / 8)
						sprintf(@cBuf, "%02X%02X%02X%02X%02X%02X%02X%02X", ;
							ASC(SUBSTR(m.cBinary, ((m.j - 1) * 8) + 1, 1)), ;
							ASC(SUBSTR(m.cBinary, ((m.j - 1) * 8) + 2, 1)), ;
							ASC(SUBSTR(m.cBinary, ((m.j - 1) * 8) + 3, 1)), ;
							ASC(SUBSTR(m.cBinary, ((m.j - 1) * 8) + 4, 1)), ;
							ASC(SUBSTR(m.cBinary, ((m.j - 1) * 8) + 5, 1)), ;
							ASC(SUBSTR(m.cBinary, ((m.j - 1) * 8) + 6, 1)), ;
							ASC(SUBSTR(m.cBinary, ((m.j - 1) * 8) + 7, 1)), ;
							ASC(SUBSTR(m.cBinary, ((m.j - 1) * 8) + 8, 1)))
						FWRITE(THIS.iHandle, m.cBuf, 16)
					ENDFOR
					IF LEN(m.cBinary) % 8 = 0
						FPUTS(THIS.iHandle, "")
					ENDIF
				ENDFOR

				IF LEN(m.cBinary) % 8 <> 0
					cBinary = RIGHT(m.cBinary, LEN(m.cBinary) % 8)
					sprintf(@cBuf, REPLICATE("%02X", LEN(m.cBinary)), ;
						ASC(SUBSTR(m.cBinary, 1, 1)), ;
						ASC(SUBSTR(m.cBinary, 2, 1)), ;
						ASC(SUBSTR(m.cBinary, 3, 1)), ;
						ASC(SUBSTR(m.cBinary, 4, 1)), ;
						ASC(SUBSTR(m.cBinary, 5, 1)), ;
						ASC(SUBSTR(m.cBinary, 6, 1)), ;
						ASC(SUBSTR(m.cBinary, 7, 1)), ;
						ASC(SUBSTR(m.cBinary, 8, 1)))
					FWRITE(THIS.iHandle, m.cBuf, LEN(m.cBinary) * 2)
					FPUTS(THIS.iHandle, "")
				ENDIF

				THIS.oThermRef.UpdateTaskMessage("")
			ENDIF
		ELSE
			FWRITE(THIS.iHandle, &cFieldname)
		ENDIF
		FPUTS(THIS.iHandle, THIS.SectionMark(m.cFieldname, .F., m.bBinary))
	ENDPROC

	PROCEDURE MethodsWrite(cFieldname)
		* write methods in alphabetical order
		FPUTS(THIS.iHandle,  THIS.SectionMark(m.cFieldname, .T.))
		FWRITE(THIS.iHandle, THIS.SortMethods(EVALUATE(m.cFieldname)))
		FPUTS(THIS.iHandle,  THIS.SectionMark(m.cFieldname, .F.))
	ENDPROC


	PROCEDURE SortMethods(tcMethods)
		LOCAL laMethods[1], lnMethods, ln, laLines[1]

		* sanity checks
		ASSERT TYPE("tcMethods") == 'C'
		IF EMPTY(tcMethods)
			RETURN
		ENDIF

		lnMethods = 0
		* for each line in the methods
		FOR ln = 1 TO ALINES(laLines, tcMethods)
			* if it's a procedure line, add a new entry
			IF LEFT(laLines[m.ln], LEN("PROCEDURE ")) == "PROCEDURE "
				lnMethods = m.lnMethods + 1
				DIMENSION laMethods[m.lnMethods]
				laMethods[m.lnMethods] = ""
			ENDIF
			* add line to current entry
			IF m.lnMethods>0
				laMethods[m.lnMethods] = laMethods[m.lnMethods] + laLines[m.ln] + CRLF
			ENDIF
		ENDFOR

		IF lnMethods > 0
			* sort the entries
			IF NOT SORT_CASE_INSENSITIVE OR VERSION(5)<700
				ASORT(laMethods)
			ELSE
				ASORT(laMethods,-1,-1,0,1)
			ENDIF
		ENDIF

		* recreate the methods in method name order
		tcMethods = ""
		FOR ln = 1 TO lnMethods
			tcMethods = m.tcMethods + laMethods[m.ln]
		ENDFOR

		RETURN m.tcMethods

	ENDFUNC

	PROCEDURE PropertiesWrite(cFieldname)
		* write Properties in alphabetical order
		FPUTS(THIS.iHandle,  THIS.SectionMark(m.cFieldname, .T.))
		FWRITE(THIS.iHandle, THIS.SortProperties(EVALUATE(m.cFieldname)))
		FPUTS(THIS.iHandle,  THIS.SectionMark(m.cFieldname, .F.))
	ENDPROC

	PROCEDURE SortProperties(tcProperties)
		* sort Properties by name
		LOCAL laProperties[1], lnProperties, lnLine

		* sanity checks
		ASSERT TYPE("tcProperties") == 'C'
		IF EMPTY(tcProperties)
			RETURN
		ENDIF

		lnProperties = ALINES(laProperties, tcProperties)
		* for each line in the Properties
		FOR lnLine = 1 TO m.lnProperties
			IF RIGHT(laProperties[m.lnLine],2)<>CRLF
				laProperties[m.lnLine] = laProperties[m.lnLine] + CRLF
			ENDIF
		ENDFOR

		* sort the entries
		IF NOT SORT_CASE_INSENSITIVE OR VERSION(5)<700
			ASORT(laProperties)
		ELSE
			ASORT(laProperties,-1,-1,0,1)
		ENDIF

		* recreate the Properties in method name order
		tcProperties = ""
		FOR lnLine = 1 TO m.lnProperties
			tcProperties = m.tcProperties + laProperties[m.lnLine]
		ENDFOR

		RETURN m.tcProperties

	ENDFUNC

	PROCEDURE HexStr2BinStr
		PARAMETERS cHexStr

		LOCAL cBinStr, i, cHexTable, cChar1, cChar2

		cBinStr = ""
		cHexTable = "0123456789ABCDEF"

		FOR i = 1 TO LEN(m.cHexStr) STEP 2
			cChar1 = SUBSTR(m.cHexStr, m.i   , 1)
			cChar2 = SUBSTR(m.cHexStr, m.i +1, 1)
			cBinStr = m.cBinStr + CHR((ATC(m.cChar1, m.cHexTable)-1) * 16 + ATC(m.cChar2, m.cHexTable)-1)
		ENDFOR

		RETURN m.cBinStr
	ENDPROC


	PROCEDURE NumWrite
		* This procedure supports the numerics found in forms, reports, etc. (basically, integers)
		PARAMETERS cFieldname
		FPUTS(THIS.iHandle, MARKFIELDSTART + m.cFieldname + MARKFIELDEND + ALLTRIM(STR(&cFieldname, 20)))
	ENDPROC

	PROCEDURE BoolWrite
		PARAMETERS cFieldname
		FPUTS(THIS.iHandle, MARKFIELDSTART + m.cFieldname + MARKFIELDEND + IIF(&cFieldname, ".T.", ".F."))
	ENDPROC

	PROCEDURE SectionMark
		PARAMETERS cFieldname, lStart, bBinary
		IF m.lStart
			IF m.bBinary
				RETURN MARKBINSTARTWORD + m.cFieldname + MARKBINSTARTWORD2
			ELSE
				RETURN MARKMEMOSTARTWORD + m.cFieldname + MARKMEMOSTARTWORD2
			ENDIF
		ELSE
			IF m.bBinary
				RETURN MARKBINENDWORD + m.cFieldname + MARKBINENDWORD2
			ELSE
				RETURN MARKMEMOENDWORD + m.cFieldname + MARKMEMOENDWORD2
			ENDIF
		ENDIF
	ENDPROC

	PROCEDURE SetCodePage
		PARAMETERS m.fname, m.iCodePage
		PRIVATE iHandle, cpbyte

		DO CASE
			CASE m.iCodePage = 437
				cpbyte = 1
			CASE m.iCodePage = 850
				cpbyte = 2
			CASE m.iCodePage = 1252
				cpbyte = 3
			CASE m.iCodePage = 10000
				cpbyte = 4
			CASE m.iCodePage = 852
				cpbyte = 100
			CASE m.iCodePage = 866
				cpbyte = 101
			CASE m.iCodePage = 865
				cpbyte = 102
			CASE m.iCodePage = 861
				cpbyte = 103
			CASE m.iCodePage = 895
				cpbyte = 104
			CASE m.iCodePage = 620
				cpbyte = 105
			CASE m.iCodePage = 737
				cpbyte = 106
			CASE m.iCodePage = 857
				cpbyte = 107
			CASE m.iCodePage = 863
				cpbyte = 108
			CASE m.iCodePage = 10007
				cpbyte = 150
			CASE m.iCodePage = 10029
				cpbyte = 151
			CASE m.iCodePage = 10006
				cpbyte = 152
			CASE m.iCodePage = 1250
				cpbyte = 200
			CASE m.iCodePage = 1251
				cpbyte = 201
			CASE m.iCodePage = 1253
				cpbyte = 203
			CASE m.iCodePage = 1254
				cpbyte = 202
			CASE m.iCodePage = 1257
				cpbyte = 204
			OTHERWISE
				* Handle the error
				RETURN .F.
		ENDCASE

		m.iHandle = FOPEN(m.fname, 2)
		IF m.iHandle = -1
			RETURN .F.
		ELSE
			FSEEK(m.iHandle, 29)
			FWRITE(m.iHandle, CHR(m.cpbyte))
			FCLOSE(m.iHandle)
		ENDIF
		RETURN .T.
	ENDPROC

ENDDEFINE

******************************************************************
DEFINE CLASS thermometer AS FORM

	AUTOCENTER   = .T.
	BORDERSTYLE  = 2
	HEIGHT       = 88
	NAME         = "thermometer"
	TITLEBAR     = 0
	WIDTH        = 356

	cCurrentTask        = ''
	cThermRef           = ""
	iBasis              = 0
	iPercentage         = 0
	shpThermbarMaxWidth = 322

	* Outer Border
	ADD OBJECT cntBorder1 AS CONTAINER WITH ;
		SPECIALEFFECT = 1,;
		LEFT          = 2 ,;
		TOP           = 2 ,;
		HEIGHT        = THISFORM.HEIGHT - 4 ,;
		WIDTH         = THISFORM.WIDTH - 4

	* Thermometer bar Border
	ADD OBJECT cntBorder2 AS CONTAINER WITH ;
		SPECIALEFFECT = 1,;
		LEFT          = 14 ,;
		TOP           = 44 ,;
		HEIGHT        = 20 ,;
		WIDTH         = 327

	ADD OBJECT lbltitle AS LABEL WITH ;
		BACKSTYLE = 0, ;
		CAPTION   = "", ;
		FONTNAME  = WIN32FONT, ;
		FONTSIZE  = 8, ;
		HEIGHT    = 16, ;
		LEFT      = 18, ;
		TOP       = 14, ;
		WIDTH     = 319

	ADD OBJECT lbltask AS LABEL WITH ;
		BACKSTYLE = 0, ;
		CAPTION   = "", ;
		FONTNAME  = WIN32FONT, ;
		FONTSIZE  = 8, ;
		HEIGHT    = 16, ;
		LEFT      = 18, ;
		TOP       = 27, ;
		WIDTH     = 319


	ADD OBJECT shpThermbar AS SHAPE WITH ;
		BORDERSTYLE = 0, ;
		FILLCOLOR   = RGB(128,128,128), ;
		FILLSTYLE   = 0, ;
		HEIGHT      = 16, ;
		LEFT        = 17, ;
		TOP         = 46, ;
		WIDTH       = 0

	ADD OBJECT lblPercentage AS LABEL WITH ;
		BACKSTYLE = 0, ;
		CAPTION   = "0%", ;
		FONTNAME  = WIN32FONT, ;
		FONTSIZE  = 8, ;
		HEIGHT    = 13, ;
		LEFT      = 170, ;
		TOP       = 47, ;
		WIDTH     = 16

	ADD OBJECT lblPercentage2 AS LABEL WITH ;
		BACKCOLOR = RGB(0,0,255), ;
		BACKSTYLE = 0, ;
		CAPTION   = "Label1", ;
		FONTNAME  = WIN32FONT, ;
		FONTSIZE  = 8, ;
		FORECOLOR = RGB(255,255,255), ;
		HEIGHT    = 13, ;
		LEFT      = 170, ;
		TOP       = 47, ;
		WIDTH     = 0

	ADD OBJECT lblescapemessage AS LABEL WITH ;
		ALIGNMENT = 2, ;
		BACKCOLOR = RGB(192,192,192), ;
		BACKSTYLE = 0, ;
		CAPTION   = "", ;
		FONTBOLD  = .F., ;
		FONTNAME  = WIN32FONT, ;
		FONTSIZE  = 8, ;
		HEIGHT    = 14, ;
		LEFT      = 17, ;
		TOP       = 68, ;
		WIDTH     = 322, ;
		WORDWRAP  = .F.

	PROCEDURE COMPLETE
		* This is the default complete message
		PARAMETERS cTask
		PRIVATE iSeconds
		IF PCOUNT() = 0
			cTask = THERMCOMPLETE_LOC
		ENDIF
		THIS.UPDATE(100,m.cTask)
	ENDPROC

	PROCEDURE UpdateTaskMessage
		&& Update the task message only, used when converting binary data
		PARAMETERS cTask
		THIS.cCurrentTask = m.cTask
		THIS.lbltask.CAPTION = THIS.cCurrentTask
	ENDPROC

	PROCEDURE UPDATE
		&& m.iProgress is the percentage complete
		&& m.cTask is displayed on the second line of the window

		PARAMETERS iProgress, cTask

		IF PCOUNT() >= 2 .AND. TYPE('m.cTask') = 'C'
			&& If we're specifically passed a null string, clear the current task,
			&& otherwise leave it alone
			THIS.cCurrentTask = m.cTask
		ENDIF

		IF NOT THIS.lbltask.CAPTION == THIS.cCurrentTask
			THIS.lbltask.CAPTION = THIS.cCurrentTask
		ENDIF

		IF THIS.iBasis <> 0
			&& interpret m.iProgress in terms of this.iBasis
			iPercentage = INT((m.iProgress / THIS.iBasis) * 100)
		ELSE
			iPercentage = m.iProgress
		ENDIF

		iPercentage = MIN(100,MAX(0,m.iPercentage))

		IF m.iPercentage = THIS.iPercentage
			RETURN
		ENDIF

		IF LEN(ALLTRIM(STR(m.iPercentage,3)))<>LEN(ALLTRIM(STR(THIS.iPercentage,3)))
			iAvgCharWidth = FONTMETRIC(6,THIS.lblPercentage.FONTNAME, ;
				THIS.lblPercentage.FONTSIZE, ;
				IIF(THIS.lblPercentage.FONTBOLD,'B','')+ ;
				IIF(THIS.lblPercentage.FONTITALIC,'I',''))

			THIS.lblPercentage.WIDTH = TXTWIDTH(ALLTRIM(STR(m.iPercentage,3)) + '%', ;
				THIS.lblPercentage.FONTNAME,THIS.lblPercentage.FONTSIZE, ;
				IIF(THIS.lblPercentage.FONTBOLD,'B','')+ ;
				IIF(THIS.lblPercentage.FONTITALIC,'I','')) * iAvgCharWidth

			THIS.lblPercentage.LEFT  = INT((THIS.shpThermbarMaxWidth- THIS.lblPercentage.WIDTH) / 2)+THIS.shpThermbar.LEFT-1
			THIS.lblPercentage2.LEFT = THIS.lblPercentage.LEFT

		ENDIF

		THIS.shpThermbar.WIDTH      = INT((THIS.shpThermbarMaxWidth)*m.iPercentage/100)
		THIS.lblPercentage.CAPTION  = ALLTRIM(STR(m.iPercentage,3)) + '%'
		THIS.lblPercentage2.CAPTION = THIS.lblPercentage.CAPTION

		IF THIS.shpThermbar.LEFT + THIS.shpThermbar.WIDTH -1 >= THIS.lblPercentage2.LEFT
			IF THIS.shpThermbar.LEFT + THIS.shpThermbar.WIDTH - 1 >= THIS.lblPercentage2.LEFT + THIS.lblPercentage.WIDTH - 1
				THIS.lblPercentage2.WIDTH = THIS.lblPercentage.WIDTH
			ELSE
				THIS.lblPercentage2.WIDTH = THIS.shpThermbar.LEFT + THIS.shpThermbar.WIDTH - THIS.lblPercentage2.LEFT - 1
			ENDIF
		ELSE
			THIS.lblPercentage2.WIDTH = 0
		ENDIF
		THIS.iPercentage = m.iPercentage
	ENDPROC

	PROCEDURE INIT
		PARAMETERS cTitle, iInterval
		&& m.cTitle is displayed on the first line of the window
		&& m.iInterval is the frequency used for updating the thermometer

		THIS.lbltitle.CAPTION = IIF(EMPTY(m.cTitle),'',m.cTitle)
		LOCAL cColor

		&& Check to see if the fontmetrics for MS Sans Serif matches
		&& those on the system developed. If not, switch to Arial.
		&& The RETURN value indicates whether the font was changed.
		IF FONTMETRIC(1, WIN32FONT, 8, '') <> 13 .OR. ;
				FONTMETRIC(4, WIN32FONT, 8, '') <> 2 .OR. ;
				FONTMETRIC(6, WIN32FONT, 8, '') <> 5 .OR. ;
				FONTMETRIC(7, WIN32FONT, 8, '') <> 11
			THIS.SETALL('FontName', WIN95FONT)
		ENDIF

	ENDPROC
ENDDEFINE


************************************************************
* EOF
************************************************************
