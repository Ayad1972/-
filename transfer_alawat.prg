* transfer_alawat.prg
* Run inside the SAME Visual FoxPro window that shows 222 records:
*   DO transfer_alawat.prg
* This ZAP the open table and loads 126 Excel rows. Python cannot do that
* while FoxPro has the file locked.

LOCAL lcDbf, lcXls, lnOld, lnNew, lnWanted
LOCAL loExcel, loBook, loSheet, llExcelOk
LOCAL lnHeader, lnCols, lnRow, lnLast, lnIns, lnBlank
LOCAL lnFCount, i, lnCol, lcField, lcHdr, lcKey, lcVal, lcType
LOCAL lnPnoCol, lnMapped, lnF, lnWidth, lnDec
LOCAL luVal, ldDate, lcAmer
LOCAL laFields[1], laCol[1]

lnWanted = 126
lcAmer = "8276"
ldDate = DATE(2026, 8, 13)
llExcelOk = .F.
lnNew = -1

SET TALK OFF
SET SAFETY OFF
SET EXACT OFF
SET DELETED ON
SET CONFIRM OFF

lcDbf = FindAlDbf()
lcXls = FindAlXls()
IF EMPTY(lcDbf)
    lcDbf = GETFILE("DBF", "AL082026.DBF")
ENDIF
IF EMPTY(lcXls)
    lcXls = GETFILE("XLSX", "Excel")
ENDIF
IF EMPTY(lcDbf) OR !FILE(lcDbf)
    MESSAGEBOX("AL082026.DBF not found")
    RETURN
ENDIF
IF EMPTY(lcXls) OR !FILE(lcXls)
    MESSAGEBOX("Excel file not found")
    RETURN
ENDIF

* Close alias then reopen exclusive so ZAP is allowed.
IF USED("AL082026")
    SELECT AL082026
    USE
ENDIF
USE (lcDbf) ALIAS AL082026 EXCLUSIVE
IF !ISEXCLUSIVE("AL082026")
    MESSAGEBOX("Cannot open AL082026 exclusive. Close BROWSE then DO again.")
    RETURN
ENDIF

lnOld = RECCOUNT()
COPY TO (ADDBS(JUSTPATH(lcDbf)) + "AL082026_BAK_BEFORE_ZAP")
ZAP
IF RECCOUNT() <> 0
    MESSAGEBOX("ZAP failed. Count is still " + TRANSFORM(RECCOUNT()))
    RETURN
ENDIF

lnFCount = AFIELDS(laFields)
DIMENSION laCol[lnFCount]
FOR i = 1 TO lnFCount
    laCol[i] = 0
ENDFOR

TRY
    loExcel = CREATEOBJECT("Excel.Application")
    loExcel.Visible = .F.
    loExcel.DisplayAlerts = .F.
    loBook = loExcel.Workbooks.Open(lcXls, 0, .T.)
    loSheet = loBook.Worksheets(1)
    llExcelOk = .T.
CATCH
    llExcelOk = .F.
ENDTRY

IF !llExcelOk
    MESSAGEBOX("Excel COM failed. Is Excel installed?")
    RETURN
ENDIF

lnCols = loSheet.UsedRange.Columns.Count
IF lnCols < 1
    lnCols = 20
ENDIF
lnLast = loSheet.UsedRange.Rows.Count
IF lnLast < 2
    lnLast = 200
ENDIF

lnHeader = 1
FOR lnRow = 1 TO MIN(15, lnLast)
    IF HeaderScore(loSheet, lnRow, lnCols, @laFields, lnFCount) ;
            > HeaderScore(loSheet, lnHeader, lnCols, @laFields, lnFCount)
        lnHeader = lnRow
    ENDIF
ENDFOR

lnMapped = 0
lnPnoCol = 0
FOR i = 1 TO lnFCount
    lcField = UPPER(ALLTRIM(laFields[i, 1]))
    FOR lnCol = 1 TO lnCols
        lcHdr = NormHdr(CellText(loSheet, lnHeader, lnCol))
        IF EMPTY(lcHdr)
            LOOP
        ENDIF
        IF lcHdr == lcField OR MatchAlias(lcField, lcHdr)
            laCol[i] = lnCol
            lnMapped = lnMapped + 1
            IF lcField == "PNO"
                lnPnoCol = lnCol
            ENDIF
            EXIT
        ENDIF
    ENDFOR
ENDFOR

IF lnPnoCol = 0
    lnPnoCol = DetectPnoCol(loSheet, lnHeader + 1, lnLast, lnCols)
    FOR i = 1 TO lnFCount
        IF UPPER(ALLTRIM(laFields[i, 1])) == "PNO"
            laCol[i] = lnPnoCol
        ENDIF
    ENDFOR
ENDIF
IF lnMapped = 0 AND lnCols = lnFCount
    FOR i = 1 TO lnFCount
        laCol[i] = i
    ENDFOR
ENDIF
IF lnPnoCol = 0
    FOR i = 1 TO lnFCount
        IF UPPER(ALLTRIM(laFields[i, 1])) == "PNO"
            lnPnoCol = laCol[i]
        ENDIF
    ENDFOR
ENDIF

lnIns = 0
lnBlank = 0
FOR lnRow = lnHeader + 1 TO lnLast
    lcKey = ""
    IF lnPnoCol > 0
        lcKey = ALLTRIM(CellText(loSheet, lnRow, lnPnoCol))
    ELSE
        lcKey = ALLTRIM(CellText(loSheet, lnRow, 1))
    ENDIF
    IF EMPTY(lcKey)
        lnBlank = lnBlank + 1
        IF lnIns > 0 AND lnBlank >= 2
            EXIT
        ENDIF
        LOOP
    ENDIF
    lnBlank = 0
    IF lnIns >= lnWanted
        EXIT
    ENDIF

    APPEND BLANK
    FOR i = 1 TO lnFCount
        lcField = ALLTRIM(laFields[i, 1])
        lcType = UPPER(laFields[i, 2])
        lnWidth = laFields[i, 3]
        lnDec = laFields[i, 4]
        lcVal = ""
        IF laCol[i] > 0
            lcVal = CellText(loSheet, lnRow, laCol[i])
        ELSE
            IF INLIST(UPPER(lcField), "AMER", "AMR", "ORDER", "ORDNO", "NOAMR")
                lcVal = lcAmer
            ENDIF
            IF INLIST(UPPER(lcField), "DT", "DATE", "FDATE")
                PutField(lcField, lcType, lnWidth, lnDec, ldDate)
                LOOP
            ENDIF
        ENDIF
        IF EMPTY(lcVal) AND INLIST(UPPER(lcField), "AMER", "AMR", "ORDER", "ORDNO", "NOAMR")
            lcVal = lcAmer
        ENDIF
        PutField(lcField, lcType, lnWidth, lnDec, lcVal)
    ENDFOR
    lnIns = lnIns + 1
ENDFOR

TRY
    loBook.Close(.F.)
    loExcel.Quit()
CATCH
ENDTRY
loSheet = .NULL.
loBook = .NULL.
loExcel = .NULL.

IF FILE(FORCEEXT(lcDbf, "CDX"))
    REINDEX
ENDIF

lnNew = RECCOUNT()
GO TOP
SET FILTER TO
SET TALK ON
COUNT
SET TALK OFF

MESSAGEBOX("Old count: " + TRANSFORM(lnOld) + CHR(13) + ;
    "New count: " + TRANSFORM(lnNew) + CHR(13) + ;
    "Excel rows loaded: " + TRANSFORM(lnIns) + CHR(13) + ;
    "Required: 126")

RETURN


FUNCTION CellText
LPARAMETERS toSheet, tnRow, tnCol
    LOCAL lu
    lu = toSheet.Cells(tnRow, tnCol).Value
    IF ISNULL(lu)
        RETURN ""
    ENDIF
    DO CASE
    CASE VARTYPE(lu) = "N"
        IF lu = INT(lu)
            RETURN ALLTRIM(STR(lu, 18, 0))
        ENDIF
        RETURN ALLTRIM(STR(lu, 18, 4))
    CASE VARTYPE(lu) = "D"
        RETURN DTOC(lu)
    CASE VARTYPE(lu) = "T"
        RETURN TTOC(lu)
    CASE VARTYPE(lu) = "L"
        RETURN IIF(lu, "T", "F")
    OTHERWISE
        RETURN ALLTRIM(TRANSFORM(lu))
    ENDCASE
ENDFUNC


FUNCTION NormHdr
LPARAMETERS tc
    LOCAL lc
    lc = UPPER(ALLTRIM(tc))
    lc = STRTRAN(lc, " ", "")
    lc = STRTRAN(lc, "_", "")
    lc = STRTRAN(lc, "-", "")
    lc = STRTRAN(lc, ".", "")
    RETURN lc
ENDFUNC


FUNCTION MatchAlias
LPARAMETERS tcField, tcHdr
    tcField = UPPER(ALLTRIM(tcField))
    tcHdr = UPPER(ALLTRIM(tcHdr))
    DO CASE
    CASE tcField = "PNO"
        RETURN INLIST(tcHdr, "PNO", "EMPNO", "EMPNO", "NO", "NUM", "NUMBER", "ID", "CODE")
    CASE tcField = "NAME"
        RETURN INLIST(tcHdr, "NAME", "ENAME", "ANAME", "EMPNAME", "FULLNAME")
    CASE INLIST(tcField, "AMT", "ALLW")
        RETURN INLIST(tcHdr, "AMT", "AMOUNT", "ALLW", "ALLOW", "ALAWA", "SAL", "VALUE", "VAL")
    CASE INLIST(tcField, "TYPE", "TYP")
        RETURN INLIST(tcHdr, "TYPE", "TYP", "KIND", "CLASS")
    CASE INLIST(tcField, "AMER", "AMR", "ORDER", "ORDNO", "NOAMR")
        RETURN INLIST(tcHdr, "AMER", "AMR", "ORDER", "ORDNO", "NOAMR", "BOOK")
    CASE INLIST(tcField, "DT", "DATE", "FDATE")
        RETURN INLIST(tcHdr, "DT", "DATE", "FDATE", "HDATE")
    CASE INLIST(tcField, "NOTE", "NOTES")
        RETURN INLIST(tcHdr, "NOTE", "NOTES", "REMARK", "REM")
    ENDCASE
    RETURN .F.
ENDFUNC


FUNCTION HeaderScore
LPARAMETERS toSheet, tnRow, tnCols, taFields, tnFCount
    EXTERNAL ARRAY taFields
    LOCAL lnScore, lnCol, lcHdr, i
    lnScore = 0
    FOR lnCol = 1 TO tnCols
        lcHdr = NormHdr(CellText(toSheet, tnRow, lnCol))
        IF EMPTY(lcHdr)
            LOOP
        ENDIF
        lnScore = lnScore + 1
        FOR i = 1 TO tnFCount
            IF lcHdr == UPPER(ALLTRIM(taFields[i, 1])) OR MatchAlias(taFields[i, 1], lcHdr)
                lnScore = lnScore + 4
            ENDIF
        ENDFOR
    ENDFOR
    RETURN lnScore
ENDFUNC


FUNCTION DetectPnoCol
LPARAMETERS toSheet, tnFrom, tnTo, tnCols
    LOCAL lnCol, lnRow, lnBest, lnBestCol, lnHits, lc
    lnBest = 0
    lnBestCol = 1
    FOR lnCol = 1 TO tnCols
        lnHits = 0
        FOR lnRow = tnFrom TO MIN(tnFrom + 25, tnTo)
            lc = ALLTRIM(CellText(toSheet, lnRow, lnCol))
            IF LEN(lc) >= 4 AND LEN(lc) <= 8 AND LEN(CHRTRAN(lc, "0123456789", "")) = 0
                lnHits = lnHits + 1
            ENDIF
        ENDFOR
        IF lnHits > lnBest
            lnBest = lnHits
            lnBestCol = lnCol
        ENDIF
    ENDFOR
    RETURN lnBestCol
ENDFUNC


PROCEDURE PutField
LPARAMETERS tcField, tcType, tnWidth, tnDec, tuVal
    LOCAL lc, ln, ld
    tcType = UPPER(tcType)
    DO CASE
    CASE tcType = "C"
        lc = ALLTRIM(TRANSFORM(tuVal))
        REPLACE (tcField) WITH LEFT(lc, tnWidth)
    CASE tcType = "N" OR tcType = "F" OR tcType = "B" OR tcType = "Y"
        IF VARTYPE(tuVal) = "N"
            ln = tuVal
        ELSE
            ln = VAL(CHRTRAN(ALLTRIM(TRANSFORM(tuVal)), ",", ""))
        ENDIF
        REPLACE (tcField) WITH ln
    CASE tcType = "D"
        IF VARTYPE(tuVal) = "D"
            ld = tuVal
        ELSE
            ld = ParseDate(TRANSFORM(tuVal))
        ENDIF
        IF !EMPTY(ld)
            REPLACE (tcField) WITH ld
        ENDIF
    CASE tcType = "L"
        REPLACE (tcField) WITH INLIST(UPPER(ALLTRIM(TRANSFORM(tuVal))), "T", "Y", "1", ".T.")
    ENDCASE
ENDPROC


FUNCTION ParseDate
LPARAMETERS tc
    LOCAL lc, d, m, y, lnAt1, lnAt2
    lc = ALLTRIM(tc)
    lc = STRTRAN(lc, "/", "-")
    lc = STRTRAN(lc, ".", "-")
    IF LEN(lc) = 8 AND LEN(CHRTRAN(lc, "0123456789", "")) = 0
        RETURN DATE(VAL(LEFT(lc, 4)), VAL(SUBSTR(lc, 5, 2)), VAL(RIGHT(lc, 2)))
    ENDIF
    lnAt1 = AT("-", lc)
    lnAt2 = RAT("-", lc)
    IF lnAt1 > 0 AND lnAt2 > lnAt1
        d = VAL(LEFT(lc, lnAt1 - 1))
        m = VAL(SUBSTR(lc, lnAt1 + 1, lnAt2 - lnAt1 - 1))
        y = VAL(SUBSTR(lc, lnAt2 + 1))
        IF y < 100
            y = y + 2000
        ENDIF
        IF d > 0 AND m > 0 AND y > 0
            RETURN DATE(y, m, d)
        ENDIF
    ENDIF
    RETURN {}
ENDFUNC


FUNCTION FindAlDbf
    LOCAL lcRoot, lc
    lcRoot = ADDBS(GETENV("USERPROFILE")) + "Desktop\092026"
    lc = FindNamed(lcRoot, "AL082026.DBF")
    IF !EMPTY(lc)
        RETURN lc
    ENDIF
    RETURN FindNamed(ADDBS(GETENV("USERPROFILE")) + "Desktop", "AL082026.DBF")
ENDFUNC


FUNCTION FindAlXls
    LOCAL lcRoot, lc
    lcRoot = ADDBS(GETENV("USERPROFILE")) + "Desktop\092026"
    lc = FindXlsxPref(lcRoot)
    IF !EMPTY(lc)
        RETURN lc
    ENDIF
    RETURN FindXlsxPref(ADDBS(GETENV("USERPROFILE")) + "Desktop")
ENDFUNC


FUNCTION FindNamed
LPARAMETERS tcRoot, tcName
    LOCAL la[1], ln, i, lc
    IF EMPTY(tcRoot) OR !DIRECTORY(tcRoot)
        RETURN ""
    ENDIF
    IF FILE(ADDBS(tcRoot) + tcName)
        RETURN ADDBS(tcRoot) + tcName
    ENDIF
    ln = ADIR(la, ADDBS(tcRoot) + "*.*", "D")
    FOR i = 1 TO ln
        IF la[i, 1] = "." OR la[i, 1] = ".."
            LOOP
        ENDIF
        IF "D" $ la[i, 5]
            lc = FindNamed(ADDBS(tcRoot) + la[i, 1], tcName)
            IF !EMPTY(lc)
                RETURN lc
            ENDIF
        ENDIF
    ENDFOR
    RETURN ""
ENDFUNC


FUNCTION FindXlsxPref
LPARAMETERS tcRoot
    LOCAL la[1], ln, i, lc, lcBest, lcName
    lcBest = ""
    IF EMPTY(tcRoot) OR !DIRECTORY(tcRoot)
        RETURN ""
    ENDIF
    ln = ADIR(la, ADDBS(tcRoot) + "*.xlsx")
    FOR i = 1 TO ln
        lcName = UPPER(la[i, 1])
        lc = ADDBS(tcRoot) + la[i, 1]
        IF "8276" $ lcName
            RETURN lc
        ENDIF
        IF EMPTY(lcBest)
            lcBest = lc
        ENDIF
    ENDFOR
    ln = ADIR(la, ADDBS(tcRoot) + "*.*", "D")
    FOR i = 1 TO ln
        IF la[i, 1] = "." OR la[i, 1] = ".."
            LOOP
        ENDIF
        IF "D" $ la[i, 5]
            lc = FindXlsxPref(ADDBS(tcRoot) + la[i, 1])
            IF !EMPTY(lc)
                IF "8276" $ UPPER(JUSTFNAME(lc))
                    RETURN lc
                ENDIF
                IF EMPTY(lcBest)
                    lcBest = lc
                ENDIF
            ENDIF
        ENDIF
    ENDFOR
    RETURN lcBest
ENDFUNC
