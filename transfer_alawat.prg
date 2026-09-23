* transfer_alawat.prg
* DO C:\Users\ngc\Downloads\transfer_alawat.prg
* Use the already-open AL082026 table and the xlsx sitting next to it.

LOCAL lcDbf, lcXls, lcDir, lnOld, lnNew, lnWanted
LOCAL loExcel, loBook, loSheet, llExcelOk
LOCAL lnHeader, lnCols, lnRow, lnLast, lnIns, lnBlank
LOCAL lnFCount, i, lnCol, lcField, lcHdr, lcKey, lcVal, lcType
LOCAL lnPnoCol, lnNameCol, lnWidth, lnDec
LOCAL lnExcelNums, lnUsedNum
LOCAL ldDate, lcAmer
LOCAL laFields[1], laCol[1], laExcelNum[1], laDir[1], lnDir

lnWanted = 126
lcAmer = "8276"
ldDate = DATE(2026, 8, 13)
llExcelOk = .F.

SET TALK OFF
SET SAFETY OFF
SET EXACT OFF
SET DELETED ON
SET CONFIRM OFF

lcDbf = ""
IF USED("AL082026")
    lcDbf = DBF("AL082026")
ENDIF
IF EMPTY(lcDbf) AND !EMPTY(ALIAS())
    IF ATC("AL082026", ALIAS()) > 0 OR ATC("AL082026", DBF()) > 0
        lcDbf = DBF()
    ENDIF
ENDIF
IF EMPTY(lcDbf)
    lcDbf = GETFILE("DBF", "AL082026")
ENDIF
IF EMPTY(lcDbf)
    MESSAGEBOX("AL082026.DBF not found")
    RETURN
ENDIF

lcDir = LEFT(lcDbf, RAT("\", lcDbf))
lcXls = ""
lnDir = ADIR(laDir, lcDir + "*.xlsx")
IF TYPE("lnDir") = "N" AND lnDir > 0
    FOR i = 1 TO lnDir
        IF ATC("8276", laDir[i, 1]) > 0
            lcXls = lcDir + laDir[i, 1]
            EXIT
        ENDIF
    ENDFOR
    IF EMPTY(lcXls)
        lcXls = lcDir + laDir[1, 1]
    ENDIF
ENDIF
IF EMPTY(lcXls)
    lcXls = GETFILE("XLSX", "Excel")
ENDIF
IF EMPTY(lcXls)
    MESSAGEBOX("Excel file not found in" + CHR(13) + lcDir)
    RETURN
ENDIF

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
    MESSAGEBOX("Cannot open Excel. Is Microsoft Excel installed?")
    RETURN
ENDIF

IF USED("AL082026")
    SELECT AL082026
    USE
ENDIF
USE (lcDbf) ALIAS AL082026 EXCLUSIVE
IF !ISEXCLUSIVE("AL082026")
    TRY
        loBook.Close(.F.)
        loExcel.Quit()
    CATCH
    ENDTRY
    MESSAGEBOX("Close BROWSE first, then DO the program again.")
    RETURN
ENDIF

lnOld = RECCOUNT()
COPY TO (lcDir + "AL082026_BAK_BEFORE_ZAP")
ZAP

lnFCount = AFIELDS(laFields)
DIMENSION laCol[lnFCount]
FOR i = 1 TO lnFCount
    laCol[i] = 0
ENDFOR

lnCols = 20
lnLast = 200
TRY
    lnCols = loSheet.UsedRange.Columns.Count
    lnLast = loSheet.UsedRange.Rows.Count
CATCH
ENDTRY
IF VARTYPE(lnCols) <> "N" OR lnCols < 1
    lnCols = 20
ENDIF
IF VARTYPE(lnLast) <> "N" OR lnLast < 2
    lnLast = 200
ENDIF

lnHeader = 1
lnPnoCol = DetectPnoCol(loSheet, 2, MIN(30, lnLast), lnCols)
IF lnPnoCol = 0
    lnPnoCol = 1
ENDIF
lnNameCol = DetectNameCol(loSheet, 2, MIN(30, lnLast), lnCols, lnPnoCol)

FOR i = 1 TO lnFCount
    lcField = UPPER(ALLTRIM(laFields[i, 1]))
    IF lcField == "PNO"
        laCol[i] = lnPnoCol
    ENDIF
    IF lcField == "NAME"
        laCol[i] = lnNameCol
    ENDIF
ENDFOR

lnExcelNums = 0
FOR lnCol = 1 TO lnCols
    IF lnCol = lnPnoCol OR lnCol = lnNameCol
        LOOP
    ENDIF
    IF LooksLikeAmountCol(loSheet, 2, MIN(30, lnLast), lnCol)
        lnExcelNums = lnExcelNums + 1
        DIMENSION laExcelNum[lnExcelNums]
        laExcelNum[lnExcelNums] = lnCol
    ENDIF
ENDFOR

lnUsedNum = 0
FOR i = 1 TO lnFCount
    IF laCol[i] > 0
        LOOP
    ENDIF
    lcField = UPPER(ALLTRIM(laFields[i, 1]))
    IF INLIST(lcField, "MONEYCOM", "MONEYPOST", "FRK", "BAS", "AMT", "ALLW")
        lnUsedNum = lnUsedNum + 1
        IF lnUsedNum <= lnExcelNums
            laCol[i] = laExcelNum[lnUsedNum]
        ENDIF
    ENDIF
ENDFOR

* If header row looks like titles, skip it.
lcKey = CellText(loSheet, 1, lnPnoCol)
IF LEN(CHRTRAN(ALLTRIM(lcKey), "0123456789", "")) > 0
    lnHeader = 1
ELSE
    lnHeader = 0
ENDIF

lnIns = 0
lnBlank = 0
FOR lnRow = lnHeader + 1 TO lnLast
    lcKey = ALLTRIM(CellText(loSheet, lnRow, lnPnoCol))
    IF EMPTY(lcKey) OR LEN(CHRTRAN(lcKey, "0123456789", "")) > 0
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
        ENDIF
        IF EMPTY(lcVal)
            DO CASE
            CASE UPPER(lcField) = "NO"
                lcVal = lcAmer
            CASE UPPER(lcField) = "MO"
                lcVal = "8"
            CASE lcType = "D"
                PutField(lcField, lcType, lnWidth, lnDec, ldDate)
                LOOP
            ENDCASE
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

lnNew = RECCOUNT()
GO TOP
SET FILTER TO
SET TALK ON
COUNT
SET TALK OFF

MESSAGEBOX("Old: " + TRANSFORM(lnOld) + CHR(13) + ;
    "New: " + TRANSFORM(lnNew) + CHR(13) + ;
    "Inserted: " + TRANSFORM(lnIns) + CHR(13) + ;
    "PNO col=" + TRANSFORM(lnPnoCol) + " Name col=" + TRANSFORM(lnNameCol) + CHR(13) + ;
    "Excel: " + lcXls)

RETURN


FUNCTION CellText
LPARAMETERS toSheet, tnRow, tnCol
    LOCAL lu, lc
    lc = ""
    IF VARTYPE(tnRow) <> "N" OR VARTYPE(tnCol) <> "N" OR tnRow < 1 OR tnCol < 1
        RETURN ""
    ENDIF
    TRY
        lu = toSheet.Cells(tnRow, tnCol).Value
        DO CASE
        CASE ISNULL(lu)
            lc = ""
        CASE VARTYPE(lu) = "N"
            IF lu = INT(lu)
                lc = ALLTRIM(STR(lu, 18, 0))
            ELSE
                lc = ALLTRIM(STR(lu, 18, 4))
            ENDIF
        CASE VARTYPE(lu) = "D"
            lc = DTOC(lu)
        CASE VARTYPE(lu) = "T"
            lc = TTOC(lu)
        CASE VARTYPE(lu) = "C"
            lc = ALLTRIM(lu)
        OTHERWISE
            lc = ALLTRIM(TRANSFORM(lu))
        ENDCASE
    CATCH
        lc = ""
    ENDTRY
    RETURN lc
ENDFUNC


FUNCTION DetectPnoCol
LPARAMETERS toSheet, tnFrom, tnTo, tnCols
    LOCAL lnCol, lnRow, lnBest, lnBestCol, lnHits, lc
    lnBest = 0
    lnBestCol = 1
    IF tnFrom < 1
        tnFrom = 1
    ENDIF
    IF tnTo < tnFrom
        tnTo = tnFrom
    ENDIF
    FOR lnCol = 1 TO tnCols
        lnHits = 0
        FOR lnRow = tnFrom TO tnTo
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


FUNCTION DetectNameCol
LPARAMETERS toSheet, tnFrom, tnTo, tnCols, tnPnoCol
    LOCAL lnCol, lnRow, lnBest, lnBestCol, lnHits, lc
    lnBest = 0
    lnBestCol = 0
    IF tnFrom < 1
        tnFrom = 1
    ENDIF
    IF tnTo < tnFrom
        tnTo = tnFrom
    ENDIF
    FOR lnCol = 1 TO tnCols
        IF lnCol = tnPnoCol
            LOOP
        ENDIF
        lnHits = 0
        FOR lnRow = tnFrom TO tnTo
            lc = ALLTRIM(CellText(toSheet, lnRow, lnCol))
            IF LEN(lc) >= 4 AND LEN(CHRTRAN(lc, "0123456789 .,-/", "")) > 2
                lnHits = lnHits + 1
            ENDIF
        ENDFOR
        IF lnHits > lnBest
            lnBest = lnHits
            lnBestCol = lnCol
        ENDIF
    ENDFOR
    IF lnBestCol = 0 AND tnPnoCol < tnCols
        lnBestCol = tnPnoCol + 1
    ENDIF
    RETURN lnBestCol
ENDFUNC


FUNCTION LooksLikeAmountCol
LPARAMETERS toSheet, tnFrom, tnTo, tnCol
    LOCAL lnRow, lnHits, lc
    lnHits = 0
    IF tnFrom < 1
        tnFrom = 1
    ENDIF
    IF tnTo < tnFrom
        RETURN .F.
    ENDIF
    FOR lnRow = tnFrom TO tnTo
        lc = ALLTRIM(CellText(toSheet, lnRow, tnCol))
        IF !EMPTY(lc) AND VAL(CHRTRAN(lc, ",", "")) <> 0
            lnHits = lnHits + 1
        ENDIF
    ENDFOR
    RETURN lnHits >= 3
ENDFUNC


PROCEDURE PutField
LPARAMETERS tcField, tcType, tnWidth, tnDec, tuVal
    LOCAL lc, ln, ld, lnY, lnM, lnD
    tcType = UPPER(tcType)
    DO CASE
    CASE tcType = "C"
        lc = ALLTRIM(TRANSFORM(tuVal))
        IF tnWidth > 0
            lc = LEFT(lc, tnWidth)
        ENDIF
        REPLACE (tcField) WITH lc
    CASE tcType = "N" OR tcType = "F" OR tcType = "B" OR tcType = "Y" OR tcType = "I"
        IF VARTYPE(tuVal) = "N"
            ln = tuVal
        ELSE
            ln = VAL(CHRTRAN(ALLTRIM(TRANSFORM(tuVal)), ",", ""))
        ENDIF
        REPLACE (tcField) WITH ln
    CASE tcType = "D"
        ld = {}
        IF VARTYPE(tuVal) = "D"
            ld = tuVal
        ELSE
            lc = ALLTRIM(TRANSFORM(tuVal))
            lc = STRTRAN(STRTRAN(lc, "/", "-"), ".", "-")
            IF LEN(lc) >= 8 AND AT("-", lc) = 0 AND LEN(CHRTRAN(LEFT(lc, 8), "0123456789", "")) = 0
                lnY = VAL(LEFT(lc, 4))
                lnM = VAL(SUBSTR(lc, 5, 2))
                lnD = VAL(RIGHT(LEFT(lc, 8), 2))
                IF lnY >= 1900 AND lnM >= 1 AND lnM <= 12 AND lnD >= 1 AND lnD <= 31
                    ld = DATE(lnY, lnM, lnD)
                ENDIF
            ENDIF
        ENDIF
        IF !EMPTY(ld)
            REPLACE (tcField) WITH ld
        ENDIF
    CASE tcType = "L"
        REPLACE (tcField) WITH INLIST(UPPER(ALLTRIM(TRANSFORM(tuVal))), "T", "Y", "1", ".T.")
    ENDCASE
ENDPROC
