Option Explicit

Private Const CLASS_NAME As String = "modInputLHS"
Private Const LHS_SHEET_NAME As String = "FAMOO-Full Inventory IFRS"
Private Const ERR_LHS As Long = vbObjectError + 6400

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strFullPath - selected LHS input; lngDataRows - returned row count
' Returns:       Variant - 2D output array including header row
' Description:   Extracts the 14 required LHS columns in the requested order.
'-------------------------------------------------------------------------------
Public Function LoadLHSInput( _
    ByVal strFullPath As String, ByRef lngDataRows As Long) As Variant

    Const METHOD_NAME As String = "LoadLHSInput"
    Dim arrColumn As Variant
    Dim arrColumns As Variant
    Dim arrHeaders As Variant
    Dim arrOutput() As Variant
    Dim blnOpenedByMacro As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColIndex As Long
    Dim lngLastRow As Long
    Dim lngOutputColumn As Long
    Dim lngRow As Long
    Dim wkbSource As Excel.Workbook
    Dim wksSource As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    ' Fixed LHS source columns: ED=134, EK=141, BP=68 and BS=71.
    arrColumns = Array(1&, 9&, 115&, 114&, 116&, 130&, 134&, 141&, 68&, 71&, 66&, 35&, 15&, 16&)
    arrHeaders = Array( _
        "FUND CODE", _
        "EXTERNAL VALUE CODE", _
        "LAST COUPON DATE", _
        "NEXT COUPON DATE", _
        "NBR OF DAYS SINCE LAST COUPON", _
        "GTI NAME", _
        "ACCOUNTING DATE_1", _
        "GENERATION DATE AND TIME", _
        "Coupon frequency/IRS-CDS", _
        "BOND CALC METHOD/IRS-CDS", _
        "START DATE", _
        "INTEREST RATE", _
        "ACCRUED INTEREST IN ACC'S CCY", _
        "ACCRUED INTEREST IN FUND CCY")

    Set wkbSource = GetOrOpenWorkbookReadOnly(strFullPath, blnOpenedByMacro)
    Set wksSource = GetWorksheetByName(wkbSource, LHS_SHEET_NAME)

    Call ValidateLHSHeaders(wksSource, arrColumns, arrHeaders)

    lngLastRow = wksSource.Cells(wksSource.Rows.Count, 1).End(xlUp).Row
    If lngLastRow < 2 Then
        lngDataRows = 0
    Else
        lngDataRows = lngLastRow - 1
    End If

    ReDim arrOutput(1 To lngDataRows + 1, 1 To 14)

    For lngColIndex = LBound(arrHeaders) To UBound(arrHeaders)
        arrOutput(1, lngColIndex + 1) = CStr(arrHeaders(lngColIndex))
    Next lngColIndex

    If lngDataRows > 0 Then
        For lngColIndex = LBound(arrColumns) To UBound(arrColumns)
            arrColumn = wksSource.Range( _
                wksSource.Cells(2, CLng(arrColumns(lngColIndex))), _
                wksSource.Cells(lngLastRow, CLng(arrColumns(lngColIndex)))).Value2

            lngOutputColumn = lngColIndex + 1

            If lngDataRows = 1 Then
                If IsLHSDateOutputColumn(lngOutputColumn) Then
                    arrOutput(2, lngOutputColumn) = NormalizeLHSDateValue(arrColumn)
                ElseIf IsLHSDateTimeOutputColumn(lngOutputColumn) Then
                    arrOutput(2, lngOutputColumn) = NormalizeLHSDateTimeValue(arrColumn)
                Else
                    arrOutput(2, lngOutputColumn) = arrColumn
                End If
            Else
                For lngRow = 1 To lngDataRows
                    If IsLHSDateOutputColumn(lngOutputColumn) Then
                        arrOutput(lngRow + 1, lngOutputColumn) = _
                            NormalizeLHSDateValue(arrColumn(lngRow, 1))
                    ElseIf IsLHSDateTimeOutputColumn(lngOutputColumn) Then
                        arrOutput(lngRow + 1, lngOutputColumn) = _
                            NormalizeLHSDateTimeValue(arrColumn(lngRow, 1))
                    Else
                        arrOutput(lngRow + 1, lngOutputColumn) = arrColumn(lngRow, 1)
                    End If
                Next lngRow
            End If
        Next lngColIndex
    End If

    LoadLHSInput = arrOutput

ExitPoint:
    Set wksSource = Nothing

    If blnOpenedByMacro Then
        Call CloseWorkbookSafe(wkbSource, False)
    End If

    Set wkbSource = Nothing

    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strFullPath", strFullPath)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    lngOutputColumn - column position in the Input LHS output array
' Returns:       Boolean - True for LHS columns containing dates
' Description:   Identifies all LHS dates that must remain DD/MM/YYYY.
'-------------------------------------------------------------------------------
Private Function IsLHSDateOutputColumn(ByVal lngOutputColumn As Long) As Boolean
    Const METHOD_NAME As String = "IsLHSDateOutputColumn"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Select Case lngOutputColumn
        Case 3, 4, 7, 11
            IsLHSDateOutputColumn = True
        Case Else
            IsLHSDateOutputColumn = False
    End Select

ExitPoint:
    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "lngOutputColumn", lngOutputColumn)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    lngOutputColumn - column position in the Input LHS output array
' Returns:       Boolean - True for the Generation Date and Time column
' Description:   Identifies the LHS date-time column that must retain its time.
'-------------------------------------------------------------------------------
Private Function IsLHSDateTimeOutputColumn(ByVal lngOutputColumn As Long) As Boolean
    Const METHOD_NAME As String = "IsLHSDateTimeOutputColumn"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    IsLHSDateTimeOutputColumn = (lngOutputColumn = 8)

ExitPoint:
    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "lngOutputColumn", lngOutputColumn)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    vValue - source LHS Generation Date and Time value
' Returns:       Variant - Excel date-time serial or unchanged source value
' Description:   Parses DD/MM/YYYY with optional HH:MM[:SS] without locale CDate.
'-------------------------------------------------------------------------------
Private Function NormalizeLHSDateTimeValue(ByVal vValue As Variant) As Variant
    Const METHOD_NAME As String = "NormalizeLHSDateTimeValue"
    Dim arrDateParts As Variant
    Dim arrTimeParts As Variant
    Dim datParsed As Date
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngDay As Long
    Dim lngHour As Long
    Dim lngMinute As Long
    Dim lngMonth As Long
    Dim lngSecond As Long
    Dim lngYear As Long
    Dim strDatePart As String
    Dim strTimePart As String
    Dim strValue As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(vValue) Or IsEmpty(vValue) Then
        NormalizeLHSDateTimeValue = vValue
        GoTo ExitPoint
    End If

    If VarType(vValue) = vbDate Then
        NormalizeLHSDateTimeValue = CDbl(CDate(vValue))
        GoTo ExitPoint
    End If

    If IsNumeric(vValue) Then
        NormalizeLHSDateTimeValue = CDbl(vValue)
        GoTo ExitPoint
    End If

    strValue = Trim$(CStr(vValue))
    If Len(strValue) = 0 Then
        NormalizeLHSDateTimeValue = vbNullString
        GoTo ExitPoint
    End If

    If Len(strValue) < 10 Then
        NormalizeLHSDateTimeValue = vValue
        GoTo ExitPoint
    End If

    strDatePart = Left$(strValue, 10)
    strTimePart = Trim$(Mid$(strValue, 11))
    arrDateParts = Split(strDatePart, "/")

    If UBound(arrDateParts) - LBound(arrDateParts) <> 2 Then
        NormalizeLHSDateTimeValue = vValue
        GoTo ExitPoint
    End If

    If Not IsNumeric(arrDateParts(0)) Or Not IsNumeric(arrDateParts(1)) Or _
       Not IsNumeric(arrDateParts(2)) Then
        NormalizeLHSDateTimeValue = vValue
        GoTo ExitPoint
    End If

    lngDay = CLng(arrDateParts(0))
    lngMonth = CLng(arrDateParts(1))
    lngYear = CLng(arrDateParts(2))

    If lngDay < 1 Or lngDay > 31 Or lngMonth < 1 Or lngMonth > 12 Or _
       lngYear < 1900 Or lngYear > 9999 Then
        NormalizeLHSDateTimeValue = vValue
        GoTo ExitPoint
    End If

    datParsed = DateSerial(lngYear, lngMonth, lngDay)
    If Day(datParsed) <> lngDay Or Month(datParsed) <> lngMonth Or _
       Year(datParsed) <> lngYear Then
        NormalizeLHSDateTimeValue = vValue
        GoTo ExitPoint
    End If

    If Len(strTimePart) > 0 Then
        arrTimeParts = Split(strTimePart, ":")

        If UBound(arrTimeParts) < 1 Or UBound(arrTimeParts) > 2 Then
            NormalizeLHSDateTimeValue = vValue
            GoTo ExitPoint
        End If

        If Not IsNumeric(arrTimeParts(0)) Or Not IsNumeric(arrTimeParts(1)) Then
            NormalizeLHSDateTimeValue = vValue
            GoTo ExitPoint
        End If

        lngHour = CLng(arrTimeParts(0))
        lngMinute = CLng(arrTimeParts(1))

        If UBound(arrTimeParts) = 2 Then
            If Not IsNumeric(arrTimeParts(2)) Then
                NormalizeLHSDateTimeValue = vValue
                GoTo ExitPoint
            End If
            lngSecond = CLng(arrTimeParts(2))
        End If

        If lngHour < 0 Or lngHour > 23 Or lngMinute < 0 Or lngMinute > 59 Or _
           lngSecond < 0 Or lngSecond > 59 Then
            NormalizeLHSDateTimeValue = vValue
            GoTo ExitPoint
        End If

        datParsed = datParsed + TimeSerial(lngHour, lngMinute, lngSecond)
    End If

    NormalizeLHSDateTimeValue = CDbl(datParsed)

ExitPoint:
    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    vValue - source LHS date value
' Returns:       Variant - Excel date serial or unchanged non-date value
' Description:   Parses explicit DD/MM/YYYY text without locale-dependent CDate.
'-------------------------------------------------------------------------------
Private Function NormalizeLHSDateValue(ByVal vValue As Variant) As Variant
    Const METHOD_NAME As String = "NormalizeLHSDateValue"
    Dim arrParts As Variant
    Dim datParsed As Date
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngDay As Long
    Dim lngMonth As Long
    Dim lngYear As Long
    Dim strValue As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(vValue) Or IsEmpty(vValue) Then
        NormalizeLHSDateValue = vValue
        GoTo ExitPoint
    End If

    If VarType(vValue) = vbDate Then
        NormalizeLHSDateValue = CDbl(CDate(vValue))
        GoTo ExitPoint
    End If

    If IsNumeric(vValue) Then
        NormalizeLHSDateValue = CDbl(vValue)
        GoTo ExitPoint
    End If

    strValue = Trim$(CStr(vValue))
    If Len(strValue) = 0 Then
        NormalizeLHSDateValue = vbNullString
        GoTo ExitPoint
    End If

    arrParts = Split(strValue, "/")

    If UBound(arrParts) - LBound(arrParts) = 2 Then
        If Len(CStr(arrParts(0))) <= 2 And Len(CStr(arrParts(1))) <= 2 And _
           Len(CStr(arrParts(2))) = 4 Then

            If IsNumeric(arrParts(0)) And IsNumeric(arrParts(1)) And IsNumeric(arrParts(2)) Then
                lngDay = CLng(arrParts(0))
                lngMonth = CLng(arrParts(1))
                lngYear = CLng(arrParts(2))

                If lngDay >= 1 And lngDay <= 31 And lngMonth >= 1 And lngMonth <= 12 And _
                   lngYear >= 1900 And lngYear <= 9999 Then

                    datParsed = DateSerial(lngYear, lngMonth, lngDay)

                    If Day(datParsed) = lngDay And Month(datParsed) = lngMonth And _
                       Year(datParsed) = lngYear Then
                        NormalizeLHSDateValue = CDbl(datParsed)
                        GoTo ExitPoint
                    End If
                End If
            End If
        End If
    End If

    ' Preserve unexpected non-empty source content instead of guessing its meaning.
    NormalizeLHSDateValue = vValue

ExitPoint:
    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    wksSource - LHS sheet; arrColumns/arrHeaders - required layout
' Returns:       None
' Description:   Validates row-1 headers before fixed-column extraction.
'-------------------------------------------------------------------------------
Private Sub ValidateLHSHeaders( _
    ByVal wksSource As Excel.Worksheet, _
    ByVal arrColumns As Variant, _
    ByVal arrHeaders As Variant)

    Const METHOD_NAME As String = "ValidateLHSHeaders"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngIndex As Long
    Dim strActual As String
    Dim strExpected As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngIndex = LBound(arrColumns) To UBound(arrColumns)
        strActual = Trim$(CStr(wksSource.Cells(1, CLng(arrColumns(lngIndex))).Value2))
        strExpected = CStr(arrHeaders(lngIndex))

        If StrComp(strActual, strExpected, vbTextCompare) <> 0 Then
            Err.Raise ERR_LHS, METHOD_NAME, _
                      "Unexpected LHS header in column " & _
                      wksSource.Cells(1, CLng(arrColumns(lngIndex))).Address(False, False) & _
                      ". Expected '" & strExpected & "', found '" & strActual & "'."
        End If
    Next lngIndex

ExitPoint:
    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Sub
