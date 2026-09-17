Option Explicit

Private Const CLASS_NAME As String = "modInputLHS"
Private Const LHS_SHEET_NAME As String = "FAMOO-Full Inventory IFRS"
Private Const LHS_QUANTITY_COLUMN As Long = 10
Private Const ERR_LHS As Long = vbObjectError + 6400

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strFullPath - selected LHS input; lngDataRows - returned row count
' Returns:       Variant - 2D output array including header row
' Description:   Extracts the required LHS columns plus schedule/accrual diagnostics.
'-------------------------------------------------------------------------------
Public Function LoadLHSInput( _
    ByVal strFullPath As String, ByRef lngDataRows As Long) As Variant

    Const METHOD_NAME As String = "LoadLHSInput"
    Dim arrColumn As Variant
    Dim arrColumns As Variant
    Dim arrHeaders As Variant
    Dim arrOptionalHeaders As Variant
    Dim arrOutput() As Variant
    Dim blnOpenedByMacro As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColIndex As Long
    Dim lngLastRow As Long
    Dim lngOutputColumn As Long
    Dim lngOptionalColumn As Long
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

    ' Optional diagnostic fields are resolved by row-1 header name so their physical
    ' source column can move without changing the macro. Missing optional fields are
    ' left blank and handled as REVIEW downstream. QUANTITY is fixed in source J.
    arrOptionalHeaders = Array( _
        "MATURITY DATE", _
        "FIRST COUPON DATE", _
        "INTEREST RATE TYPE", _
        "SECURITY INDEXATION MODE")

    Set wkbSource = GetOrOpenWorkbookReadOnly(strFullPath, blnOpenedByMacro)
    Set wksSource = GetWorksheetByName(wkbSource, LHS_SHEET_NAME)

    Call ValidateLHSHeaders(wksSource, arrColumns, arrHeaders)
    Call ValidateFixedHeader(wksSource, LHS_QUANTITY_COLUMN, "QUANTITY")

    lngLastRow = wksSource.Cells(wksSource.Rows.Count, 1).End(xlUp).Row
    If lngLastRow < 2 Then
        lngDataRows = 0
    Else
        lngDataRows = lngLastRow - 1
    End If

    ReDim arrOutput(1 To lngDataRows + 1, 1 To 19)

    For lngColIndex = LBound(arrHeaders) To UBound(arrHeaders)
        arrOutput(1, lngColIndex + 1) = CStr(arrHeaders(lngColIndex))
    Next lngColIndex

    For lngColIndex = LBound(arrOptionalHeaders) To UBound(arrOptionalHeaders)
        arrOutput(1, 15 + lngColIndex) = CStr(arrOptionalHeaders(lngColIndex))
    Next lngColIndex
    arrOutput(1, 19) = "QUANTITY"

    If lngDataRows > 0 Then
        For lngColIndex = LBound(arrColumns) To UBound(arrColumns)
            arrColumn = wksSource.Range( _
                wksSource.Cells(2, CLng(arrColumns(lngColIndex))), _
                wksSource.Cells(lngLastRow, CLng(arrColumns(lngColIndex)))).Value2

            lngOutputColumn = lngColIndex + 1

            If lngDataRows = 1 Then
                If IsLHSDateOutputColumn(lngOutputColumn) Then
                    arrOutput(2, lngOutputColumn) = NormalizeLHSDateValue(arrColumn)
                Else
                    arrOutput(2, lngOutputColumn) = arrColumn
                End If
            Else
                For lngRow = 1 To lngDataRows
                    If IsLHSDateOutputColumn(lngOutputColumn) Then
                        arrOutput(lngRow + 1, lngOutputColumn) = _
                            NormalizeLHSDateValue(arrColumn(lngRow, 1))
                    Else
                        arrOutput(lngRow + 1, lngOutputColumn) = arrColumn(lngRow, 1)
                    End If
                Next lngRow
            End If
        Next lngColIndex

        For lngColIndex = LBound(arrOptionalHeaders) To UBound(arrOptionalHeaders)
            lngOptionalColumn = FindOptionalHeaderColumn( _
                wksSource, CStr(arrOptionalHeaders(lngColIndex)))

            If lngOptionalColumn > 0 Then
                arrColumn = wksSource.Range( _
                    wksSource.Cells(2, lngOptionalColumn), _
                    wksSource.Cells(lngLastRow, lngOptionalColumn)).Value2

                lngOutputColumn = 15 + lngColIndex

                If lngDataRows = 1 Then
                    If IsLHSDateOutputColumn(lngOutputColumn) Then
                        arrOutput(2, lngOutputColumn) = NormalizeLHSDateValue(arrColumn)
                    Else
                        arrOutput(2, lngOutputColumn) = arrColumn
                    End If
                Else
                    For lngRow = 1 To lngDataRows
                        If IsLHSDateOutputColumn(lngOutputColumn) Then
                            arrOutput(lngRow + 1, lngOutputColumn) = _
                                NormalizeLHSDateValue(arrColumn(lngRow, 1))
                        Else
                            arrOutput(lngRow + 1, lngOutputColumn) = arrColumn(lngRow, 1)
                        End If
                    Next lngRow
                End If
            End If
        Next lngColIndex


        arrColumn = wksSource.Range( _
            wksSource.Cells(2, LHS_QUANTITY_COLUMN), _
            wksSource.Cells(lngLastRow, LHS_QUANTITY_COLUMN)).Value2

        If lngDataRows = 1 Then
            arrOutput(2, 19) = arrColumn
        Else
            For lngRow = 1 To lngDataRows
                arrOutput(lngRow + 1, 19) = arrColumn(lngRow, 1)
            Next lngRow
        End If
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
        Case 3, 4, 7, 8, 11, 15, 16
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
' Parameters:    vValue - source LHS date value
' Returns:       Variant - Excel date serial or unchanged non-date value
' Description:   Parses DD/MM/YYYY and strips any time component from date-only fields.
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
        NormalizeLHSDateValue = Int(CDbl(CDate(vValue)))
        GoTo ExitPoint
    End If

    If IsNumeric(vValue) Then
        NormalizeLHSDateValue = Int(CDbl(vValue))
        GoTo ExitPoint
    End If

    strValue = Trim$(CStr(vValue))
    If Len(strValue) >= 10 Then strValue = Left$(strValue, 10)
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
' Creation date: 2026-09-17
' Parameters:    wksSource - LHS sheet; strHeader - optional row-1 header
' Returns:       Long - matching source column, or 0 when the header is absent
' Description:   Resolves optional diagnostic fields without depending on a fixed
'                source column letter. Comparison is trimmed and case-insensitive.
'-------------------------------------------------------------------------------
Private Function FindOptionalHeaderColumn( _
    ByVal wksSource As Excel.Worksheet, ByVal strHeader As String) As Long

    Const METHOD_NAME As String = "FindOptionalHeaderColumn"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngLastColumn As Long
    Dim strActual As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastColumn = wksSource.Cells(1, wksSource.Columns.Count).End(xlToLeft).Column

    For lngColumn = 1 To lngLastColumn
        strActual = Trim$(CStr(wksSource.Cells(1, lngColumn).Value2))
        If StrComp(strActual, Trim$(strHeader), vbTextCompare) = 0 Then
            FindOptionalHeaderColumn = lngColumn
            Exit For
        End If
    Next lngColumn

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
        "strHeader", strHeader)
    GoTo ExitPoint
End Function

Private Sub ValidateFixedHeader( _
    ByVal wksSource As Excel.Worksheet, ByVal lngColumn As Long, _
    ByVal strExpected As String)

    Const METHOD_NAME As String = "ValidateFixedHeader"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strActual As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strActual = Trim$(CStr(wksSource.Cells(1, lngColumn).Value2))

    If StrComp(strActual, strExpected, vbTextCompare) <> 0 Then
        Err.Raise ERR_LHS, METHOD_NAME, _
            "Unexpected LHS header in column " & _
            wksSource.Cells(1, lngColumn).Address(False, False) & _
            ". Expected '" & strExpected & "', found '" & strActual & "'."
    End If

ExitPoint:
    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "lngColumn", lngColumn)
    GoTo ExitPoint
End Sub

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

