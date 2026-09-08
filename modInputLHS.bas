Option Explicit

Private Const CLASS_NAME As String = "modInputLHS"
Private Const LHS_SHEET_NAME As String = "FAMOO-Full Inventory IFRS"
Private Const ERR_LHS As Long = vbObjectError + 6400

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strFullPath - selected LHS input; lngDataRows - returned row count
' Returns:       Variant - 2D output array including header row
' Description:   Extracts the 13 required LHS columns in the requested order.
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
    Dim lngRow As Long
    Dim wkbSource As Excel.Workbook
    Dim wksSource As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrColumns = Array(1&, 9&, 115&, 114&, 116&, 130&, 134&, 71&, 68&, 66&, 35&, 15&, 16&)
    arrHeaders = Array( _
        "FUND CODE", _
        "EXTERNAL VALUE CODE", _
        "LAST COUPON DATE", _
        "NEXT COUPON DATE", _
        "NBR OF DAYS SINCE LAST COUPON", _
        "GTI NAME", _
        "ACCOUNTING DATE_1", _
        "BOND CALC METHOD/IRS-CDS", _
        "Coupon frequency/IRS-CDS", _
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

    ReDim arrOutput(1 To lngDataRows + 1, 1 To 13)

    For lngColIndex = LBound(arrHeaders) To UBound(arrHeaders)
        arrOutput(1, lngColIndex + 1) = CStr(arrHeaders(lngColIndex))
    Next lngColIndex

    If lngDataRows > 0 Then
        For lngColIndex = LBound(arrColumns) To UBound(arrColumns)
            arrColumn = wksSource.Range( _
                wksSource.Cells(2, CLng(arrColumns(lngColIndex))), _
                wksSource.Cells(lngLastRow, CLng(arrColumns(lngColIndex)))).Value2

            If lngDataRows = 1 Then
                arrOutput(2, lngColIndex + 1) = arrColumn
            Else
                For lngRow = 1 To lngDataRows
                    arrOutput(lngRow + 1, lngColIndex + 1) = arrColumn(lngRow, 1)
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
