Option Explicit

Private Const CLASS_NAME As String = "modInputNeolink"
Private Const ERR_NEOLINK As Long = vbObjectError + 6500

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strFullPath - selected Neolink file; dictFund - Fondsliste lookup;
'                lngIncludedRows/lngUnmatchedRows - returned statistics
' Returns:       Variant - filtered 2D output array including header row
' Description:   Keeps only rows whose last 11 account characters start with 103.
'-------------------------------------------------------------------------------
Public Function LoadNeolinkInput( _
    ByVal strFullPath As String, _
    ByVal dictFund As Object, _
    ByRef lngIncludedRows As Long, _
    ByRef lngUnmatchedRows As Long) As Variant

    Const METHOD_NAME As String = "LoadNeolinkInput"
    Dim arrHeaders As Variant
    Dim arrKeys() As String
    Dim arrOutput() As Variant
    Dim arrSource As Variant
    Dim blnOpenedByMacro As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngOutputRow As Long
    Dim lngRow As Long
    Dim lngSourceCol As Long
    Dim lngLastRow As Long
    Dim strAccount As String
    Dim strFundCode As String
    Dim strKey As String
    Dim wkbSource As Excel.Workbook
    Dim wksSource As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If dictFund Is Nothing Then
        Err.Raise ERR_NEOLINK, METHOD_NAME, "Fondsliste lookup is not initialized."
    End If

    arrHeaders = Array( _
        "Fund Code", _
        "GL:Securities account", _
        "GL:Type code", _
        "GL:Type", _
        "CA:ISIN code", _
        "CM:Value date", _
        "CM:Payment date")

    Set wkbSource = GetOrOpenWorkbookReadOnly(strFullPath, blnOpenedByMacro)
    Set wksSource = GetFirstWorksheet(wkbSource)

    Call ValidateNeolinkHeaders(wksSource)

    lngLastRow = wksSource.Cells(wksSource.Rows.Count, 1).End(xlUp).Row

    If lngLastRow < 2 Then
        lngIncludedRows = 0
        lngUnmatchedRows = 0
        ReDim arrOutput(1 To 1, 1 To 7)

        For lngSourceCol = LBound(arrHeaders) To UBound(arrHeaders)
            arrOutput(1, lngSourceCol + 1) = CStr(arrHeaders(lngSourceCol))
        Next lngSourceCol

        LoadNeolinkInput = arrOutput
        GoTo ExitPoint
    End If

    arrSource = wksSource.Range(wksSource.Cells(1, 1), wksSource.Cells(lngLastRow, 6)).Value2
    ReDim arrKeys(1 To UBound(arrSource, 1))

    lngIncludedRows = 0
    lngUnmatchedRows = 0

    For lngRow = 2 To UBound(arrSource, 1)
        strAccount = Replace(Trim$(CStr(arrSource(lngRow, 1))), " ", vbNullString)

        If Len(strAccount) >= 11 Then
            strKey = UCase$(Right$(strAccount, 11))

            If Left$(strKey, 3) = "103" Then
                arrKeys(lngRow) = strKey
                lngIncludedRows = lngIncludedRows + 1
            End If
        End If
    Next lngRow

    ReDim arrOutput(1 To lngIncludedRows + 1, 1 To 7)

    For lngSourceCol = LBound(arrHeaders) To UBound(arrHeaders)
        arrOutput(1, lngSourceCol + 1) = CStr(arrHeaders(lngSourceCol))
    Next lngSourceCol

    lngOutputRow = 2

    For lngRow = 2 To UBound(arrSource, 1)
        strKey = arrKeys(lngRow)

        If Len(strKey) > 0 Then
            strFundCode = vbNullString

            If dictFund.Exists(strKey) Then
                strFundCode = CStr(dictFund(strKey))
            Else
                lngUnmatchedRows = lngUnmatchedRows + 1
            End If

            arrOutput(lngOutputRow, 1) = strFundCode

            For lngSourceCol = 1 To 6
                arrOutput(lngOutputRow, lngSourceCol + 1) = arrSource(lngRow, lngSourceCol)
            Next lngSourceCol

            lngOutputRow = lngOutputRow + 1
        End If
    Next lngRow

    LoadNeolinkInput = arrOutput

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
' Parameters:    wksSource - first worksheet of the Neolink file
' Returns:       None
' Description:   Validates the required A:F headers in row 1.
'-------------------------------------------------------------------------------
Private Sub ValidateNeolinkHeaders(ByVal wksSource As Excel.Worksheet)
    Const METHOD_NAME As String = "ValidateNeolinkHeaders"
    Dim arrExpected As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngIndex As Long
    Dim strActual As String
    Dim strExpected As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrExpected = Array( _
        "GL:Securities account", _
        "GL:Type code", _
        "GL:Type", _
        "CA:ISIN code", _
        "CM:Value date", _
        "CM:Payment date")

    For lngIndex = LBound(arrExpected) To UBound(arrExpected)
        strActual = Trim$(CStr(wksSource.Cells(1, lngIndex + 1).Value2))
        strExpected = CStr(arrExpected(lngIndex))

        If StrComp(strActual, strExpected, vbTextCompare) <> 0 Then
            Err.Raise ERR_NEOLINK, METHOD_NAME, _
                      "Unexpected Neolink header in column " & _
                      wksSource.Cells(1, lngIndex + 1).Address(False, False) & _
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
