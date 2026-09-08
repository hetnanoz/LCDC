Option Explicit

Private Const CLASS_NAME As String = "modOutput"
Private Const OUTPUT_BASE_NAME As String = "Last_Coupon_Date_Checker_"
Private Const ERR_OUTPUT As Long = vbObjectError + 6600
Private Const LOOKUP_SEPARATOR As String = "|"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    arrLHS - LHS output array; arrNeolink - Neolink output array
' Returns:       Excel.Workbook - newly created output workbook
' Description:   Creates Input LHS, Input Neolink and Output_All sheets.
'-------------------------------------------------------------------------------
Public Function CreateOutputWorkbook( _
    ByVal arrLHS As Variant, ByVal arrNeolink As Variant) As Excel.Workbook

    Const METHOD_NAME As String = "CreateOutputWorkbook"
    Dim arrOutputAll As Variant
    Dim blnCompleted As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim wkbOutput As Excel.Workbook
    Dim wksLHS As Excel.Worksheet
    Dim wksNeolink As Excel.Worksheet
    Dim wksOutputAll As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set wkbOutput = Application.Workbooks.Add(xlWBATWorksheet)
    Set wksLHS = wkbOutput.Worksheets(1)
    wksLHS.Name = "Input LHS"

    Set wksNeolink = wkbOutput.Worksheets.Add(After:=wksLHS)
    wksNeolink.Name = "Input Neolink"

    Set wksOutputAll = wkbOutput.Worksheets.Add(After:=wksNeolink)
    wksOutputAll.Name = "Output_All"

    ' Text identifiers must not be converted to numbers/scientific notation.
    wksLHS.Columns("A:B").NumberFormat = "@"
    wksNeolink.Columns("A:B").NumberFormat = "@"
    wksOutputAll.Columns("A:B").NumberFormat = "@"

    Call WriteArrayToWorksheet(wksLHS, arrLHS)
    Call WriteArrayToWorksheet(wksNeolink, arrNeolink)

    arrOutputAll = BuildOutputAllArray(arrLHS, arrNeolink)
    Call WriteArrayToWorksheet(wksOutputAll, arrOutputAll)

    ' Force all known date columns to the required DD/MM/YYYY display.
    wksLHS.Columns("C:D").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("G:G").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("J:J").NumberFormat = "dd/mm/yyyy"
    wksNeolink.Columns("F:G").NumberFormat = "dd/mm/yyyy"
    wksOutputAll.Columns("C:F").NumberFormat = "dd/mm/yyyy"
    wksOutputAll.Columns("H:H").NumberFormat = "0"

    Call ApplyMatchFormatting(wksOutputAll)

    wksLHS.Rows(1).Font.Bold = True
    wksNeolink.Rows(1).Font.Bold = True
    wksOutputAll.Rows(1).Font.Bold = True

    wksLHS.UsedRange.Columns.AutoFit
    wksNeolink.UsedRange.Columns.AutoFit
    wksOutputAll.UsedRange.Columns.AutoFit

    Set CreateOutputWorkbook = wkbOutput
    blnCompleted = True

ExitPoint:
    Set wksLHS = Nothing
    Set wksNeolink = Nothing
    Set wksOutputAll = Nothing

    If Not blnCompleted Then
        Call CloseWorkbookSafe(wkbOutput, False)
    End If

    Set wkbOutput = Nothing

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
' Parameters:    arrLHS - complete Input LHS array; arrNeolink - Neolink array
' Returns:       Variant - eight-column Output_All array including headers
' Description:   Adds latest INTR Neolink Payment Date for each Fund + ISIN and
'                compares it with the LHS Last Coupon Date.
'-------------------------------------------------------------------------------
Private Function BuildOutputAllArray( _
    ByVal arrLHS As Variant, ByVal arrNeolink As Variant) As Variant

    Const METHOD_NAME As String = "BuildOutputAllArray"
    Dim arrResult() As Variant
    Dim blnHasLastCouponDate As Boolean
    Dim dictLatestPayment As Object
    Dim dblLastCouponDate As Double
    Dim dblPaymentDate As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLastCouponDay As Long
    Dim lngPaymentDay As Long
    Dim lngRow As Long
    Dim lngRows As Long
    Dim strKey As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictLatestPayment = BuildLatestNeolinkPaymentLookup(arrNeolink)
    lngRows = UBound(arrLHS, 1)

    ReDim arrResult(1 To lngRows, 1 To 8)

    arrResult(1, 1) = arrLHS(1, 1)
    arrResult(1, 2) = arrLHS(1, 2)
    arrResult(1, 3) = arrLHS(1, 3)
    arrResult(1, 4) = arrLHS(1, 4)
    arrResult(1, 5) = arrLHS(1, 7)
    arrResult(1, 6) = "Neolink Payment Date"
    arrResult(1, 7) = "Last Coupon Date Match"
    arrResult(1, 8) = "Difference Days"

    For lngRow = 2 To lngRows
        arrResult(lngRow, 1) = arrLHS(lngRow, 1)
        arrResult(lngRow, 2) = arrLHS(lngRow, 2)
        arrResult(lngRow, 3) = arrLHS(lngRow, 3)
        arrResult(lngRow, 4) = arrLHS(lngRow, 4)
        arrResult(lngRow, 5) = arrLHS(lngRow, 7)

        strKey = BuildFundIsinKey(arrLHS(lngRow, 1), arrLHS(lngRow, 2))

        If Len(strKey) = 0 Or Not dictLatestPayment.Exists(strKey) Then
            arrResult(lngRow, 7) = "NOT FOUND"
        Else
            dblPaymentDate = CDbl(dictLatestPayment(strKey))
            arrResult(lngRow, 6) = dblPaymentDate

            blnHasLastCouponDate = TryGetExcelDateSerial( _
                arrLHS(lngRow, 3), dblLastCouponDate)

            If blnHasLastCouponDate Then
                lngLastCouponDay = CLng(Int(dblLastCouponDate))
                lngPaymentDay = CLng(Int(dblPaymentDate))

                If lngLastCouponDay = lngPaymentDay Then
                    arrResult(lngRow, 7) = "TRUE"
                Else
                    arrResult(lngRow, 7) = "FALSE"
                End If

                ' Positive value means Neolink Payment Date is later than LHS Last Coupon Date.
                arrResult(lngRow, 8) = lngPaymentDay - lngLastCouponDay
            Else
                arrResult(lngRow, 7) = "NO LHS DATE"
            End If
        End If
    Next lngRow

    BuildOutputAllArray = arrResult

ExitPoint:
    Set dictLatestPayment = Nothing

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
' Parameters:    arrNeolink - complete Input Neolink array including headers
' Returns:       Object - Fund + ISIN -> latest valid INTR Payment Date
' Description:   Uses only GL:Type code = INTR and keeps the greatest Payment Date.
'-------------------------------------------------------------------------------
Private Function BuildLatestNeolinkPaymentLookup(ByVal arrNeolink As Variant) As Object
    Const METHOD_NAME As String = "BuildLatestNeolinkPaymentLookup"
    Dim dictLatestPayment As Object
    Dim dblPaymentDate As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngRow As Long
    Dim strKey As String
    Dim strTypeCode As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictLatestPayment = CreateObject("Scripting.Dictionary")
    dictLatestPayment.CompareMode = vbTextCompare

    For lngRow = 2 To UBound(arrNeolink, 1)
        strTypeCode = UCase$(Trim$(CStr(arrNeolink(lngRow, 3))))

        If strTypeCode = "INTR" Then
            strKey = BuildFundIsinKey(arrNeolink(lngRow, 1), arrNeolink(lngRow, 5))

            If Len(strKey) > 0 Then
                If TryGetExcelDateSerial(arrNeolink(lngRow, 7), dblPaymentDate) Then
                    If dictLatestPayment.Exists(strKey) Then
                        If dblPaymentDate > CDbl(dictLatestPayment(strKey)) Then
                            dictLatestPayment(strKey) = dblPaymentDate
                        End If
                    Else
                        dictLatestPayment.Add strKey, dblPaymentDate
                    End If
                End If
            End If
        End If
    Next lngRow

    Set BuildLatestNeolinkPaymentLookup = dictLatestPayment

ExitPoint:
    If errNumber <> 0 Then
        Set dictLatestPayment = Nothing
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
' Parameters:    vFundCode - Fund Code; vIsin - External Value Code / CA:ISIN code
' Returns:       String - normalized compound lookup key or empty string
' Description:   Matches Fund + ISIN case-insensitively and ignores spaces.
'-------------------------------------------------------------------------------
Private Function BuildFundIsinKey(ByVal vFundCode As Variant, ByVal vIsin As Variant) As String
    Const METHOD_NAME As String = "BuildFundIsinKey"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strFundCode As String
    Dim strIsin As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(vFundCode) Or IsError(vIsin) Then GoTo ExitPoint

    strFundCode = UCase$(Replace(Trim$(CStr(vFundCode)), " ", vbNullString))
    strIsin = UCase$(Replace(Trim$(CStr(vIsin)), " ", vbNullString))

    If Len(strFundCode) > 0 And Len(strIsin) > 0 Then
        BuildFundIsinKey = strFundCode & LOOKUP_SEPARATOR & strIsin
    End If

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
' Parameters:    vValue - candidate date; dblDateSerial - returned Excel date serial
' Returns:       Boolean - True when the value can be used as a date
' Description:   Accepts normalized Excel serials, Date values and DD/MM/YYYY text.
'-------------------------------------------------------------------------------
Private Function TryGetExcelDateSerial( _
    ByVal vValue As Variant, ByRef dblDateSerial As Double) As Boolean

    Const METHOD_NAME As String = "TryGetExcelDateSerial"
    Dim arrParts As Variant
    Dim datParsed As Date
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngDay As Long
    Dim lngMonth As Long
    Dim lngYear As Long
    Dim strValue As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    dblDateSerial = 0

    If IsError(vValue) Or IsEmpty(vValue) Then GoTo ExitPoint

    If VarType(vValue) = vbDate Then
        dblDateSerial = CDbl(CDate(vValue))
        TryGetExcelDateSerial = True
        GoTo ExitPoint
    End If

    If IsNumeric(vValue) Then
        If CDbl(vValue) > 0 Then
            dblDateSerial = CDbl(vValue)
            TryGetExcelDateSerial = True
        End If
        GoTo ExitPoint
    End If

    strValue = Trim$(CStr(vValue))
    If Len(strValue) = 0 Then GoTo ExitPoint

    arrParts = Split(strValue, "/")

    If UBound(arrParts) - LBound(arrParts) = 2 Then
        If IsNumeric(arrParts(0)) And IsNumeric(arrParts(1)) And IsNumeric(arrParts(2)) Then
            lngDay = CLng(arrParts(0))
            lngMonth = CLng(arrParts(1))
            lngYear = CLng(arrParts(2))

            If lngDay >= 1 And lngDay <= 31 And lngMonth >= 1 And lngMonth <= 12 And _
               lngYear >= 1900 And lngYear <= 9999 Then

                datParsed = DateSerial(lngYear, lngMonth, lngDay)

                If Day(datParsed) = lngDay And Month(datParsed) = lngMonth And _
                   Year(datParsed) = lngYear Then
                    dblDateSerial = CDbl(datParsed)
                    TryGetExcelDateSerial = True
                End If
            End If
        End If
    End If

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
' Parameters:    wksOutputAll - Output_All worksheet
' Returns:       None
' Description:   Colors TRUE matches green and FALSE matches red in column G.
'-------------------------------------------------------------------------------
Private Sub ApplyMatchFormatting(ByVal wksOutputAll As Excel.Worksheet)
    Const METHOD_NAME As String = "ApplyMatchFormatting"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLastRow As Long
    Dim rngMatch As Excel.Range

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastRow = wksOutputAll.Cells(wksOutputAll.Rows.Count, 1).End(xlUp).Row
    If lngLastRow < 2 Then GoTo ExitPoint

    Set rngMatch = wksOutputAll.Range("G2:G" & CStr(lngLastRow))
    rngMatch.FormatConditions.Delete

    With rngMatch.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$G2=""TRUE""")
        .Interior.Color = RGB(198, 239, 206)
    End With

    With rngMatch.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$G2=""FALSE""")
        .Interior.Color = RGB(255, 199, 206)
    End With

ExitPoint:
    Set rngMatch = Nothing

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

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strSaveFolder - configured save_path folder
' Returns:       String - unique full .xlsx output path
' Description:   Uses timestamp yyyymmdd_hhnnss and avoids accidental overwrite.
'-------------------------------------------------------------------------------
Public Function BuildOutputPath(ByVal strSaveFolder As String) As String
    Const METHOD_NAME As String = "BuildOutputPath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngSuffix As Long
    Dim strCandidate As String
    Dim strFolder As String
    Dim strStem As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Call ValidateFolderExists(strSaveFolder, "Output")

    strFolder = Trim$(strSaveFolder)
    If Right$(strFolder, 1) <> Application.PathSeparator Then
        strFolder = strFolder & Application.PathSeparator
    End If

    strStem = OUTPUT_BASE_NAME & Format$(Now, "yyyymmdd_hhnnss")
    strCandidate = strFolder & strStem & ".xlsx"

    Do While Len(Dir$(strCandidate)) > 0
        lngSuffix = lngSuffix + 1
        strCandidate = strFolder & strStem & "_" & Format$(lngSuffix, "00") & ".xlsx"
    Loop

    BuildOutputPath = strCandidate

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
        "strSaveFolder", strSaveFolder)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    wkbOutput - output workbook; strFullPath - exact destination path
' Returns:       None
' Description:   Applies the Internal sensitivity label and saves as .xlsx.
'-------------------------------------------------------------------------------
Public Sub SaveOutputWorkbook( _
    ByVal wkbOutput As Excel.Workbook, ByVal strFullPath As String)

    Const METHOD_NAME As String = "SaveOutputWorkbook"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If wkbOutput Is Nothing Then
        Err.Raise ERR_OUTPUT, METHOD_NAME, "Output workbook is not initialized."
    End If

    Call ApplySensitivityLabel(wkbOutput, MIP_INTERNAL_LABEL_ID)

    wkbOutput.SaveAs _
        Filename:=strFullPath, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False

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
        "strFullPath", strFullPath)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    wkbTarget - output workbook; strLabelId - sensitivity label ID
' Returns:       None
' Description:   Applies Microsoft Information Protection Internal label.
'-------------------------------------------------------------------------------
Private Sub ApplySensitivityLabel( _
    ByVal wkbTarget As Excel.Workbook, _
    ByVal strLabelId As String)

    Const METHOD_NAME As String = "ApplySensitivityLabel"
    Dim errDescription As String
    Dim errNumber As Long
    Dim objLabelInfo As Object

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If wkbTarget Is Nothing Then
        Err.Raise ERR_OUTPUT, METHOD_NAME, "Target workbook is not initialized."
    End If

    Set objLabelInfo = wkbTarget.SensitivityLabel.CreateLabelInfo()

    With objLabelInfo
        .LabelId = strLabelId
        .SiteId = MIP_SITE_ID
        .AssignmentMethod = MIP_ASSIGNMENT_METHOD
        .ContentBits = MIP_CONTENT_BITS
        .IsEnabled = True
    End With

    Call wkbTarget.SensitivityLabel.SetLabel(objLabelInfo, objLabelInfo)

ExitPoint:
    Set objLabelInfo = Nothing

    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strLabelId", strLabelId)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    wksTarget - destination sheet; arrData - 2D source array
' Returns:       None
' Description:   Writes the complete data array to the worksheet in one operation.
'-------------------------------------------------------------------------------
Private Sub WriteArrayToWorksheet( _
    ByVal wksTarget As Excel.Worksheet, ByVal arrData As Variant)

    Const METHOD_NAME As String = "WriteArrayToWorksheet"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumns As Long
    Dim lngRows As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngRows = UBound(arrData, 1)
    lngColumns = UBound(arrData, 2)

    wksTarget.Range("A1").Resize(lngRows, lngColumns).Value2 = arrData

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
        "sheet", wksTarget.Name)
    GoTo ExitPoint
End Sub
