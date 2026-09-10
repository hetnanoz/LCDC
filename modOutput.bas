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
    wksOutputAll.Columns("A:C").NumberFormat = "@"

    Call WriteArrayToWorksheet(wksLHS, arrLHS)
    Call WriteArrayToWorksheet(wksNeolink, arrNeolink)

    arrOutputAll = BuildOutputAllArray(arrLHS, arrNeolink)
    Call WriteArrayToWorksheet(wksOutputAll, arrOutputAll)

    ' Force all known date columns to the required DD/MM/YYYY display.
    wksLHS.Columns("C:D").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("G:G").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("H:H").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("K:K").NumberFormat = "dd/mm/yyyy"
    wksNeolink.Columns("F:G").NumberFormat = "dd/mm/yyyy"
    wksOutputAll.Columns("D:D").NumberFormat = "dd/mm/yyyy"
    wksOutputAll.Columns("E:E").NumberFormat = "dd/mm/yyyy"
    wksOutputAll.Columns("H:K").NumberFormat = "dd/mm/yyyy"
    wksOutputAll.Columns("M:M").NumberFormat = "0"

    Call ApplyMatchFormatting(wksOutputAll)
    Call ApplyOutputAllDateHighlighting(wksOutputAll)
    Call ApplyWeekendDateHighlighting(wksOutputAll)
    Call ApplyValueDateMismatchFormatting(wksOutputAll)
    Call ApplyMappingTooltips(wksOutputAll)

    Call FormatWorksheetLayout(wksLHS)
    Call FormatWorksheetLayout(wksNeolink)
    Call FormatWorksheetLayout(wksOutputAll)

    Call FreezeTopRowOnAllSheets(wkbOutput)

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
' Returns:       Variant - thirteen-column Output_All array including headers
' Description:   Keeps only LHS rows with a non-empty Last Coupon Date. Adds LHS
'                generation/coupon fields plus Value Date and latest INTR Payment
'                Date from the same Neolink row, then compares Last Coupon Date.
'-------------------------------------------------------------------------------
Private Function BuildOutputAllArray( _
    ByVal arrLHS As Variant, ByVal arrNeolink As Variant) As Variant

    Const METHOD_NAME As String = "BuildOutputAllArray"
    Dim arrLookupValue As Variant
    Dim arrResult() As Variant
    Dim blnHasLastCouponDate As Boolean
    Dim dictLatestPayment As Object
    Dim dblLastCouponDate As Double
    Dim dblPaymentDate As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLastCouponDay As Long
    Dim lngOutputRow As Long
    Dim lngPaymentDay As Long
    Dim lngRow As Long
    Dim lngRows As Long
    Dim lngRowsWithLastCouponDate As Long
    Dim strKey As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictLatestPayment = BuildLatestNeolinkPaymentLookup(arrNeolink)
    lngRows = UBound(arrLHS, 1)

    For lngRow = 2 To lngRows
        If HasNonEmptyValue(arrLHS(lngRow, 3)) Then
            lngRowsWithLastCouponDate = lngRowsWithLastCouponDate + 1
        End If
    Next lngRow

    ReDim arrResult(1 To lngRowsWithLastCouponDate + 1, 1 To 13)

    arrResult(1, 1) = arrLHS(1, 1)
    arrResult(1, 2) = "Securities account"
    arrResult(1, 3) = arrLHS(1, 2)
    arrResult(1, 4) = arrLHS(1, 7)
    arrResult(1, 5) = arrLHS(1, 8)
    arrResult(1, 6) = arrLHS(1, 9)
    arrResult(1, 7) = arrLHS(1, 10)
    arrResult(1, 8) = "Neolink Value Date"
    arrResult(1, 9) = arrLHS(1, 4)
    arrResult(1, 10) = arrLHS(1, 3)
    arrResult(1, 11) = "Neolink Payment Date"
    arrResult(1, 12) = "Last Coupon Date Match"
    arrResult(1, 13) = "Difference Days"

    lngOutputRow = 1

    For lngRow = 2 To lngRows
        If HasNonEmptyValue(arrLHS(lngRow, 3)) Then
            lngOutputRow = lngOutputRow + 1

            arrResult(lngOutputRow, 1) = arrLHS(lngRow, 1)
            arrResult(lngOutputRow, 3) = arrLHS(lngRow, 2)
            arrResult(lngOutputRow, 4) = arrLHS(lngRow, 7)
            arrResult(lngOutputRow, 5) = arrLHS(lngRow, 8)
            arrResult(lngOutputRow, 6) = arrLHS(lngRow, 9)
            arrResult(lngOutputRow, 7) = arrLHS(lngRow, 10)
            arrResult(lngOutputRow, 9) = arrLHS(lngRow, 4)
            arrResult(lngOutputRow, 10) = arrLHS(lngRow, 3)

            strKey = BuildFundIsinKey(arrLHS(lngRow, 1), arrLHS(lngRow, 2))

            If Len(strKey) = 0 Or Not dictLatestPayment.Exists(strKey) Then
                arrResult(lngOutputRow, 12) = "NOT FOUND"
            Else
                arrLookupValue = dictLatestPayment(strKey)
                dblPaymentDate = CDbl(arrLookupValue(0))

                arrResult(lngOutputRow, 2) = CStr(arrLookupValue(1))
                If CDbl(arrLookupValue(2)) > 0 Then
                    arrResult(lngOutputRow, 8) = CDbl(arrLookupValue(2))
                End If
                arrResult(lngOutputRow, 11) = dblPaymentDate

                blnHasLastCouponDate = TryGetExcelDateSerial( _
                    arrLHS(lngRow, 3), dblLastCouponDate)

                If blnHasLastCouponDate Then
                    lngLastCouponDay = CLng(Int(dblLastCouponDate))
                    lngPaymentDay = CLng(Int(dblPaymentDate))

                    If lngLastCouponDay = lngPaymentDay Then
                        arrResult(lngOutputRow, 12) = "TRUE"
                    Else
                        arrResult(lngOutputRow, 12) = "FALSE"
                    End If

                    ' Positive value means Neolink Payment Date is later than LHS Last Coupon Date.
                    arrResult(lngOutputRow, 13) = lngPaymentDay - lngLastCouponDay
                Else
                    arrResult(lngOutputRow, 12) = "NO LHS DATE"
                End If
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
' Returns:       Object - Fund + ISIN -> Array(Payment Date, account, Value Date)
' Description:   Uses only GL:Type code = INTR and keeps the row with the greatest
'                Payment Date. Account and Value Date come from that same row.
'-------------------------------------------------------------------------------
Private Function BuildLatestNeolinkPaymentLookup(ByVal arrNeolink As Variant) As Object
    Const METHOD_NAME As String = "BuildLatestNeolinkPaymentLookup"
    Dim arrExisting As Variant
    Dim blnHasValueDate As Boolean
    Dim dictLatestPayment As Object
    Dim dblPaymentDate As Double
    Dim dblValueDate As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngRow As Long
    Dim strAccount As String
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
                    strAccount = GetLastElevenAccountCharacters(arrNeolink(lngRow, 2))
                    blnHasValueDate = TryGetExcelDateSerial( _
                        arrNeolink(lngRow, 6), dblValueDate)

                    If Not blnHasValueDate Then dblValueDate = 0

                    If dictLatestPayment.Exists(strKey) Then
                        arrExisting = dictLatestPayment(strKey)

                        If dblPaymentDate > CDbl(arrExisting(0)) Then
                            dictLatestPayment(strKey) = _
                                Array(dblPaymentDate, strAccount, dblValueDate)
                        End If
                    Else
                        dictLatestPayment.Add strKey, _
                            Array(dblPaymentDate, strAccount, dblValueDate)
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
' Parameters:    vValue - Neolink GL:Securities account value
' Returns:       String - last 11 normalized account characters, or empty string
' Description:   Removes spaces and returns the last 11 characters.
'-------------------------------------------------------------------------------
Private Function GetLastElevenAccountCharacters(ByVal vValue As Variant) As String
    Const METHOD_NAME As String = "GetLastElevenAccountCharacters"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strAccount As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(vValue) Or IsEmpty(vValue) Then GoTo ExitPoint

    strAccount = Replace(Trim$(CStr(vValue)), " ", vbNullString)

    If Len(strAccount) >= 11 Then
        GetLastElevenAccountCharacters = Right$(strAccount, 11)
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
' Parameters:    vValue - candidate value
' Returns:       Boolean - True when the value is not empty/blank
' Description:   Used to exclude LHS rows whose Last Coupon Date is blank.
'-------------------------------------------------------------------------------
Private Function HasNonEmptyValue(ByVal vValue As Variant) As Boolean
    Const METHOD_NAME As String = "HasNonEmptyValue"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(vValue) Or IsEmpty(vValue) Then GoTo ExitPoint

    HasNonEmptyValue = (Len(Trim$(CStr(vValue))) > 0)

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
' Description:   Colors TRUE matches green and FALSE matches red in column L.
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

    Set rngMatch = wksOutputAll.Range("L2:L" & CStr(lngLastRow))
    rngMatch.FormatConditions.Delete

    With rngMatch.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$L2=""TRUE""")
        .Interior.Color = RGB(198, 239, 206)
    End With

    With rngMatch.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=$L2=""FALSE""")
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
' Parameters:    wksOutputAll - Output_All worksheet
' Returns:       None
' Description:   Visually distinguishes the two dates being compared. Uses very
'                light fills so the output remains readable and print-friendly.
'-------------------------------------------------------------------------------
Private Sub ApplyOutputAllDateHighlighting(ByVal wksOutputAll As Excel.Worksheet)
    Const METHOD_NAME As String = "ApplyOutputAllDateHighlighting"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLastRow As Long
    Dim rngLastCoupon As Excel.Range
    Dim rngPaymentDate As Excel.Range

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastRow = wksOutputAll.Cells(wksOutputAll.Rows.Count, 1).End(xlUp).Row

    With wksOutputAll.Range("J1")
        .Font.Bold = True
        .Interior.Color = RGB(255, 246, 214)
    End With

    With wksOutputAll.Range("K1")
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With

    If lngLastRow < 2 Then GoTo ExitPoint

    Set rngLastCoupon = wksOutputAll.Range("J2:J" & CStr(lngLastRow))
    Set rngPaymentDate = wksOutputAll.Range("K2:K" & CStr(lngLastRow))

    rngLastCoupon.Interior.Color = RGB(255, 252, 240)
    rngPaymentDate.Interior.Color = RGB(245, 250, 255)

ExitPoint:
    Set rngLastCoupon = Nothing
    Set rngPaymentDate = Nothing

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
' Creation date: 2026-09-09
' Parameters:    wksOutputAll - Output_All worksheet
' Returns:       None
' Description:   Gives Neolink Value Date a light red fill when it differs from
'                Payment Date. Blank or non-date Value Dates are not highlighted.
'-------------------------------------------------------------------------------
Private Sub ApplyValueDateMismatchFormatting(ByVal wksOutputAll As Excel.Worksheet)
    Const METHOD_NAME As String = "ApplyValueDateMismatchFormatting"
    Dim arrPaymentDate As Variant
    Dim arrValueDate As Variant
    Dim blnMismatch As Boolean
    Dim dblPaymentDate As Double
    Dim dblValueDate As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    Dim lngRunStart As Long
    Dim rngMismatch As Excel.Range
    Dim rngRun As Excel.Range
    Dim rngValueDate As Excel.Range

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastRow = wksOutputAll.Cells(wksOutputAll.Rows.Count, 1).End(xlUp).Row
    If lngLastRow < 2 Then GoTo ExitPoint

    Set rngValueDate = wksOutputAll.Range("H2:H" & CStr(lngLastRow))
    arrValueDate = rngValueDate.Value2
    arrPaymentDate = wksOutputAll.Range("K2:K" & CStr(lngLastRow)).Value2

    ' Remove only formatting previously applied by this routine.
    rngValueDate.FormatConditions.Delete
    rngValueDate.Interior.Pattern = xlNone

    ' Compare the values in VBA instead of using a conditional-format formula.
    ' This avoids locale-dependent Formula1 parsing that can raise VBA Error 5.
    For lngIndex = 1 To UBound(arrValueDate, 1)
        blnMismatch = False

        If TryGetExcelDateSerial(arrValueDate(lngIndex, 1), dblValueDate) Then
            If TryGetExcelDateSerial(arrPaymentDate(lngIndex, 1), dblPaymentDate) Then
                blnMismatch = (CLng(Int(dblValueDate)) <> CLng(Int(dblPaymentDate)))
            End If
        End If

        If blnMismatch Then
            If lngRunStart = 0 Then lngRunStart = lngIndex
        ElseIf lngRunStart > 0 Then
            Set rngRun = rngValueDate.Cells(lngRunStart, 1).Resize(lngIndex - lngRunStart, 1)
            If rngMismatch Is Nothing Then
                Set rngMismatch = rngRun
            Else
                Set rngMismatch = Application.Union(rngMismatch, rngRun)
            End If
            lngRunStart = 0
        End If
    Next lngIndex

    If lngRunStart > 0 Then
        Set rngRun = rngValueDate.Cells(lngRunStart, 1).Resize( _
            UBound(arrValueDate, 1) - lngRunStart + 1, 1)
        If rngMismatch Is Nothing Then
            Set rngMismatch = rngRun
        Else
            Set rngMismatch = Application.Union(rngMismatch, rngRun)
        End If
    End If

    If Not rngMismatch Is Nothing Then
        rngMismatch.Interior.Color = RGB(255, 230, 230)
    End If

ExitPoint:
    Set rngRun = Nothing
    Set rngMismatch = Nothing
    Set rngValueDate = Nothing

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
' Creation date: 2026-09-10
' Parameters:    wksOutputAll - Output_All worksheet
' Returns:       None
' Description:   Adds lightweight Data Validation input-message tooltips to the
'                coupon frequency and bond calculation method columns. One rule
'                per column keeps the runtime impact negligible even for large files.
'-------------------------------------------------------------------------------
Private Sub ApplyMappingTooltips(ByVal wksOutputAll As Excel.Worksheet)
    Const METHOD_NAME As String = "ApplyMappingTooltips"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLastRow As Long
    Dim rngBondCalc As Excel.Range
    Dim rngCouponFrequency As Excel.Range
    Dim strBondTooltip As String
    Dim strCouponTooltip As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastRow = wksOutputAll.Cells(wksOutputAll.Rows.Count, 1).End(xlUp).Row
    If lngLastRow < 2 Then GoTo ExitPoint

    strCouponTooltip = _
        "1 - End of Month" & vbLf & _
        "2 - End of Quarter" & vbLf & _
        "3 - End of half-year" & vbLf & _
        "4 - End of Year" & vbLf & _
        "5 - Monthly" & vbLf & _
        "6 - Quarterly" & vbLf & _
        "7 - Half-year" & vbLf & _
        "8 - Year" & vbLf & _
        "9 - Maturity"

    strBondTooltip = _
        "1 - 365-6/365" & vbLf & _
        "2 - 365-6/360" & vbLf & _
        "3 - 360/360" & vbLf & _
        "4 - 360/365" & vbLf & _
        "5 - 365-6/365-6" & vbLf & _
        "6 - 365-6/365-6 = Civil" & vbLf & _
        "7 - 360/360 US" & vbLf & _
        "8 - Actual/364" & vbLf & _
        "9 - Actual/252"

    Set rngCouponFrequency = wksOutputAll.Range("F2:F" & CStr(lngLastRow))
    Set rngBondCalc = wksOutputAll.Range("G2:G" & CStr(lngLastRow))

    With rngCouponFrequency.Validation
        .Add Type:=xlValidateInputOnly
        .IgnoreBlank = True
        .InputTitle = "Coupon Frequency"
        .InputMessage = strCouponTooltip
        .ShowInput = True
        .ShowError = False
    End With

    With rngBondCalc.Validation
        .Add Type:=xlValidateInputOnly
        .IgnoreBlank = True
        .InputTitle = "Interest calculation type"
        .InputMessage = strBondTooltip
        .ShowInput = True
        .ShowError = False
    End With

ExitPoint:
    Set rngCouponFrequency = Nothing
    Set rngBondCalc = Nothing

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
' Creation date: 2026-09-10
' Parameters:    wksOutputAll - Output_All worksheet
' Returns:       None
' Description:   Highlights weekend dates in Next Coupon Date, Last Coupon Date
'                and Neolink Payment Date. Saturday and Sunday use distinct fills.
'-------------------------------------------------------------------------------
Private Sub ApplyWeekendDateHighlighting(ByVal wksOutputAll As Excel.Worksheet)
    Const METHOD_NAME As String = "ApplyWeekendDateHighlighting"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLastRow As Long
    Dim rngWeekendDates As Excel.Range

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastRow = wksOutputAll.Cells(wksOutputAll.Rows.Count, 1).End(xlUp).Row
    If lngLastRow < 2 Then GoTo ExitPoint

    Set rngWeekendDates = wksOutputAll.Range("I2:K" & CStr(lngLastRow))

    ' Simple formulas without list separators keep the rules locale-safe.
    With rngWeekendDates.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=ISNUMBER(I2)*(WEEKDAY(I2)=7)")
        .Interior.Color = RGB(226, 239, 218)
        .StopIfTrue = True
    End With

    With rngWeekendDates.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=ISNUMBER(I2)*(WEEKDAY(I2)=1)")
        .Interior.Color = RGB(252, 228, 214)
        .StopIfTrue = True
    End With

ExitPoint:
    Set rngWeekendDates = Nothing

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
' Creation date: 2026-09-10
' Parameters:    wksTarget - worksheet to format
' Returns:       None
' Description:   Wraps and enlarges the frozen header row, auto-fits columns and
'                caps excessive widths so long headers do not make sheets unwieldy.
'-------------------------------------------------------------------------------
Private Sub FormatWorksheetLayout(ByVal wksTarget As Excel.Worksheet)
    Const METHOD_NAME As String = "FormatWorksheetLayout"
    Const MAX_COLUMN_WIDTH As Double = 20
    Const HEADER_HEIGHT_FACTOR As Double = 3
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngLastColumn As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastColumn = wksTarget.Cells(1, wksTarget.Columns.Count).End(xlToLeft).Column

    With wksTarget.Rows(1)
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = wksTarget.StandardHeight * HEADER_HEIGHT_FACTOR
    End With

    wksTarget.UsedRange.Columns.AutoFit

    For lngColumn = 1 To lngLastColumn
        If wksTarget.Columns(lngColumn).ColumnWidth > MAX_COLUMN_WIDTH Then
            wksTarget.Columns(lngColumn).ColumnWidth = MAX_COLUMN_WIDTH
        End If
    Next lngColumn

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

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    wkbTarget - output workbook
' Returns:       None
' Description:   Freezes row 1 on every worksheet. Excel exposes FreezePanes at
'                window level, so each sheet must briefly become the active sheet.
'-------------------------------------------------------------------------------
Private Sub FreezeTopRowOnAllSheets(ByVal wkbTarget As Excel.Workbook)
    Const METHOD_NAME As String = "FreezeTopRowOnAllSheets"
    Dim errDescription As String
    Dim errNumber As Long
    Dim wndTarget As Excel.Window
    Dim wksOriginal As Excel.Worksheet
    Dim wksTarget As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If wkbTarget Is Nothing Then
        Err.Raise ERR_OUTPUT, METHOD_NAME, "Output workbook is not initialized."
    End If

    If wkbTarget.Windows.Count = 0 Then
        Err.Raise ERR_OUTPUT, METHOD_NAME, "Output workbook does not have an Excel window."
    End If

    Set wndTarget = wkbTarget.Windows(1)

    If TypeName(wkbTarget.ActiveSheet) = "Worksheet" Then
        Set wksOriginal = wkbTarget.ActiveSheet
    End If

    wndTarget.Activate

    For Each wksTarget In wkbTarget.Worksheets
        ' FreezePanes is a Window property and applies to the active worksheet.
        wksTarget.Activate

        With wndTarget
            .FreezePanes = False
            .SplitColumn = 0
            .SplitRow = 1
            .FreezePanes = True
        End With
    Next wksTarget

    If Not wksOriginal Is Nothing Then
        wksOriginal.Activate
    End If

ExitPoint:
    Set wksTarget = Nothing
    Set wksOriginal = Nothing
    Set wndTarget = Nothing

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

