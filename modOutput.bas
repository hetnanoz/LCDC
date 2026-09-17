Option Explicit

Private Const CLASS_NAME As String = "modOutput"
Private Const OUTPUT_BASE_NAME As String = "Last_Coupon_Date_Checker_"
Private Const ERR_OUTPUT As Long = vbObjectError + 6600
Private Const LOOKUP_SEPARATOR As String = "|"
Private Const PAYMENT_PREVIOUS_WINDOW_DAYS As Long = 10
Private Const VERIFIED_ACCRUAL_METHOD As Long = 1
Private Const ACCRUAL_MATCH_TOLERANCE As Double = 0.05
Private Const CURRENT_FACE_MATCH_TOLERANCE As Double = 0.01

' Output_All column positions. Keep these centralized so formatting and analysis
' stay aligned when diagnostic columns are added.
Private Const COL_FUND As Long = 1
Private Const COL_ACCOUNT As Long = 2
Private Const COL_ISIN As Long = 3
Private Const COL_GTI_NAME As Long = 4
Private Const COL_ACCOUNTING_DATE As Long = 5
Private Const COL_GENERATION_DATE As Long = 6
Private Const COL_COUPON_FREQ As Long = 7
Private Const COL_BOND_CALC As Long = 8
Private Const COL_START_DATE As Long = 9
Private Const COL_FIRST_COUPON As Long = 10
Private Const COL_MATURITY As Long = 11
Private Const COL_INTEREST_RATE As Long = 12
Private Const COL_QUANTITY As Long = 13
Private Const COL_FACE_VALUE As Long = 14
Private Const COL_REIMBURSEMENT_FACTOR As Long = 15
Private Const COL_CALCULATED_CURRENT_FACE As Long = 16
Private Const COL_CURRENT_FACE_MATCH As Long = 17
Private Const COL_INTEREST_RATE_TYPE As Long = 18
Private Const COL_INDEXATION_MODE As Long = 19
Private Const COL_DAYS_SINCE_LAST As Long = 20
Private Const COL_VALUE_DATE As Long = 21
Private Const COL_LHS_NEXT As Long = 22
Private Const COL_EXPECTED_NEXT As Long = 23
Private Const COL_LHS_LAST As Long = 24
Private Const COL_EXPECTED_LAST As Long = 25
Private Const COL_PAYMENT As Long = 26
Private Const COL_SCHEDULE_MATCH As Long = 27
Private Const COL_PAYMENT_SHIFT As Long = 28
Private Const COL_VALIDATION_RESULT As Long = 29
Private Const COL_PRODUCT_REVIEW As Long = 30
Private Const COL_STUB_FLAG As Long = 31
Private Const COL_PAYMENT_MATCH As Long = 32
Private Const COL_PAYMENT_DIFF As Long = 33
Private Const COL_PAYMENT_SELECTION As Long = 34
Private Const COL_LATEST_PAYMENT As Long = 35
Private Const COL_ACCRUAL_CONVENTION As Long = 36
Private Const COL_ACCRUAL_PERIOD As Long = 37
Private Const COL_ACCRUAL_DAYS_SCHEDULE As Long = 38
Private Const COL_ACCRUAL_DAYS_DIFF As Long = 39
Private Const COL_ACCRUAL_DAYS_LHS As Long = 40
Private Const COL_COUPON_PERIOD_DAYS As Long = 41
Private Const COL_YEAR_FRACTION As Long = 42
Private Const COL_DAILY_ACCRUAL As Long = 43
Private Const COL_EXPECTED_FULL_COUPON As Long = 44
Private Const COL_ACCRUED_PERIOD_PCT As Long = 45
Private Const COL_ACCRUAL_REMAINING As Long = 46
Private Const COL_EXPECTED_ACCRUAL As Long = 47
Private Const COL_EXPECTED_ACCRUAL_LHS As Long = 48
Private Const COL_LHS_ACCRUED As Long = 49
Private Const COL_ACCRUAL_DIFF As Long = 50
Private Const COL_ACCRUAL_DIFF_PCT As Long = 51
Private Const COL_ACCRUAL_MATCH As Long = 52
Private Const COL_ACCRUAL_RESULT As Long = 53
Private Const OUTPUT_ALL_COLUMN_COUNT As Long = 53

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    arrLHS - LHS output array; arrNeolink - Neolink output array
' Returns:       Excel.Workbook - newly created output workbook
' Description:   Creates input, validation, schedule and accrual analysis sheets.
'-------------------------------------------------------------------------------
Public Function CreateOutputWorkbook( _
    ByVal arrLHS As Variant, ByVal arrNeolink As Variant) As Excel.Workbook

    Const METHOD_NAME As String = "CreateOutputWorkbook"
    Dim arrOutputAll As Variant
    Dim blnCompleted As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim wkbOutput As Excel.Workbook
    Dim wksAccrualAnalysis As Excel.Worksheet
    Dim wksAnalysis As Excel.Worksheet
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

    Set wksAnalysis = wkbOutput.Worksheets.Add(After:=wksOutputAll)
    wksAnalysis.Name = "Analysis"

    Set wksAccrualAnalysis = wkbOutput.Worksheets.Add(After:=wksAnalysis)
    wksAccrualAnalysis.Name = "Accrual Analysis"

    ' Text identifiers must not be converted to numbers/scientific notation.
    wksLHS.Columns("A:B").NumberFormat = "@"
    wksNeolink.Columns("A:B").NumberFormat = "@"
    wksOutputAll.Columns(COL_FUND).NumberFormat = "@"
    wksOutputAll.Columns(COL_ACCOUNT).NumberFormat = "@"
    wksOutputAll.Columns(COL_ISIN).NumberFormat = "@"

    Application.StatusBar = "Last Coupon Date Checker: Writing input sheets..."
    Call WriteArrayToWorksheet(wksLHS, arrLHS)
    Call WriteArrayToWorksheet(wksNeolink, arrNeolink)

    Application.StatusBar = "Last Coupon Date Checker: Calculating schedules and accruals..."
    arrOutputAll = BuildOutputAllArray(arrLHS, arrNeolink)
    Application.StatusBar = "Last Coupon Date Checker: Writing Output_All..."
    Call WriteArrayToWorksheet(wksOutputAll, arrOutputAll)

    Application.StatusBar = "Last Coupon Date Checker: Building analysis sheets..."
    Call BuildAnalysisSheet(wksAnalysis, arrOutputAll)
    Call BuildAccrualAnalysisSheet(wksAccrualAnalysis, arrOutputAll)

    ' LHS date displays. Optional diagnostic dates are appended in columns O:P.
    wksLHS.Columns("C:D").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("G:H").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("K:K").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("O:P").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("S:T").NumberFormat = "#,##0.00"
    wksLHS.Columns("U:U").NumberFormat = "0.000000000"
    wksNeolink.Columns("F:G").NumberFormat = "dd/mm/yyyy"

    Application.StatusBar = "Last Coupon Date Checker: Formatting output..."
    Call SetOutputAllNumberFormats(wksOutputAll)
    Call ApplyMatchFormatting(wksOutputAll)
    Call ApplyOutputAllDateHighlighting(wksOutputAll)
    Call ApplyWeekendDateHighlighting(wksOutputAll)
    Call ApplyValueDateMismatchFormatting(wksOutputAll)
    Call ApplyPreviousPaymentSelectionFormatting(wksOutputAll)
    Call ApplyMappingTooltips(wksOutputAll)

    Call FormatWorksheetLayout(wksLHS)
    Call FormatWorksheetLayout(wksNeolink)
    Call FormatWorksheetLayout(wksOutputAll)
    Call FormatWorksheetLayout(wksAnalysis)
    Call FormatWorksheetLayout(wksAccrualAnalysis)

    Call FreezeTopRowOnAllSheets(wkbOutput)
    Application.StatusBar = "Last Coupon Date Checker: Output created."

    Set CreateOutputWorkbook = wkbOutput
    blnCompleted = True

ExitPoint:
    Set wksLHS = Nothing
    Set wksNeolink = Nothing
    Set wksOutputAll = Nothing
    Set wksAnalysis = Nothing
    Set wksAccrualAnalysis = Nothing

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
' Returns:       Variant - diagnostic Output_All array including headers
' Description:   Keeps only LHS rows with a non-empty Last Coupon Date, builds an
'                independent expected coupon schedule and keeps Neolink payment data
'                as supporting evidence rather than the primary schedule source.
'-------------------------------------------------------------------------------
Private Function BuildOutputAllArray( _
    ByVal arrLHS As Variant, ByVal arrNeolink As Variant) As Variant

    Const METHOD_NAME As String = "BuildOutputAllArray"
    Dim arrAccrualConventionText As Variant
    Dim arrAccrualPeriodText As Variant
    Dim arrLookupValue As Variant
    Dim arrResult() As Variant
    Dim arrScheduleValue As Variant
    Dim blnHasLastCouponDate As Boolean
    Dim blnHasPayment As Boolean
    Dim blnScheduleBuilt As Boolean
    Dim dictPreferredPayment As Object
    Dim dictProductReview As Object
    Dim dictScheduleCache As Object
    Dim dblAccountingDate As Double
    Dim dblExpectedLast As Double
    Dim dblExpectedNext As Double
    Dim dblCalculatedCurrentFace As Double
    Dim dblFaceValue As Double
    Dim dblReimbursementFactor As Double
    Dim dblQuantityForFace As Double
    Dim dblFirstCoupon As Double
    Dim dblLastCouponDate As Double
    Dim dblMaturityDate As Double
    Dim dblPaymentDate As Double
    Dim dblStartDate As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngBondCalc As Long
    Dim lngFrequency As Long
    Dim lngLastCouponDay As Long
    Dim lngLhsCols As Long
    Dim lngOutputRow As Long
    Dim lngPaymentDay As Long
    Dim lngRow As Long
    Dim lngRows As Long
    Dim lngRowsWithLastCouponDate As Long
    Dim strGtiName As String
    Dim strKey As String
    Dim strProductReview As String
    Dim strScheduleCacheKey As String
    Dim strScheduleReason As String
    Dim strStubFlag As String
    Dim vLastCoupon As Variant

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictPreferredPayment = BuildPreferredNeolinkPaymentLookup(arrNeolink)
    Set dictProductReview = CreateObject("Scripting.Dictionary")
    dictProductReview.CompareMode = vbTextCompare
    Set dictScheduleCache = CreateObject("Scripting.Dictionary")
    dictScheduleCache.CompareMode = vbBinaryCompare

    arrAccrualConventionText = Array( _
        vbNullString, "1 - 365-6/365", "2 - 365-6/360", "3 - 360/360", _
        "4 - 360/365", "5 - 365-6/365-6", "6 - 365-6/365-6 = Civil", _
        "7 - 360/360 US", "8 - Actual/364", "9 - Actual/252")
    arrAccrualPeriodText = Array( _
        vbNullString, "End of Month", "End of Quarter", "End of half-year", _
        "End of Year", "Monthly", "Quarterly", "Half-year", "Year", "Maturity")

    lngRows = UBound(arrLHS, 1)
    lngLhsCols = UBound(arrLHS, 2)

    ' Count qualifying rows without calling a helper for every LHS record.
    For lngRow = 2 To lngRows
        vLastCoupon = arrLHS(lngRow, 3)
        If Not IsError(vLastCoupon) And Not IsEmpty(vLastCoupon) Then
            If VarType(vLastCoupon) <> vbString Or LenB(Trim$(CStr(vLastCoupon))) > 0 Then
                lngRowsWithLastCouponDate = lngRowsWithLastCouponDate + 1
            End If
        End If
    Next lngRow

    ReDim arrResult(1 To lngRowsWithLastCouponDate + 1, 1 To OUTPUT_ALL_COLUMN_COUNT)
    Call SetOutputAllHeaders(arrResult)

    lngOutputRow = 1

    For lngRow = 2 To lngRows
        vLastCoupon = arrLHS(lngRow, 3)
        blnHasLastCouponDate = False
        If Not IsError(vLastCoupon) And Not IsEmpty(vLastCoupon) Then
            If VarType(vLastCoupon) <> vbString Or LenB(Trim$(CStr(vLastCoupon))) > 0 Then
                blnHasLastCouponDate = True
            End If
        End If

        If blnHasLastCouponDate Then
            lngOutputRow = lngOutputRow + 1

            arrResult(lngOutputRow, COL_FUND) = arrLHS(lngRow, 1)
            arrResult(lngOutputRow, COL_ISIN) = arrLHS(lngRow, 2)
            arrResult(lngOutputRow, COL_GTI_NAME) = arrLHS(lngRow, 6)
            arrResult(lngOutputRow, COL_ACCOUNTING_DATE) = arrLHS(lngRow, 7)
            arrResult(lngOutputRow, COL_GENERATION_DATE) = arrLHS(lngRow, 8)
            arrResult(lngOutputRow, COL_COUPON_FREQ) = arrLHS(lngRow, 9)
            arrResult(lngOutputRow, COL_BOND_CALC) = arrLHS(lngRow, 10)
            arrResult(lngOutputRow, COL_START_DATE) = arrLHS(lngRow, 11)
            arrResult(lngOutputRow, COL_INTEREST_RATE) = arrLHS(lngRow, 12)
            If lngLhsCols >= 19 Then arrResult(lngOutputRow, COL_QUANTITY) = arrLHS(lngRow, 19)
            If lngLhsCols >= 20 Then arrResult(lngOutputRow, COL_FACE_VALUE) = arrLHS(lngRow, 20)
            If lngLhsCols >= 21 Then _
                arrResult(lngOutputRow, COL_REIMBURSEMENT_FACTOR) = arrLHS(lngRow, 21)

            ' MBS/current-face diagnostic kept inline for the hot loop.
            If lngLhsCols >= 21 Then
                If IsNumeric(arrLHS(lngRow, 20)) And IsNumeric(arrLHS(lngRow, 21)) Then
                    dblFaceValue = CDbl(arrLHS(lngRow, 20))
                    dblReimbursementFactor = CDbl(arrLHS(lngRow, 21))
                    If dblFaceValue <> 0 And dblReimbursementFactor <> 0 Then
                        dblCalculatedCurrentFace = dblFaceValue * dblReimbursementFactor
                        arrResult(lngOutputRow, COL_CALCULATED_CURRENT_FACE) = _
                            dblCalculatedCurrentFace

                        If IsNumeric(arrLHS(lngRow, 19)) Then
                            dblQuantityForFace = CDbl(arrLHS(lngRow, 19))
                            If Abs(dblQuantityForFace - dblCalculatedCurrentFace) <= _
                               CURRENT_FACE_MATCH_TOLERANCE Then
                                arrResult(lngOutputRow, COL_CURRENT_FACE_MATCH) = "TRUE"
                            Else
                                arrResult(lngOutputRow, COL_CURRENT_FACE_MATCH) = "FALSE"
                            End If
                        Else
                            arrResult(lngOutputRow, COL_CURRENT_FACE_MATCH) = "REVIEW"
                        End If
                    End If
                End If
            End If

            arrResult(lngOutputRow, COL_DAYS_SINCE_LAST) = arrLHS(lngRow, 5)
            arrResult(lngOutputRow, COL_LHS_NEXT) = arrLHS(lngRow, 4)
            arrResult(lngOutputRow, COL_LHS_LAST) = vLastCoupon

            If lngLhsCols >= 15 Then arrResult(lngOutputRow, COL_MATURITY) = arrLHS(lngRow, 15)
            If lngLhsCols >= 16 Then arrResult(lngOutputRow, COL_FIRST_COUPON) = arrLHS(lngRow, 16)
            If lngLhsCols >= 17 Then arrResult(lngOutputRow, COL_INTEREST_RATE_TYPE) = arrLHS(lngRow, 17)
            If lngLhsCols >= 18 Then arrResult(lngOutputRow, COL_INDEXATION_MODE) = arrLHS(lngRow, 18)

            ' Cache GTI classification because the same product names repeat many times.
            strGtiName = vbNullString
            If Not IsError(arrLHS(lngRow, 6)) And Not IsEmpty(arrLHS(lngRow, 6)) Then
                strGtiName = Trim$(CStr(arrLHS(lngRow, 6)))
            End If
            If dictProductReview.Exists(strGtiName) Then
                strProductReview = CStr(dictProductReview(strGtiName))
            Else
                strProductReview = GetProductReviewFlag(strGtiName)
                dictProductReview.Add strGtiName, strProductReview
            End If
            arrResult(lngOutputRow, COL_PRODUCT_REVIEW) = strProductReview

            strKey = BuildFundIsinKey(arrLHS(lngRow, 1), arrLHS(lngRow, 2))
            blnHasPayment = False
            dblPaymentDate = 0

            If Len(strKey) > 0 And dictPreferredPayment.Exists(strKey) Then
                arrLookupValue = dictPreferredPayment(strKey)
                dblPaymentDate = CDbl(arrLookupValue(0))
                blnHasPayment = True

                arrResult(lngOutputRow, COL_ACCOUNT) = CStr(arrLookupValue(1))
                If CDbl(arrLookupValue(2)) > 0 Then
                    arrResult(lngOutputRow, COL_VALUE_DATE) = CDbl(arrLookupValue(2))
                End If
                arrResult(lngOutputRow, COL_PAYMENT) = dblPaymentDate

                If CBool(arrLookupValue(3)) Then
                    arrResult(lngOutputRow, COL_PAYMENT_SELECTION) = "PREVIOUS DATE USED"
                    arrResult(lngOutputRow, COL_LATEST_PAYMENT) = CDbl(arrLookupValue(4))
                End If
            End If

            ' LHS dates are normalized to numeric Excel serials by modInputLHS.
            dblLastCouponDate = 0
            If IsNumeric(vLastCoupon) Then
                dblLastCouponDate = CDbl(vLastCoupon)
            Else
                Call TryGetExcelDateSerial(vLastCoupon, dblLastCouponDate)
            End If

            If blnHasPayment And dblLastCouponDate > 0 Then
                lngLastCouponDay = CLng(Int(dblLastCouponDate))
                lngPaymentDay = CLng(Int(dblPaymentDate))
                If lngLastCouponDay = lngPaymentDay Then
                    arrResult(lngOutputRow, COL_PAYMENT_MATCH) = "TRUE"
                Else
                    arrResult(lngOutputRow, COL_PAYMENT_MATCH) = "FALSE"
                End If
                arrResult(lngOutputRow, COL_PAYMENT_DIFF) = lngPaymentDay - lngLastCouponDay
            ElseIf Not blnHasPayment Then
                arrResult(lngOutputRow, COL_PAYMENT_MATCH) = "NOT FOUND"
            End If

            dblAccountingDate = FastDateSerial(arrLHS(lngRow, 7))
            dblStartDate = FastDateSerial(arrLHS(lngRow, 11))
            dblMaturityDate = 0
            dblFirstCoupon = 0
            If lngLhsCols >= 15 Then dblMaturityDate = FastDateSerial(arrLHS(lngRow, 15))
            If lngLhsCols >= 16 Then dblFirstCoupon = FastDateSerial(arrLHS(lngRow, 16))

            lngFrequency = FastLongCode(arrLHS(lngRow, 9))
            lngBondCalc = FastLongCode(arrLHS(lngRow, 10))

            ' Cache schedule results. Identical securities/dates no longer recalculate
            ' DateDiff/DateAdd and stub checks for every repeated holding.
            strScheduleCacheKey = CStr(CLng(Int(dblAccountingDate))) & LOOKUP_SEPARATOR & _
                CStr(CLng(Int(dblStartDate))) & LOOKUP_SEPARATOR & _
                CStr(CLng(Int(dblFirstCoupon))) & LOOKUP_SEPARATOR & _
                CStr(CLng(Int(dblMaturityDate))) & LOOKUP_SEPARATOR & CStr(lngFrequency)

            If dictScheduleCache.Exists(strScheduleCacheKey) Then
                arrScheduleValue = dictScheduleCache(strScheduleCacheKey)
                blnScheduleBuilt = CBool(arrScheduleValue(0))
                dblExpectedLast = CDbl(arrScheduleValue(1))
                dblExpectedNext = CDbl(arrScheduleValue(2))
                strStubFlag = CStr(arrScheduleValue(3))
                strScheduleReason = CStr(arrScheduleValue(4))
            Else
                dblExpectedLast = 0
                dblExpectedNext = 0
                strStubFlag = vbNullString
                strScheduleReason = vbNullString
                blnScheduleBuilt = TryBuildExpectedCouponSchedule( _
                    dblAccountingDate, dblStartDate, dblFirstCoupon, dblMaturityDate, _
                    lngFrequency, dblExpectedLast, dblExpectedNext, _
                    strStubFlag, strScheduleReason)
                dictScheduleCache.Add strScheduleCacheKey, _
                    Array(blnScheduleBuilt, dblExpectedLast, dblExpectedNext, _
                          strStubFlag, strScheduleReason)
            End If

            arrResult(lngOutputRow, COL_STUB_FLAG) = strStubFlag

            If blnScheduleBuilt Then
                arrResult(lngOutputRow, COL_EXPECTED_LAST) = dblExpectedLast
                If dblExpectedNext > 0 Then arrResult(lngOutputRow, COL_EXPECTED_NEXT) = dblExpectedNext

                If dblLastCouponDate > 0 Then
                    If CLng(Int(dblLastCouponDate)) = CLng(Int(dblExpectedLast)) Then
                        arrResult(lngOutputRow, COL_SCHEDULE_MATCH) = "TRUE"
                    Else
                        arrResult(lngOutputRow, COL_SCHEDULE_MATCH) = "FALSE"
                    End If
                Else
                    arrResult(lngOutputRow, COL_SCHEDULE_MATCH) = "REVIEW"
                End If

                If blnHasPayment Then
                    arrResult(lngOutputRow, COL_PAYMENT_SHIFT) = _
                        CLng(Int(dblPaymentDate)) - CLng(Int(dblExpectedLast))
                End If

                arrResult(lngOutputRow, COL_VALIDATION_RESULT) = BuildValidationResult( _
                    CStr(arrResult(lngOutputRow, COL_SCHEDULE_MATCH)), _
                    strProductReview, strStubFlag, blnHasPayment, _
                    arrResult(lngOutputRow, COL_PAYMENT_SHIFT))
            Else
                arrResult(lngOutputRow, COL_SCHEDULE_MATCH) = "REVIEW"
                arrResult(lngOutputRow, COL_VALIDATION_RESULT) = strScheduleReason
            End If

            If lngBondCalc >= 1 And lngBondCalc <= 9 Then
                arrResult(lngOutputRow, COL_ACCRUAL_CONVENTION) = _
                    arrAccrualConventionText(lngBondCalc)
            Else
                arrResult(lngOutputRow, COL_ACCRUAL_CONVENTION) = "(missing)"
            End If
            If lngFrequency >= 1 And lngFrequency <= 9 Then
                arrResult(lngOutputRow, COL_ACCRUAL_PERIOD) = arrAccrualPeriodText(lngFrequency)
            Else
                arrResult(lngOutputRow, COL_ACCRUAL_PERIOD) = "(missing)"
            End If

            If lngBondCalc = VERIFIED_ACCRUAL_METHOD Then
                Call PopulateVerifiedMethod1AccrualFast( _
                    arrResult, lngOutputRow, arrLHS, lngRow, lngFrequency, _
                    blnScheduleBuilt, dblAccountingDate, dblExpectedLast, _
                    dblExpectedNext, strProductReview, strStubFlag)
            ElseIf lngBondCalc = 0 Then
                arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
                arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
                    "REVIEW - BOND CALC METHOD MISSING"
            Else
                arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
                arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
                    "REVIEW - METHOD NOT VERIFIED"
            End If
        End If

        If lngRow Mod 5000 = 0 Then
            Application.StatusBar = "Last Coupon Date Checker: Calculating row " & _
                CStr(lngRow) & " / " & CStr(lngRows) & "..."
            DoEvents
        End If
    Next lngRow

    BuildOutputAllArray = arrResult

ExitPoint:
    Set dictPreferredPayment = Nothing
    Set dictProductReview = Nothing
    Set dictScheduleCache = Nothing

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


Private Function FastDateSerial(ByVal vValue As Variant) As Double
    Const METHOD_NAME As String = "FastDateSerial"
    Dim dblParsed As Double
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(vValue) Or IsEmpty(vValue) Then GoTo ExitPoint
    If IsNumeric(vValue) Then
        If CDbl(vValue) > 0 Then FastDateSerial = CDbl(vValue)
        GoTo ExitPoint
    End If

    If TryGetExcelDateSerial(vValue, dblParsed) Then FastDateSerial = dblParsed

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

Private Function FastLongCode(ByVal vValue As Variant) As Long
    Const METHOD_NAME As String = "FastLongCode"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strValue As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(vValue) Or IsEmpty(vValue) Then GoTo ExitPoint
    If IsNumeric(vValue) Then
        FastLongCode = CLng(vValue)
        GoTo ExitPoint
    End If
    strValue = Trim$(CStr(vValue))
    If Len(strValue) > 0 And IsNumeric(strValue) Then FastLongCode = CLng(strValue)

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

Private Sub PopulateVerifiedMethod1AccrualFast( _
    ByRef arrResult As Variant, ByVal lngOutputRow As Long, _
    ByVal arrLHS As Variant, ByVal lngLHSRow As Long, _
    ByVal lngFrequency As Long, ByVal blnScheduleBuilt As Boolean, _
    ByVal dblAccountingDate As Double, ByVal dblExpectedLast As Double, _
    ByVal dblExpectedNext As Double, ByVal strProductReview As String, _
    ByVal strStubFlag As String)

    Const METHOD_NAME As String = "PopulateVerifiedMethod1AccrualFast"
    Dim blnHasLhsAccrued As Boolean
    Dim dblCouponAmount As Double
    Dim dblCouponFactor As Double
    Dim dblDailyAccrual As Double
    Dim dblExpectedAccrual As Double
    Dim dblExpectedAccrualLhs As Double
    Dim dblLhsAccrued As Double
    Dim dblLhsLast As Double
    Dim dblLhsNext As Double
    Dim dblQuantity As Double
    Dim dblRatePercent As Double
    Dim dblYearFraction As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngAccrualDays As Long
    Dim lngAccrualDaysLhs As Long
    Dim lngCouponPeriodDays As Long
    Dim lngCouponPeriodDaysLhs As Long
    Dim lngLhsDays As Long
    Dim strMatch As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not blnScheduleBuilt Or dblExpectedLast <= 0 Or dblExpectedNext <= 0 Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
            "REVIEW - EXPECTED COUPON SCHEDULE NOT AVAILABLE"
        GoTo ExitPoint
    End If

    If UBound(arrLHS, 2) < 19 Or Not IsNumeric(arrLHS(lngLHSRow, 19)) Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = "REVIEW - QUANTITY MISSING OR ZERO"
        GoTo ExitPoint
    End If
    dblQuantity = CDbl(arrLHS(lngLHSRow, 19))
    If dblQuantity = 0 Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = "REVIEW - QUANTITY MISSING OR ZERO"
        GoTo ExitPoint
    End If

    If Not IsNumeric(arrLHS(lngLHSRow, 12)) Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = "REVIEW - INTEREST RATE MISSING"
        GoTo ExitPoint
    End If
    dblRatePercent = CDbl(arrLHS(lngLHSRow, 12))

    Select Case lngFrequency
        Case 1, 5: dblCouponFactor = 1# / 12#
        Case 2, 6: dblCouponFactor = 1# / 4#
        Case 3, 7: dblCouponFactor = 1# / 2#
        Case 4, 8: dblCouponFactor = 1#
        Case Else
            arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
            arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
                "REVIEW - ACCRUAL FREQUENCY NOT SUPPORTED"
            GoTo ExitPoint
    End Select

    lngCouponPeriodDays = CLng(Int(dblExpectedNext)) - CLng(Int(dblExpectedLast))
    lngAccrualDays = CLng(Int(dblAccountingDate)) - CLng(Int(dblExpectedLast)) + 1
    If lngCouponPeriodDays <= 0 Or lngAccrualDays < 0 Or _
       lngAccrualDays > lngCouponPeriodDays Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
            "REVIEW - ACCOUNTING DATE OUTSIDE COUPON PERIOD"
        GoTo ExitPoint
    End If

    dblCouponAmount = dblQuantity * (dblRatePercent / 100#) * dblCouponFactor
    dblYearFraction = dblCouponFactor * CDbl(lngAccrualDays) / CDbl(lngCouponPeriodDays)
    dblDailyAccrual = dblCouponAmount / CDbl(lngCouponPeriodDays)
    dblExpectedAccrual = dblQuantity * (dblRatePercent / 100#) * dblYearFraction

    arrResult(lngOutputRow, COL_ACCRUAL_DAYS_SCHEDULE) = lngAccrualDays
    arrResult(lngOutputRow, COL_COUPON_PERIOD_DAYS) = lngCouponPeriodDays
    arrResult(lngOutputRow, COL_YEAR_FRACTION) = dblYearFraction
    arrResult(lngOutputRow, COL_DAILY_ACCRUAL) = dblDailyAccrual
    arrResult(lngOutputRow, COL_EXPECTED_FULL_COUPON) = dblCouponAmount
    arrResult(lngOutputRow, COL_ACCRUED_PERIOD_PCT) = _
        CDbl(lngAccrualDays) / CDbl(lngCouponPeriodDays)
    arrResult(lngOutputRow, COL_ACCRUAL_REMAINING) = _
        dblCouponAmount - dblExpectedAccrual
    arrResult(lngOutputRow, COL_EXPECTED_ACCRUAL) = dblExpectedAccrual

    If IsNumeric(arrLHS(lngLHSRow, 5)) Then
        lngLhsDays = CLng(arrLHS(lngLHSRow, 5))
        arrResult(lngOutputRow, COL_ACCRUAL_DAYS_DIFF) = lngLhsDays - lngAccrualDays
    End If

    dblLhsLast = FastDateSerial(arrLHS(lngLHSRow, 3))
    dblLhsNext = FastDateSerial(arrLHS(lngLHSRow, 4))
    If dblLhsLast > 0 And dblLhsNext > dblLhsLast Then
        lngCouponPeriodDaysLhs = CLng(Int(dblLhsNext)) - CLng(Int(dblLhsLast))
        lngAccrualDaysLhs = CLng(Int(dblAccountingDate)) - CLng(Int(dblLhsLast)) + 1
        If lngAccrualDaysLhs >= 0 And lngAccrualDaysLhs <= lngCouponPeriodDaysLhs Then
            dblExpectedAccrualLhs = dblQuantity * (dblRatePercent / 100#) * _
                dblCouponFactor * CDbl(lngAccrualDaysLhs) / CDbl(lngCouponPeriodDaysLhs)
            arrResult(lngOutputRow, COL_ACCRUAL_DAYS_LHS) = lngAccrualDaysLhs
            arrResult(lngOutputRow, COL_EXPECTED_ACCRUAL_LHS) = dblExpectedAccrualLhs
        End If
    End If

    blnHasLhsAccrued = IsNumeric(arrLHS(lngLHSRow, 13))
    If Not blnHasLhsAccrued Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
            "REVIEW - LHS ACCRUED INTEREST MISSING"
        GoTo ExitPoint
    End If

    dblLhsAccrued = CDbl(arrLHS(lngLHSRow, 13))
    arrResult(lngOutputRow, COL_LHS_ACCRUED) = dblLhsAccrued
    arrResult(lngOutputRow, COL_ACCRUAL_DIFF) = dblLhsAccrued - dblExpectedAccrual
    If Abs(dblExpectedAccrual) > 0.0000001 Then
        arrResult(lngOutputRow, COL_ACCRUAL_DIFF_PCT) = _
            (dblLhsAccrued - dblExpectedAccrual) / dblExpectedAccrual
    End If

    If Abs(dblLhsAccrued - dblExpectedAccrual) <= ACCRUAL_MATCH_TOLERANCE Then
        strMatch = "TRUE"
    Else
        strMatch = "FALSE"
    End If
    arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = strMatch
    arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = BuildAccrualValidationResult( _
        strMatch, strProductReview, strStubFlag)

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
        "lngLHSRow", lngLHSRow)
    GoTo ExitPoint
End Sub

Private Sub SetOutputAllHeaders(ByRef arrResult As Variant)
    Const METHOD_NAME As String = "SetOutputAllHeaders"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrResult(1, COL_FUND) = "FUND CODE"
    arrResult(1, COL_ACCOUNT) = "Securities account"
    arrResult(1, COL_ISIN) = "EXTERNAL VALUE CODE"
    arrResult(1, COL_GTI_NAME) = "GTI NAME"
    arrResult(1, COL_ACCOUNTING_DATE) = "ACCOUNTING DATE_1"
    arrResult(1, COL_GENERATION_DATE) = "GENERATION DATE AND TIME"
    arrResult(1, COL_COUPON_FREQ) = "Coupon frequency/IRS-CDS"
    arrResult(1, COL_BOND_CALC) = "BOND CALC METHOD/IRS-CDS"
    arrResult(1, COL_START_DATE) = "START DATE"
    arrResult(1, COL_FIRST_COUPON) = "FIRST COUPON DATE"
    arrResult(1, COL_MATURITY) = "MATURITY DATE"
    arrResult(1, COL_INTEREST_RATE) = "INTEREST RATE"
    arrResult(1, COL_QUANTITY) = "QUANTITY"
    arrResult(1, COL_FACE_VALUE) = "FACE VALUE"
    arrResult(1, COL_REIMBURSEMENT_FACTOR) = "REIMBURSEMENT FACTOR (MSB)"
    arrResult(1, COL_CALCULATED_CURRENT_FACE) = "Calculated Current Face"
    arrResult(1, COL_CURRENT_FACE_MATCH) = "Quantity vs Current Face Match"
    arrResult(1, COL_INTEREST_RATE_TYPE) = "INTEREST RATE TYPE"
    arrResult(1, COL_INDEXATION_MODE) = "SECURITY INDEXATION MODE"
    arrResult(1, COL_DAYS_SINCE_LAST) = "NBR OF DAYS SINCE LAST COUPON"
    arrResult(1, COL_VALUE_DATE) = "Neolink Value Date"
    arrResult(1, COL_LHS_NEXT) = "NEXT COUPON DATE"
    arrResult(1, COL_EXPECTED_NEXT) = "Expected Next Coupon Date"
    arrResult(1, COL_LHS_LAST) = "LAST COUPON DATE"
    arrResult(1, COL_EXPECTED_LAST) = "Expected Last Coupon Date"
    arrResult(1, COL_PAYMENT) = "Neolink Payment Date"
    arrResult(1, COL_SCHEDULE_MATCH) = "Schedule Match"
    arrResult(1, COL_PAYMENT_SHIFT) = "Payment Shift Days"
    arrResult(1, COL_VALIDATION_RESULT) = "Validation Result"
    arrResult(1, COL_PRODUCT_REVIEW) = "Product Review Flag"
    arrResult(1, COL_STUB_FLAG) = "Possible Stub Flag"
    arrResult(1, COL_PAYMENT_MATCH) = "Last Coupon vs Payment Match"
    arrResult(1, COL_PAYMENT_DIFF) = "Payment vs LHS Difference Days"
    arrResult(1, COL_PAYMENT_SELECTION) = "Payment Date Selection"
    arrResult(1, COL_LATEST_PAYMENT) = "Latest Neolink Payment Date"
    arrResult(1, COL_ACCRUAL_CONVENTION) = "Accrual Convention"
    arrResult(1, COL_ACCRUAL_PERIOD) = "Accrual Period"
    arrResult(1, COL_ACCRUAL_DAYS_SCHEDULE) = "Accrual Days - Schedule"
    arrResult(1, COL_ACCRUAL_DAYS_DIFF) = "Accrual Days Difference (LHS - Expected)"
    arrResult(1, COL_ACCRUAL_DAYS_LHS) = "Accrual Days - LHS Date"
    arrResult(1, COL_COUPON_PERIOD_DAYS) = "Coupon Period Days"
    arrResult(1, COL_YEAR_FRACTION) = "Year Fraction - Schedule"
    arrResult(1, COL_DAILY_ACCRUAL) = "Daily Accrual"
    arrResult(1, COL_EXPECTED_FULL_COUPON) = _
        "Simple Expected Full Coupon (Current LHS Rate)"
    arrResult(1, COL_ACCRUED_PERIOD_PCT) = "Accrued % of Coupon Period"
    arrResult(1, COL_ACCRUAL_REMAINING) = "Accrual Remaining to Next Coupon"
    arrResult(1, COL_EXPECTED_ACCRUAL) = "Expected Accrual - Schedule"
    arrResult(1, COL_EXPECTED_ACCRUAL_LHS) = "Expected Accrual - LHS Date"
    arrResult(1, COL_LHS_ACCRUED) = "LHS Accrued Interest ACC'S CCY"
    arrResult(1, COL_ACCRUAL_DIFF) = "Accrual Difference (LHS - Expected)"
    arrResult(1, COL_ACCRUAL_DIFF_PCT) = "Accrual Difference %"
    arrResult(1, COL_ACCRUAL_MATCH) = "Accrual Match"
    arrResult(1, COL_ACCRUAL_RESULT) = "Accrual Validation Result"

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

Private Function TryBuildExpectedCouponSchedule( _
    ByVal dblAccountingDate As Double, ByVal dblStartDate As Double, _
    ByVal dblFirstCoupon As Double, ByVal dblMaturityDate As Double, _
    ByVal vFrequency As Variant, ByRef dblExpectedLast As Double, _
    ByRef dblExpectedNext As Double, ByRef strStubFlag As String, _
    ByRef strReason As String) As Boolean

    Const METHOD_NAME As String = "TryBuildExpectedCouponSchedule"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngFrequency As Long
    Dim lngMonths As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    dblExpectedLast = 0
    dblExpectedNext = 0
    strStubFlag = vbNullString
    strReason = vbNullString

    If dblAccountingDate <= 0 Then
        strReason = "REVIEW - ACCOUNTING DATE MISSING"
        GoTo ExitPoint
    End If

    If Not TryGetLongCode(vFrequency, lngFrequency) Then
        strReason = "REVIEW - COUPON FREQUENCY MISSING"
        GoTo ExitPoint
    End If

    Select Case lngFrequency
        Case 1 To 4
            Call BuildCalendarEndSchedule( _
                CDate(dblAccountingDate), lngFrequency, dblExpectedLast, dblExpectedNext)
            strStubFlag = "CALENDAR-END FREQUENCY - REVIEW"
            TryBuildExpectedCouponSchedule = True

        Case 5 To 8
            Select Case lngFrequency
                Case 5: lngMonths = 1
                Case 6: lngMonths = 3
                Case 7: lngMonths = 6
                Case 8: lngMonths = 12
            End Select

            If dblMaturityDate <= 0 And dblFirstCoupon <= 0 Then
                strReason = "REVIEW - MATURITY/FIRST COUPON DATE MISSING"
                GoTo ExitPoint
            End If

            If Not BuildPeriodicSchedule( _
                CDate(dblAccountingDate), dblStartDate, dblFirstCoupon, _
                dblMaturityDate, lngMonths, dblExpectedLast, dblExpectedNext, _
                strStubFlag, strReason) Then
                GoTo ExitPoint
            End If

            TryBuildExpectedCouponSchedule = True

        Case 9
            strReason = "REVIEW - MATURITY-ONLY FREQUENCY"

        Case Else
            strReason = "REVIEW - UNSUPPORTED COUPON FREQUENCY"
    End Select

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

Private Function BuildPeriodicSchedule( _
    ByVal datAccounting As Date, ByVal dblStartDate As Double, _
    ByVal dblFirstCoupon As Double, ByVal dblMaturityDate As Double, _
    ByVal lngMonths As Long, ByRef dblExpectedLast As Double, _
    ByRef dblExpectedNext As Double, ByRef strStubFlag As String, _
    ByRef strReason As String) As Boolean

    Const METHOD_NAME As String = "BuildPeriodicSchedule"
    Dim datCurrent As Date
    Dim datFirstCoupon As Date
    Dim datMaturity As Date
    Dim datNext As Date
    Dim datStart As Date
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngMonthDifference As Long
    Dim lngSteps As Long
    Dim strFirstStub As String
    Dim strLastStub As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    ' Performance-critical routine: calculate the relevant coupon period directly
    ' with DateDiff/DateAdd. Do not walk the complete coupon history row by row.
    If dblFirstCoupon > 0 Then
        datFirstCoupon = CDate(dblFirstCoupon)

        If datAccounting < datFirstCoupon Then
            strReason = "REVIEW - ACCOUNTING DATE BEFORE FIRST COUPON"
            GoTo ExitPoint
        End If

        If dblMaturityDate > 0 Then
            datMaturity = CDate(dblMaturityDate)
            If datAccounting >= datMaturity Then
                dblExpectedLast = CDbl(datMaturity)
                dblExpectedNext = 0
                GoTo StubChecks
            End If
        End If

        lngMonthDifference = DateDiff("m", datFirstCoupon, datAccounting)
        lngSteps = lngMonthDifference \ lngMonths
        datCurrent = DateAdd("m", lngSteps * lngMonths, datFirstCoupon)

        If datCurrent > datAccounting Then
            lngSteps = lngSteps - 1
            datCurrent = DateAdd("m", lngSteps * lngMonths, datFirstCoupon)
        End If

        datNext = DateAdd("m", (lngSteps + 1) * lngMonths, datFirstCoupon)
        If datNext <= datAccounting Then
            lngSteps = lngSteps + 1
            datCurrent = datNext
            datNext = DateAdd("m", (lngSteps + 1) * lngMonths, datFirstCoupon)
        End If

        If dblMaturityDate > 0 And datNext > datMaturity Then datNext = datMaturity

        dblExpectedLast = CDbl(datCurrent)
        dblExpectedNext = CDbl(datNext)

StubChecks:
        If dblStartDate > 0 Then
            datStart = CDate(dblStartDate)
            If CLng(DateAdd("m", -lngMonths, datFirstCoupon)) <> CLng(datStart) Then
                strFirstStub = "POSSIBLE FIRST STUB"
            End If
        End If

        If dblMaturityDate > 0 Then
            strLastStub = DetectLastStub(datFirstCoupon, datMaturity, lngMonths)
        End If
    Else
        datMaturity = CDate(dblMaturityDate)

        If datAccounting >= datMaturity Then
            dblExpectedLast = CDbl(datMaturity)
            dblExpectedNext = 0
        Else
            lngMonthDifference = DateDiff("m", datAccounting, datMaturity)
            lngSteps = lngMonthDifference \ lngMonths
            datCurrent = DateAdd("m", -lngSteps * lngMonths, datMaturity)

            If datCurrent > datAccounting Then
                lngSteps = lngSteps + 1
                datCurrent = DateAdd("m", -lngSteps * lngMonths, datMaturity)
            End If

            datNext = DateAdd("m", -(lngSteps - 1) * lngMonths, datMaturity)

            If dblStartDate > 0 Then
                datStart = CDate(dblStartDate)
                If datCurrent < datStart Then
                    strReason = "REVIEW - ACCOUNTING DATE BEFORE FIRST REGULAR COUPON"
                    GoTo ExitPoint
                End If
            End If

            dblExpectedLast = CDbl(datCurrent)
            dblExpectedNext = CDbl(datNext)
        End If

        If dblStartDate > 0 Then
            strFirstStub = DetectStartMaturityMisalignment( _
                CDate(dblStartDate), datMaturity, lngMonths)
        End If
    End If

    strStubFlag = JoinFlags(strFirstStub, strLastStub)
    BuildPeriodicSchedule = True

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
Private Function DetectLastStub( _
    ByVal datFirstCoupon As Date, ByVal datMaturity As Date, _
    ByVal lngMonths As Long) As String

    Const METHOD_NAME As String = "DetectLastStub"
    Dim datCandidate As Date
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngMonthDifference As Long
    Dim lngSteps As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If datMaturity < datFirstCoupon Then
        DetectLastStub = "POSSIBLE LAST STUB"
        GoTo ExitPoint
    End If

    lngMonthDifference = DateDiff("m", datFirstCoupon, datMaturity)
    lngSteps = lngMonthDifference \ lngMonths
    datCandidate = DateAdd("m", lngSteps * lngMonths, datFirstCoupon)

    If CLng(datCandidate) <> CLng(datMaturity) Then
        DetectLastStub = "POSSIBLE LAST STUB"
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
Private Function DetectStartMaturityMisalignment( _
    ByVal datStart As Date, ByVal datMaturity As Date, _
    ByVal lngMonths As Long) As String

    Const METHOD_NAME As String = "DetectStartMaturityMisalignment"
    Dim datCandidate As Date
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngMonthDifference As Long
    Dim lngSteps As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If datMaturity < datStart Then
        DetectStartMaturityMisalignment = _
            "POSSIBLE STUB - START/MATURITY NOT ALIGNED"
        GoTo ExitPoint
    End If

    lngMonthDifference = DateDiff("m", datStart, datMaturity)
    lngSteps = lngMonthDifference \ lngMonths
    datCandidate = DateAdd("m", -lngSteps * lngMonths, datMaturity)

    If datCandidate > datStart Then
        lngSteps = lngSteps + 1
        datCandidate = DateAdd("m", -lngSteps * lngMonths, datMaturity)
    End If

    If CLng(datCandidate) <> CLng(datStart) Then
        DetectStartMaturityMisalignment = _
            "POSSIBLE STUB - START/MATURITY NOT ALIGNED"
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
Private Sub BuildCalendarEndSchedule( _
    ByVal datAccounting As Date, ByVal lngFrequency As Long, _
    ByRef dblExpectedLast As Double, ByRef dblExpectedNext As Double)

    Const METHOD_NAME As String = "BuildCalendarEndSchedule"
    Dim datPeriodEnd As Date
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngMonth As Long
    Dim lngPeriodMonths As Long
    Dim lngYear As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngYear = Year(datAccounting)
    lngMonth = Month(datAccounting)

    Select Case lngFrequency
        Case 1
            lngPeriodMonths = 1
            datPeriodEnd = DateSerial(lngYear, lngMonth + 1, 0)
        Case 2
            lngPeriodMonths = 3
            datPeriodEnd = DateSerial(lngYear, ((lngMonth - 1) \ 3 + 1) * 3 + 1, 0)
        Case 3
            lngPeriodMonths = 6
            If lngMonth <= 6 Then
                datPeriodEnd = DateSerial(lngYear, 7, 0)
            Else
                datPeriodEnd = DateSerial(lngYear + 1, 1, 0)
            End If
        Case 4
            lngPeriodMonths = 12
            datPeriodEnd = DateSerial(lngYear + 1, 1, 0)
    End Select

    If CLng(datAccounting) >= CLng(datPeriodEnd) Then
        dblExpectedLast = CDbl(datPeriodEnd)
        dblExpectedNext = CDbl(DateSerial( _
            Year(DateAdd("m", lngPeriodMonths, datPeriodEnd)), _
            Month(DateAdd("m", lngPeriodMonths, datPeriodEnd)) + 1, 0))
    Else
        dblExpectedNext = CDbl(datPeriodEnd)
        dblExpectedLast = CDbl(DateSerial( _
            Year(DateAdd("m", -lngPeriodMonths, datPeriodEnd)), _
            Month(DateAdd("m", -lngPeriodMonths, datPeriodEnd)) + 1, 0))
    End If

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

Private Function TryGetLongCode(ByVal vValue As Variant, ByRef lngCode As Long) As Boolean
    Const METHOD_NAME As String = "TryGetLongCode"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strValue As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngCode = 0
    If IsError(vValue) Or IsEmpty(vValue) Then GoTo ExitPoint

    strValue = Trim$(CStr(vValue))
    If Len(strValue) = 0 Then GoTo ExitPoint
    If Not IsNumeric(strValue) Then GoTo ExitPoint

    lngCode = CLng(strValue)
    TryGetLongCode = True

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

Private Function GetProductReviewFlag(ByVal vGtiName As Variant) As String
    Const METHOD_NAME As String = "GetProductReviewFlag"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strFlags As String
    Dim strName As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If IsError(vGtiName) Or IsEmpty(vGtiName) Then GoTo ExitPoint
    strName = UCase$(Trim$(CStr(vGtiName)))

    If InStr(1, strName, "MBS", vbTextCompare) > 0 Then
        strFlags = AppendReviewFlag(strFlags, "MBS")
    End If
    If InStr(1, strName, "ABS", vbTextCompare) > 0 Then
        strFlags = AppendReviewFlag(strFlags, "ABS")
    End If
    If InStr(1, strName, "AMORT", vbTextCompare) > 0 Then
        strFlags = AppendReviewFlag(strFlags, "AMORTIZING")
    End If
    If InStr(1, strName, "FLOAT", vbTextCompare) > 0 Or _
       InStr(1, strName, "FRN", vbTextCompare) > 0 Or _
       InStr(1, strName, "VARIABLE", vbTextCompare) > 0 Then
        strFlags = AppendReviewFlag(strFlags, "FLOATING/VARIABLE")
    End If
    If InStr(1, strName, "CALLABLE", vbTextCompare) > 0 Then
        strFlags = AppendReviewFlag(strFlags, "CALLABLE")
    End If

    If Len(strFlags) > 0 Then GetProductReviewFlag = "REVIEW - " & strFlags

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

Private Function AppendReviewFlag( _
    ByVal strExisting As String, ByVal strFlag As String) As String

    If Len(strExisting) = 0 Then
        AppendReviewFlag = strFlag
    Else
        AppendReviewFlag = strExisting & "; " & strFlag
    End If
End Function

Private Function BuildValidationResult( _
    ByVal strScheduleMatch As String, ByVal strProductReview As String, _
    ByVal strStubFlag As String, ByVal blnHasPayment As Boolean, _
    ByVal vPaymentShift As Variant) As String

    Const METHOD_NAME As String = "BuildValidationResult"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strReview As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strReview = JoinFlags(strProductReview, strStubFlag)

    If Len(strReview) > 0 Then
        If UCase$(strScheduleMatch) = "TRUE" Then
            BuildValidationResult = strReview & " / SCHEDULE MATCH"
        ElseIf UCase$(strScheduleMatch) = "FALSE" Then
            BuildValidationResult = strReview & " / SCHEDULE MISMATCH"
        Else
            BuildValidationResult = strReview
        End If
        GoTo ExitPoint
    End If

    If UCase$(strScheduleMatch) = "FALSE" Then
        BuildValidationResult = "LHS LAST COUPON MISMATCH"
    ElseIf UCase$(strScheduleMatch) = "TRUE" Then
        If Not blnHasPayment Then
            BuildValidationResult = "SCHEDULE OK - PAYMENT NOT FOUND"
        ElseIf IsNumeric(vPaymentShift) And CLng(vPaymentShift) <> 0 Then
            BuildValidationResult = "SCHEDULE OK - PAYMENT SHIFTED"
        Else
            BuildValidationResult = "SCHEDULE OK"
        End If
    Else
        BuildValidationResult = "REVIEW"
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

Private Function JoinFlags(ByVal strFirst As String, ByVal strSecond As String) As String
    Const METHOD_NAME As String = "JoinFlags"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Len(Trim$(strFirst)) > 0 And Len(Trim$(strSecond)) > 0 Then
        JoinFlags = Trim$(strFirst) & "; " & Trim$(strSecond)
    ElseIf Len(Trim$(strFirst)) > 0 Then
        JoinFlags = Trim$(strFirst)
    Else
        JoinFlags = Trim$(strSecond)
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
' Creation date: 2026-09-10
' Parameters:    arrNeolink - complete Input Neolink array including headers
' Returns:       Object - Fund + ISIN -> Array(selected Payment Date, account, Value Date,
'                previous-date-used flag, original latest Payment Date)
' Description:   Uses only GL:Type code = INTR. Tracks the latest and immediately
'                preceding distinct Payment Date for each Fund + ISIN. If the
'                preceding date is within 10 calendar days before the latest date,
'                the preceding date is selected; otherwise the latest date is used.
'                Account and Value Date always come from the selected Neolink row.
'-------------------------------------------------------------------------------
Private Function BuildPreferredNeolinkPaymentLookup(ByVal arrNeolink As Variant) As Object
    Const METHOD_NAME As String = "BuildPreferredNeolinkPaymentLookup"
    Dim arrExisting As Variant
    Dim arrKeys As Variant
    Dim blnHasValueDate As Boolean
    Dim dictPayment As Object
    Dim dblPaymentDate As Double
    Dim dblValueDate As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLatestDay As Long
    Dim lngPreviousDay As Long
    Dim lngRow As Long
    Dim lngKeyIndex As Long
    Dim strAccount As String
    Dim strKey As String
    Dim strTypeCode As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictPayment = CreateObject("Scripting.Dictionary")
    dictPayment.CompareMode = vbTextCompare

    For lngRow = 2 To UBound(arrNeolink, 1)
        strTypeCode = UCase$(Trim$(CStr(arrNeolink(lngRow, 3))))

        If strTypeCode = "INTR" Then
            strKey = BuildFundIsinKey(arrNeolink(lngRow, 1), arrNeolink(lngRow, 5))

            If Len(strKey) > 0 Then
                If TryGetExcelDateSerial(arrNeolink(lngRow, 7), dblPaymentDate) Then
                    ' Compare Payment Dates as calendar dates, not date-time values.
                    dblPaymentDate = CDbl(CLng(Int(dblPaymentDate)))
                    strAccount = GetLastElevenAccountCharacters(arrNeolink(lngRow, 2))
                    blnHasValueDate = TryGetExcelDateSerial( _
                        arrNeolink(lngRow, 6), dblValueDate)

                    If blnHasValueDate Then
                        dblValueDate = CDbl(CLng(Int(dblValueDate)))
                    Else
                        dblValueDate = 0
                    End If

                    If dictPayment.Exists(strKey) Then
                        arrExisting = dictPayment(strKey)

                        If dblPaymentDate > CDbl(arrExisting(0)) Then
                            ' The former latest date becomes the immediately preceding date.
                            arrExisting(3) = arrExisting(0)
                            arrExisting(4) = arrExisting(1)
                            arrExisting(5) = arrExisting(2)
                            arrExisting(0) = dblPaymentDate
                            arrExisting(1) = strAccount
                            arrExisting(2) = dblValueDate
                            dictPayment(strKey) = arrExisting
                        ElseIf dblPaymentDate < CDbl(arrExisting(0)) Then
                            If CDbl(arrExisting(3)) = 0 Or _
                               dblPaymentDate > CDbl(arrExisting(3)) Then

                                arrExisting(3) = dblPaymentDate
                                arrExisting(4) = strAccount
                                arrExisting(5) = dblValueDate
                                dictPayment(strKey) = arrExisting
                            End If
                        End If
                    Else
                        ' 0-2 = latest row; 3-5 = immediately preceding distinct row.
                        dictPayment.Add strKey, _
                            Array(dblPaymentDate, strAccount, dblValueDate, 0#, vbNullString, 0#)
                    End If
                End If
            End If
        End If
    Next lngRow

    ' Reduce the internal six-value record to the values expected by Output_All.
    If dictPayment.Count > 0 Then
        arrKeys = dictPayment.Keys

        For lngKeyIndex = LBound(arrKeys) To UBound(arrKeys)
            strKey = CStr(arrKeys(lngKeyIndex))
            arrExisting = dictPayment(strKey)
            lngLatestDay = CLng(arrExisting(0))
            lngPreviousDay = CLng(arrExisting(3))

            If lngPreviousDay > 0 And _
               lngLatestDay - lngPreviousDay <= PAYMENT_PREVIOUS_WINDOW_DAYS Then

                dictPayment(strKey) = _
                    Array(arrExisting(3), arrExisting(4), arrExisting(5), True, arrExisting(0))
            Else
                dictPayment(strKey) = _
                    Array(arrExisting(0), arrExisting(1), arrExisting(2), False, arrExisting(0))
            End If
        Next lngKeyIndex
    End If

    Set BuildPreferredNeolinkPaymentLookup = dictPayment

ExitPoint:
    If errNumber <> 0 Then
        Set dictPayment = Nothing
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

    Set rngMatch = wksOutputAll.Range( _
        wksOutputAll.Cells(2, COL_SCHEDULE_MATCH), _
        wksOutputAll.Cells(lngLastRow, COL_SCHEDULE_MATCH))
    Call AddTrueFalseFormatting(rngMatch)

    Set rngMatch = wksOutputAll.Range( _
        wksOutputAll.Cells(2, COL_PAYMENT_MATCH), _
        wksOutputAll.Cells(lngLastRow, COL_PAYMENT_MATCH))
    Call AddTrueFalseFormatting(rngMatch)

    Set rngMatch = wksOutputAll.Range( _
        wksOutputAll.Cells(2, COL_ACCRUAL_MATCH), _
        wksOutputAll.Cells(lngLastRow, COL_ACCRUAL_MATCH))
    Call AddTrueFalseFormatting(rngMatch)

    Set rngMatch = wksOutputAll.Range( _
        wksOutputAll.Cells(2, COL_CURRENT_FACE_MATCH), _
        wksOutputAll.Cells(lngLastRow, COL_CURRENT_FACE_MATCH))
    Call AddTrueFalseFormatting(rngMatch)

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

Private Sub AddTrueFalseFormatting(ByVal rngTarget As Excel.Range)
    Const METHOD_NAME As String = "AddTrueFalseFormatting"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strAddress As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    rngTarget.FormatConditions.Delete
    strAddress = rngTarget.Cells(1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=True)

    With rngTarget.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=" & strAddress & "=" & Chr$(34) & "TRUE" & Chr$(34))
        .Interior.Color = RGB(198, 239, 206)
    End With

    With rngTarget.FormatConditions.Add( _
        Type:=xlExpression, Formula1:="=" & strAddress & "=" & Chr$(34) & "FALSE" & Chr$(34))
        .Interior.Color = RGB(255, 199, 206)
    End With

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

Private Sub ApplyOutputAllDateHighlighting(ByVal wksOutputAll As Excel.Worksheet)
    Const METHOD_NAME As String = "ApplyOutputAllDateHighlighting"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLastRow As Long
    Dim rngLastCoupon As Excel.Range
    Dim rngPaymentDate As Excel.Range

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastRow = wksOutputAll.Cells(wksOutputAll.Rows.Count, 1).End(xlUp).Row

    With wksOutputAll.Cells(1, COL_LHS_LAST)
        .Font.Bold = True
        .Interior.Color = RGB(255, 246, 214)
    End With

    With wksOutputAll.Cells(1, COL_PAYMENT)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With

    With wksOutputAll.Cells(1, COL_EXPECTED_LAST)
        .Font.Bold = True
        .Interior.Color = RGB(226, 239, 218)
    End With

    If lngLastRow < 2 Then GoTo ExitPoint

    Set rngLastCoupon = wksOutputAll.Range( _
        wksOutputAll.Cells(2, COL_LHS_LAST), _
        wksOutputAll.Cells(lngLastRow, COL_LHS_LAST))
    Set rngPaymentDate = wksOutputAll.Range( _
        wksOutputAll.Cells(2, COL_PAYMENT), _
        wksOutputAll.Cells(lngLastRow, COL_PAYMENT))

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
    Dim rngValueDate As Excel.Range

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastRow = wksOutputAll.Cells(wksOutputAll.Rows.Count, 1).End(xlUp).Row
    If lngLastRow < 2 Then GoTo ExitPoint

    Set rngValueDate = wksOutputAll.Range( _
        wksOutputAll.Cells(2, COL_VALUE_DATE), _
        wksOutputAll.Cells(lngLastRow, COL_VALUE_DATE))
    arrValueDate = rngValueDate.Value2
    arrPaymentDate = wksOutputAll.Range( _
        wksOutputAll.Cells(2, COL_PAYMENT), _
        wksOutputAll.Cells(lngLastRow, COL_PAYMENT)).Value2

    rngValueDate.Interior.Pattern = xlNone

    For lngIndex = 1 To UBound(arrValueDate, 1)
        blnMismatch = False
        If TryGetExcelDateSerial(arrValueDate(lngIndex, 1), dblValueDate) Then
            If TryGetExcelDateSerial(arrPaymentDate(lngIndex, 1), dblPaymentDate) Then
                blnMismatch = (CLng(Int(dblValueDate)) <> CLng(Int(dblPaymentDate)))
            End If
        End If

        If blnMismatch Then
            rngValueDate.Cells(lngIndex, 1).Interior.Color = RGB(255, 230, 230)
        End If
    Next lngIndex

ExitPoint:
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

Private Sub ApplyPreviousPaymentSelectionFormatting(ByVal wksOutputAll As Excel.Worksheet)
    Const METHOD_NAME As String = "ApplyPreviousPaymentSelectionFormatting"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngRow As Long
    Dim lngLastRow As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastRow = wksOutputAll.Cells(wksOutputAll.Rows.Count, 1).End(xlUp).Row
    If lngLastRow < 2 Then GoTo ExitPoint

    With wksOutputAll.Range( _
        wksOutputAll.Cells(1, COL_PAYMENT_SELECTION), _
        wksOutputAll.Cells(1, COL_LATEST_PAYMENT))
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)
    End With

    For lngRow = 2 To lngLastRow
        If UCase$(Trim$(CStr( _
            wksOutputAll.Cells(lngRow, COL_PAYMENT_SELECTION).Value2))) = _
            "PREVIOUS DATE USED" Then

            wksOutputAll.Cells(lngRow, COL_ISIN).Interior.Color = RGB(255, 242, 204)
            wksOutputAll.Range( _
                wksOutputAll.Cells(lngRow, COL_PAYMENT_SELECTION), _
                wksOutputAll.Cells(lngRow, COL_LATEST_PAYMENT)).Interior.Color = _
                RGB(255, 242, 204)
        End If
    Next lngRow

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

    Set rngCouponFrequency = wksOutputAll.Range( _
        wksOutputAll.Cells(2, COL_COUPON_FREQ), _
        wksOutputAll.Cells(lngLastRow, COL_COUPON_FREQ))
    Set rngBondCalc = wksOutputAll.Range( _
        wksOutputAll.Cells(2, COL_BOND_CALC), _
        wksOutputAll.Cells(lngLastRow, COL_BOND_CALC))

    With rngCouponFrequency.Validation
        .Delete
        .Add Type:=xlValidateInputOnly
        .IgnoreBlank = True
        .InputTitle = "Coupon Frequency"
        .InputMessage = strCouponTooltip
        .ShowInput = True
        .ShowError = False
    End With

    With rngBondCalc.Validation
        .Delete
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

Private Sub ApplyWeekendDateHighlighting(ByVal wksOutputAll As Excel.Worksheet)
    Const METHOD_NAME As String = "ApplyWeekendDateHighlighting"
    Dim arrColumns As Variant
    Dim dblDate As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumnIndex As Long
    Dim lngLastRow As Long
    Dim lngRow As Long
    Dim lngWeekday As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastRow = wksOutputAll.Cells(wksOutputAll.Rows.Count, 1).End(xlUp).Row
    If lngLastRow < 2 Then GoTo ExitPoint

    arrColumns = Array(COL_LHS_NEXT, COL_LHS_LAST, COL_PAYMENT)

    For lngColumnIndex = LBound(arrColumns) To UBound(arrColumns)
        For lngRow = 2 To lngLastRow
            If TryGetExcelDateSerial( _
                wksOutputAll.Cells(lngRow, CLng(arrColumns(lngColumnIndex))).Value2, _
                dblDate) Then

                lngWeekday = Weekday(CDate(dblDate), vbSunday)
                If lngWeekday = vbSaturday Then
                    wksOutputAll.Cells( _
                        lngRow, CLng(arrColumns(lngColumnIndex))).Interior.Color = _
                        RGB(226, 239, 218)
                ElseIf lngWeekday = vbSunday Then
                    wksOutputAll.Cells( _
                        lngRow, CLng(arrColumns(lngColumnIndex))).Interior.Color = _
                        RGB(252, 228, 214)
                End If
            End If
        Next lngRow
    Next lngColumnIndex

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

Private Sub PopulateAccrualDiagnostics( _
    ByRef arrResult As Variant, ByVal lngOutputRow As Long, _
    ByVal arrLHS As Variant, ByVal lngLHSRow As Long, _
    ByVal blnScheduleBuilt As Boolean, ByVal dblAccountingDate As Double, _
    ByVal dblExpectedLast As Double, ByVal dblExpectedNext As Double, _
    ByVal strProductReview As String, ByVal strStubFlag As String)

    Const METHOD_NAME As String = "PopulateAccrualDiagnostics"
    Dim blnHasLhsAccrued As Boolean
    Dim blnHasLhsLast As Boolean
    Dim blnHasLhsNext As Boolean
    Dim blnHasQuantity As Boolean
    Dim blnHasRate As Boolean
    Dim blnModelCalculated As Boolean
    Dim dblDailyAccrual As Double
    Dim dblExpectedAccrual As Double
    Dim dblExpectedAccrualLhs As Double
    Dim dblLhsAccrued As Double
    Dim dblLhsDaysValue As Double
    Dim dblLhsLast As Double
    Dim dblLhsNext As Double
    Dim dblQuantity As Double
    Dim dblRatePercent As Double
    Dim dblYearFraction As Double
    Dim dblYearFractionLhs As Double
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngAccrualDays As Long
    Dim lngAccrualDaysLhs As Long
    Dim lngBondCalc As Long
    Dim lngCouponPeriodDays As Long
    Dim lngCouponPeriodDaysLhs As Long
    Dim lngFrequency As Long
    Dim lngLhsDays As Long
    Dim strCalcReason As String
    Dim strMatch As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrResult(lngOutputRow, COL_ACCRUAL_CONVENTION) = _
        GetAccrualConventionDescription(arrLHS(lngLHSRow, 10))
    arrResult(lngOutputRow, COL_ACCRUAL_PERIOD) = _
        GetAccrualPeriodDescription(arrLHS(lngLHSRow, 9))

    If Not TryGetLongCode(arrLHS(lngLHSRow, 10), lngBondCalc) Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
            "REVIEW - BOND CALC METHOD MISSING"
        GoTo ExitPoint
    End If

    If lngBondCalc <> VERIFIED_ACCRUAL_METHOD Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
            "REVIEW - METHOD NOT VERIFIED"
        GoTo ExitPoint
    End If

    If Not blnScheduleBuilt Or dblExpectedLast <= 0 Or dblExpectedNext <= 0 Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
            "REVIEW - EXPECTED COUPON SCHEDULE NOT AVAILABLE"
        GoTo ExitPoint
    End If

    If UBound(arrLHS, 2) >= 19 Then
        blnHasQuantity = TryGetDoubleValue(arrLHS(lngLHSRow, 19), dblQuantity)
    End If
    blnHasRate = TryGetDoubleValue(arrLHS(lngLHSRow, 12), dblRatePercent)
    blnHasLhsAccrued = TryGetDoubleValue(arrLHS(lngLHSRow, 13), dblLhsAccrued)

    If Not blnHasQuantity Or dblQuantity = 0 Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
            "REVIEW - QUANTITY MISSING OR ZERO"
        GoTo ExitPoint
    End If

    If Not blnHasRate Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
            "REVIEW - INTEREST RATE MISSING"
        GoTo ExitPoint
    End If

    If Not TryGetLongCode(arrLHS(lngLHSRow, 9), lngFrequency) Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
            "REVIEW - COUPON FREQUENCY MISSING"
        GoTo ExitPoint
    End If

    blnModelCalculated = TryCalculateMethod1Accrual( _
        dblQuantity, dblRatePercent, lngFrequency, dblAccountingDate, _
        dblExpectedLast, dblExpectedNext, lngAccrualDays, lngCouponPeriodDays, _
        dblYearFraction, dblDailyAccrual, dblExpectedAccrual, strCalcReason)

    If Not blnModelCalculated Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = strCalcReason
        GoTo ExitPoint
    End If

    arrResult(lngOutputRow, COL_ACCRUAL_DAYS_SCHEDULE) = lngAccrualDays
    arrResult(lngOutputRow, COL_COUPON_PERIOD_DAYS) = lngCouponPeriodDays
    arrResult(lngOutputRow, COL_YEAR_FRACTION) = dblYearFraction
    arrResult(lngOutputRow, COL_DAILY_ACCRUAL) = dblDailyAccrual
    arrResult(lngOutputRow, COL_EXPECTED_FULL_COUPON) = _
        dblDailyAccrual * CDbl(lngCouponPeriodDays)
    arrResult(lngOutputRow, COL_ACCRUED_PERIOD_PCT) = _
        CDbl(lngAccrualDays) / CDbl(lngCouponPeriodDays)
    arrResult(lngOutputRow, COL_ACCRUAL_REMAINING) = _
        (dblDailyAccrual * CDbl(lngCouponPeriodDays)) - dblExpectedAccrual
    arrResult(lngOutputRow, COL_EXPECTED_ACCRUAL) = dblExpectedAccrual

    If TryGetDoubleValue(arrLHS(lngLHSRow, 5), dblLhsDaysValue) Then
        lngLhsDays = CLng(dblLhsDaysValue)
        arrResult(lngOutputRow, COL_ACCRUAL_DAYS_DIFF) = lngLhsDays - lngAccrualDays
    End If

    blnHasLhsLast = TryGetExcelDateSerial(arrLHS(lngLHSRow, 3), dblLhsLast)
    blnHasLhsNext = TryGetExcelDateSerial(arrLHS(lngLHSRow, 4), dblLhsNext)

    If blnHasLhsLast And blnHasLhsNext Then
        strCalcReason = vbNullString
        If TryCalculateMethod1Accrual( _
            dblQuantity, arrLHS(lngLHSRow, 12), lngFrequency, dblAccountingDate, _
            dblLhsLast, dblLhsNext, lngAccrualDaysLhs, lngCouponPeriodDaysLhs, _
            dblYearFractionLhs, dblDailyAccrual, dblExpectedAccrualLhs, strCalcReason) Then

            arrResult(lngOutputRow, COL_ACCRUAL_DAYS_LHS) = lngAccrualDaysLhs
            arrResult(lngOutputRow, COL_EXPECTED_ACCRUAL_LHS) = dblExpectedAccrualLhs
        End If
    End If

    If Not blnHasLhsAccrued Then
        arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = "REVIEW"
        arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = _
            "REVIEW - LHS ACCRUED INTEREST MISSING"
        GoTo ExitPoint
    End If

    arrResult(lngOutputRow, COL_LHS_ACCRUED) = dblLhsAccrued
    arrResult(lngOutputRow, COL_ACCRUAL_DIFF) = dblLhsAccrued - dblExpectedAccrual

    If Abs(dblExpectedAccrual) > 0.0000001 Then
        arrResult(lngOutputRow, COL_ACCRUAL_DIFF_PCT) = _
            (dblLhsAccrued - dblExpectedAccrual) / dblExpectedAccrual
    End If

    If Abs(dblLhsAccrued - dblExpectedAccrual) <= ACCRUAL_MATCH_TOLERANCE Then
        strMatch = "TRUE"
    Else
        strMatch = "FALSE"
    End If

    arrResult(lngOutputRow, COL_ACCRUAL_MATCH) = strMatch
    arrResult(lngOutputRow, COL_ACCRUAL_RESULT) = BuildAccrualValidationResult( _
        strMatch, strProductReview, strStubFlag)

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
        "lngLHSRow", lngLHSRow)
    GoTo ExitPoint
End Sub

Private Function TryCalculateMethod1Accrual( _
    ByVal dblQuantity As Double, ByVal vRatePercent As Variant, _
    ByVal lngFrequency As Long, ByVal dblAccountingDate As Double, _
    ByVal dblPeriodLast As Double, ByVal dblPeriodNext As Double, _
    ByRef lngAccrualDays As Long, ByRef lngCouponPeriodDays As Long, _
    ByRef dblYearFraction As Double, ByRef dblDailyAccrual As Double, _
    ByRef dblExpectedAccrual As Double, ByRef strReason As String) As Boolean

    Const METHOD_NAME As String = "TryCalculateMethod1Accrual"
    Dim dblCouponFactor As Double
    Dim dblCouponAmount As Double
    Dim dblRatePercent As Double
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngAccrualDays = 0
    lngCouponPeriodDays = 0
    dblYearFraction = 0
    dblDailyAccrual = 0
    dblExpectedAccrual = 0
    strReason = vbNullString

    If dblAccountingDate <= 0 Or dblPeriodLast <= 0 Or dblPeriodNext <= 0 Then
        strReason = "REVIEW - ACCRUAL DATES MISSING"
        GoTo ExitPoint
    End If

    If Not TryGetDoubleValue(vRatePercent, dblRatePercent) Then
        strReason = "REVIEW - INTEREST RATE MISSING"
        GoTo ExitPoint
    End If

    If Not TryGetCouponFrequencyFactor(lngFrequency, dblCouponFactor) Then
        strReason = "REVIEW - ACCRUAL FREQUENCY NOT SUPPORTED"
        GoTo ExitPoint
    End If

    lngCouponPeriodDays = CLng(Int(dblPeriodNext)) - CLng(Int(dblPeriodLast))
    If lngCouponPeriodDays <= 0 Then
        strReason = "REVIEW - INVALID COUPON PERIOD"
        GoTo ExitPoint
    End If

    lngAccrualDays = CLng(Int(dblAccountingDate)) - CLng(Int(dblPeriodLast)) + 1
    If lngAccrualDays < 0 Or lngAccrualDays > lngCouponPeriodDays Then
        strReason = "REVIEW - ACCOUNTING DATE OUTSIDE COUPON PERIOD"
        GoTo ExitPoint
    End If

    ' LHS stores the rate in percentage points, e.g. 4.467 means 4.467%.
    dblCouponAmount = dblQuantity * (dblRatePercent / 100#) * dblCouponFactor
    dblYearFraction = dblCouponFactor * CDbl(lngAccrualDays) / _
        CDbl(lngCouponPeriodDays)
    dblDailyAccrual = dblCouponAmount / CDbl(lngCouponPeriodDays)
    dblExpectedAccrual = dblQuantity * (dblRatePercent / 100#) * dblYearFraction
    TryCalculateMethod1Accrual = True

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

Private Function TryGetCouponFrequencyFactor( _
    ByVal lngFrequency As Long, ByRef dblFactor As Double) As Boolean

    Const METHOD_NAME As String = "TryGetCouponFrequencyFactor"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    dblFactor = 0

    Select Case lngFrequency
        Case 1, 5
            dblFactor = 1# / 12#
        Case 2, 6
            dblFactor = 1# / 4#
        Case 3, 7
            dblFactor = 1# / 2#
        Case 4, 8
            dblFactor = 1#
        Case Else
            GoTo ExitPoint
    End Select

    TryGetCouponFrequencyFactor = True

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

Private Function GetAccrualConventionDescription(ByVal vMethod As Variant) As String
    Const METHOD_NAME As String = "GetAccrualConventionDescription"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngMethod As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not TryGetLongCode(vMethod, lngMethod) Then
        GetAccrualConventionDescription = "(missing)"
        GoTo ExitPoint
    End If

    Select Case lngMethod
        Case 1: GetAccrualConventionDescription = "1 - 365-6/365"
        Case 2: GetAccrualConventionDescription = "2 - 365-6/360"
        Case 3: GetAccrualConventionDescription = "3 - 360/360"
        Case 4: GetAccrualConventionDescription = "4 - 360/365"
        Case 5: GetAccrualConventionDescription = "5 - 365-6/365-6"
        Case 6: GetAccrualConventionDescription = "6 - 365-6/365-6 = Civil"
        Case 7: GetAccrualConventionDescription = "7 - 360/360 US"
        Case 8: GetAccrualConventionDescription = "8 - Actual/364"
        Case 9: GetAccrualConventionDescription = "9 - Actual/252"
        Case Else: GetAccrualConventionDescription = CStr(lngMethod) & " - UNKNOWN"
    End Select

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

Private Function GetAccrualPeriodDescription(ByVal vFrequency As Variant) As String
    Const METHOD_NAME As String = "GetAccrualPeriodDescription"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngFrequency As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If Not TryGetLongCode(vFrequency, lngFrequency) Then
        GetAccrualPeriodDescription = "(missing)"
        GoTo ExitPoint
    End If

    Select Case lngFrequency
        Case 1: GetAccrualPeriodDescription = "End of Month"
        Case 2: GetAccrualPeriodDescription = "End of Quarter"
        Case 3: GetAccrualPeriodDescription = "End of half-year"
        Case 4: GetAccrualPeriodDescription = "End of Year"
        Case 5: GetAccrualPeriodDescription = "Monthly"
        Case 6: GetAccrualPeriodDescription = "Quarterly"
        Case 7: GetAccrualPeriodDescription = "Half-year"
        Case 8: GetAccrualPeriodDescription = "Year"
        Case 9: GetAccrualPeriodDescription = "Maturity"
        Case Else: GetAccrualPeriodDescription = CStr(lngFrequency) & " - UNKNOWN"
    End Select

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

Private Function BuildAccrualValidationResult( _
    ByVal strMatch As String, ByVal strProductReview As String, _
    ByVal strStubFlag As String) As String

    Const METHOD_NAME As String = "BuildAccrualValidationResult"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strReview As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strReview = JoinFlags(strProductReview, strStubFlag)

    If Len(strReview) > 0 Then
        If UCase$(strMatch) = "TRUE" Then
            BuildAccrualValidationResult = _
                strReview & " / SIMPLE COUPON MODEL / ACCRUAL MATCH"
        Else
            BuildAccrualValidationResult = _
                strReview & " / SIMPLE COUPON MODEL / ACCRUAL MISMATCH"
        End If
    ElseIf UCase$(strMatch) = "TRUE" Then
        BuildAccrualValidationResult = "OK - METHOD 1 MODEL"
    Else
        BuildAccrualValidationResult = "ACCRUAL MISMATCH - METHOD 1 MODEL"
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

Private Function TryGetDoubleValue( _
    ByVal vValue As Variant, ByRef dblValue As Double) As Boolean

    Const METHOD_NAME As String = "TryGetDoubleValue"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strValue As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    dblValue = 0
    If IsError(vValue) Or IsEmpty(vValue) Then GoTo ExitPoint

    If IsNumeric(vValue) Then
        dblValue = CDbl(vValue)
        TryGetDoubleValue = True
        GoTo ExitPoint
    End If

    strValue = Trim$(CStr(vValue))
    If Len(strValue) = 0 Then GoTo ExitPoint
    If IsNumeric(strValue) Then
        dblValue = CDbl(strValue)
        TryGetDoubleValue = True
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

Private Sub SetOutputAllNumberFormats(ByVal wksOutputAll As Excel.Worksheet)
    Const METHOD_NAME As String = "SetOutputAllNumberFormats"
    Dim arrDateColumns As Variant
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngIndex As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    arrDateColumns = Array( _
        COL_ACCOUNTING_DATE, COL_GENERATION_DATE, COL_START_DATE, COL_FIRST_COUPON, _
        COL_MATURITY, COL_VALUE_DATE, COL_LHS_NEXT, COL_EXPECTED_NEXT, _
        COL_LHS_LAST, COL_EXPECTED_LAST, COL_PAYMENT, COL_LATEST_PAYMENT)

    For lngIndex = LBound(arrDateColumns) To UBound(arrDateColumns)
        wksOutputAll.Columns(CLng(arrDateColumns(lngIndex))).NumberFormat = "dd/mm/yyyy"
    Next lngIndex

    wksOutputAll.Columns(COL_PAYMENT_SHIFT).NumberFormat = "0"
    wksOutputAll.Columns(COL_PAYMENT_DIFF).NumberFormat = "0"
    wksOutputAll.Columns(COL_QUANTITY).NumberFormat = "#,##0.0000"
    wksOutputAll.Columns(COL_FACE_VALUE).NumberFormat = "#,##0.00"
    wksOutputAll.Columns(COL_REIMBURSEMENT_FACTOR).NumberFormat = "0.000000000"
    wksOutputAll.Columns(COL_CALCULATED_CURRENT_FACE).NumberFormat = "#,##0.0000"
    wksOutputAll.Columns(COL_INTEREST_RATE).NumberFormat = "0.000000"
    wksOutputAll.Columns(COL_ACCRUAL_DAYS_SCHEDULE).NumberFormat = "0"
    wksOutputAll.Columns(COL_ACCRUAL_DAYS_DIFF).NumberFormat = "0"
    wksOutputAll.Columns(COL_ACCRUAL_DAYS_LHS).NumberFormat = "0"
    wksOutputAll.Columns(COL_COUPON_PERIOD_DAYS).NumberFormat = "0"
    wksOutputAll.Columns(COL_YEAR_FRACTION).NumberFormat = "0.00000000"
    wksOutputAll.Columns(COL_DAILY_ACCRUAL).NumberFormat = "#,##0.0000"
    wksOutputAll.Columns(COL_EXPECTED_FULL_COUPON).NumberFormat = "#,##0.00"
    wksOutputAll.Columns(COL_ACCRUED_PERIOD_PCT).NumberFormat = "0.00%"
    wksOutputAll.Columns(COL_ACCRUAL_REMAINING).NumberFormat = "#,##0.00"
    wksOutputAll.Columns(COL_EXPECTED_ACCRUAL).NumberFormat = "#,##0.00"
    wksOutputAll.Columns(COL_EXPECTED_ACCRUAL_LHS).NumberFormat = "#,##0.00"
    wksOutputAll.Columns(COL_LHS_ACCRUED).NumberFormat = "#,##0.00"
    wksOutputAll.Columns(COL_ACCRUAL_DIFF).NumberFormat = "#,##0.00"
    wksOutputAll.Columns(COL_ACCRUAL_DIFF_PCT).NumberFormat = "0.0000%"

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

Private Sub BuildAccrualAnalysisSheet( _
    ByVal wksAnalysis As Excel.Worksheet, ByVal arrOutputAll As Variant)

    Const METHOD_NAME As String = "BuildAccrualAnalysisSheet"
    Dim arrGtiKeys As Variant
    Dim arrMethodSummary() As Variant
    Dim arrGtiSummary() As Variant
    Dim chtObject As Excel.ChartObject
    Dim dictGtiDiffCount As Object
    Dim dictGtiDiffSum As Object
    Dim dictGtiMismatch As Object
    Dim dictGtiReview As Object
    Dim dictGtiTotal As Object
    Dim dictMethodDiffCount As Object
    Dim dictMethodDiffSum As Object
    Dim dictMethodFalse As Object
    Dim dictMethodReview As Object
    Dim dictMethodTotal As Object
    Dim dictMethodTrue As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngCalculated As Long
    Dim lngChartRows As Long
    Dim lngFalseTotal As Long
    Dim lngGtiRows As Long
    Dim lngIndex As Long
    Dim lngMethodRows As Long
    Dim lngReviewTotal As Long
    Dim lngRow As Long
    Dim lngTrueTotal As Long
    Dim strGti As String
    Dim strMatch As String
    Dim strMethod As String
    Dim vKey As Variant

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictMethodTotal = CreateObject("Scripting.Dictionary")
    Set dictMethodTrue = CreateObject("Scripting.Dictionary")
    Set dictMethodFalse = CreateObject("Scripting.Dictionary")
    Set dictMethodReview = CreateObject("Scripting.Dictionary")
    Set dictMethodDiffSum = CreateObject("Scripting.Dictionary")
    Set dictMethodDiffCount = CreateObject("Scripting.Dictionary")
    Set dictGtiTotal = CreateObject("Scripting.Dictionary")
    Set dictGtiMismatch = CreateObject("Scripting.Dictionary")
    Set dictGtiReview = CreateObject("Scripting.Dictionary")
    Set dictGtiDiffSum = CreateObject("Scripting.Dictionary")
    Set dictGtiDiffCount = CreateObject("Scripting.Dictionary")

    dictMethodTotal.CompareMode = vbTextCompare
    dictMethodTrue.CompareMode = vbTextCompare
    dictMethodFalse.CompareMode = vbTextCompare
    dictMethodReview.CompareMode = vbTextCompare
    dictMethodDiffSum.CompareMode = vbTextCompare
    dictMethodDiffCount.CompareMode = vbTextCompare
    dictGtiTotal.CompareMode = vbTextCompare
    dictGtiMismatch.CompareMode = vbTextCompare
    dictGtiReview.CompareMode = vbTextCompare
    dictGtiDiffSum.CompareMode = vbTextCompare
    dictGtiDiffCount.CompareMode = vbTextCompare

    For lngRow = 2 To UBound(arrOutputAll, 1)
        strMethod = Trim$(CStr(arrOutputAll(lngRow, COL_ACCRUAL_CONVENTION)))
        If Len(strMethod) = 0 Then strMethod = "(missing)"

        strGti = Trim$(CStr(arrOutputAll(lngRow, COL_GTI_NAME)))
        If Len(strGti) = 0 Then strGti = "(blank GTI NAME)"

        strMatch = UCase$(Trim$(CStr(arrOutputAll(lngRow, COL_ACCRUAL_MATCH))))

        Call IncrementDictionaryCount(dictMethodTotal, strMethod)
        Call IncrementDictionaryCount(dictGtiTotal, strGti)

        Select Case strMatch
            Case "TRUE"
                Call IncrementDictionaryCount(dictMethodTrue, strMethod)
                lngTrueTotal = lngTrueTotal + 1
                lngCalculated = lngCalculated + 1
            Case "FALSE"
                Call IncrementDictionaryCount(dictMethodFalse, strMethod)
                Call IncrementDictionaryCount(dictGtiMismatch, strGti)
                lngFalseTotal = lngFalseTotal + 1
                lngCalculated = lngCalculated + 1
            Case Else
                Call IncrementDictionaryCount(dictMethodReview, strMethod)
                Call IncrementDictionaryCount(dictGtiReview, strGti)
                lngReviewTotal = lngReviewTotal + 1
        End Select

        If IsNumeric(arrOutputAll(lngRow, COL_ACCRUAL_DIFF)) Then
            Call AddDictionaryDouble( _
                dictMethodDiffSum, strMethod, _
                Abs(CDbl(arrOutputAll(lngRow, COL_ACCRUAL_DIFF))))
            Call IncrementDictionaryCount(dictMethodDiffCount, strMethod)

            Call AddDictionaryDouble( _
                dictGtiDiffSum, strGti, _
                Abs(CDbl(arrOutputAll(lngRow, COL_ACCRUAL_DIFF))))
            Call IncrementDictionaryCount(dictGtiDiffCount, strGti)
        End If
    Next lngRow

    wksAnalysis.Range("A1").Value2 = "Accrual validation overview"
    wksAnalysis.Range("A2").Value2 = "Output_All rows"
    wksAnalysis.Range("B2").Value2 = UBound(arrOutputAll, 1) - 1
    wksAnalysis.Range("A3").Value2 = "Method 1 rows calculated"
    wksAnalysis.Range("B3").Value2 = lngCalculated
    wksAnalysis.Range("A4").Value2 = "Accrual matches"
    wksAnalysis.Range("B4").Value2 = lngTrueTotal
    wksAnalysis.Range("A5").Value2 = "Accrual mismatches"
    wksAnalysis.Range("B5").Value2 = lngFalseTotal
    wksAnalysis.Range("A6").Value2 = "Review / unverified"
    wksAnalysis.Range("B6").Value2 = lngReviewTotal
    wksAnalysis.Range("A7").Value2 = "Match tolerance"
    wksAnalysis.Range("B7").Value2 = ACCRUAL_MATCH_TOLERANCE

    wksAnalysis.Range("A9").Value2 = "Accrual Convention"
    wksAnalysis.Range("B9").Value2 = "Rows"
    wksAnalysis.Range("C9").Value2 = "TRUE"
    wksAnalysis.Range("D9").Value2 = "FALSE"
    wksAnalysis.Range("E9").Value2 = "Review"
    wksAnalysis.Range("F9").Value2 = "Average absolute accrual div"

    lngMethodRows = dictMethodTotal.Count
    If lngMethodRows > 0 Then
        ReDim arrMethodSummary(1 To lngMethodRows, 1 To 6)
        lngIndex = 0

        For Each vKey In dictMethodTotal.Keys
            lngIndex = lngIndex + 1
            arrMethodSummary(lngIndex, 1) = CStr(vKey)
            arrMethodSummary(lngIndex, 2) = GetDictionaryCount(dictMethodTotal, CStr(vKey))
            arrMethodSummary(lngIndex, 3) = GetDictionaryCount(dictMethodTrue, CStr(vKey))
            arrMethodSummary(lngIndex, 4) = GetDictionaryCount(dictMethodFalse, CStr(vKey))
            arrMethodSummary(lngIndex, 5) = GetDictionaryCount(dictMethodReview, CStr(vKey))

            If GetDictionaryCount(dictMethodDiffCount, CStr(vKey)) > 0 Then
                arrMethodSummary(lngIndex, 6) = _
                    GetDictionaryDouble(dictMethodDiffSum, CStr(vKey)) / _
                    GetDictionaryCount(dictMethodDiffCount, CStr(vKey))
            End If
        Next vKey

        wksAnalysis.Range("A10").Resize(lngMethodRows, 6).Value2 = arrMethodSummary
        wksAnalysis.Range("F10").Resize(lngMethodRows, 1).NumberFormat = "#,##0.00"
    End If

    wksAnalysis.Range("H9").Value2 = "GTI NAME"
    wksAnalysis.Range("I9").Value2 = "Rows"
    wksAnalysis.Range("J9").Value2 = "Accrual mismatch count"
    wksAnalysis.Range("K9").Value2 = "Review count"
    wksAnalysis.Range("L9").Value2 = "Average absolute accrual div"

    arrGtiKeys = BuildCombinedThreeKeyArray( _
        dictGtiMismatch, dictGtiReview, dictGtiTotal)

    If Not IsEmpty(arrGtiKeys) Then
        Call SortKeysByThreeCounts( _
            arrGtiKeys, dictGtiMismatch, dictGtiReview, dictGtiTotal)
        lngGtiRows = UBound(arrGtiKeys) - LBound(arrGtiKeys) + 1
        ReDim arrGtiSummary(1 To lngGtiRows, 1 To 5)

        For lngIndex = LBound(arrGtiKeys) To UBound(arrGtiKeys)
            strGti = CStr(arrGtiKeys(lngIndex))
            arrGtiSummary(lngIndex - LBound(arrGtiKeys) + 1, 1) = strGti
            arrGtiSummary(lngIndex - LBound(arrGtiKeys) + 1, 2) = _
                GetDictionaryCount(dictGtiTotal, strGti)
            arrGtiSummary(lngIndex - LBound(arrGtiKeys) + 1, 3) = _
                GetDictionaryCount(dictGtiMismatch, strGti)
            arrGtiSummary(lngIndex - LBound(arrGtiKeys) + 1, 4) = _
                GetDictionaryCount(dictGtiReview, strGti)

            If GetDictionaryCount(dictGtiDiffCount, strGti) > 0 Then
                arrGtiSummary(lngIndex - LBound(arrGtiKeys) + 1, 5) = _
                    GetDictionaryDouble(dictGtiDiffSum, strGti) / _
                    GetDictionaryCount(dictGtiDiffCount, strGti)
            End If
        Next lngIndex

        wksAnalysis.Range("H10").Resize(lngGtiRows, 5).Value2 = arrGtiSummary
        wksAnalysis.Range("L10").Resize(lngGtiRows, 1).NumberFormat = "#,##0.00"

        For lngIndex = LBound(arrGtiKeys) To UBound(arrGtiKeys)
            If GetDictionaryCount( _
                dictGtiMismatch, CStr(arrGtiKeys(lngIndex))) > 0 Then
                lngChartRows = lngChartRows + 1
                If lngChartRows = 15 Then Exit For
            Else
                Exit For
            End If
        Next lngIndex

        If lngChartRows > 0 Then
            Set chtObject = wksAnalysis.ChartObjects.Add( _
                Left:=wksAnalysis.Range("N2").Left, _
                Top:=wksAnalysis.Range("N2").Top, Width:=650, Height:=380)

            With chtObject.Chart
                .ChartType = xlBarClustered
                .SeriesCollection.NewSeries
                With .SeriesCollection(1)
                    .Name = "Accrual mismatch count"
                    .XValues = wksAnalysis.Range( _
                        wksAnalysis.Cells(10, 8), _
                        wksAnalysis.Cells(9 + lngChartRows, 8))
                    .Values = wksAnalysis.Range( _
                        wksAnalysis.Cells(10, 10), _
                        wksAnalysis.Cells(9 + lngChartRows, 10))
                    .ApplyDataLabels
                End With
                .HasTitle = True
                .ChartTitle.Text = "Accrual mismatches by GTI Name - top 15"
                .HasLegend = False
            End With
        End If
    End If

    wksAnalysis.Range("A1:F1").Font.Bold = True
    wksAnalysis.Range("A9:F9").Font.Bold = True
    wksAnalysis.Range("H9:L9").Font.Bold = True

ExitPoint:
    Set chtObject = Nothing
    Set dictMethodTotal = Nothing
    Set dictMethodTrue = Nothing
    Set dictMethodFalse = Nothing
    Set dictMethodReview = Nothing
    Set dictMethodDiffSum = Nothing
    Set dictMethodDiffCount = Nothing
    Set dictGtiTotal = Nothing
    Set dictGtiMismatch = Nothing
    Set dictGtiReview = Nothing
    Set dictGtiDiffSum = Nothing
    Set dictGtiDiffCount = Nothing

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

Private Sub AddDictionaryDouble( _
    ByVal dictTarget As Object, ByVal strKey As String, ByVal dblValue As Double)

    Const METHOD_NAME As String = "AddDictionaryDouble"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If dictTarget.Exists(strKey) Then
        dictTarget(strKey) = CDbl(dictTarget(strKey)) + dblValue
    Else
        dictTarget.Add strKey, dblValue
    End If

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

Private Function GetDictionaryDouble( _
    ByVal dictTarget As Object, ByVal strKey As String) As Double

    Const METHOD_NAME As String = "GetDictionaryDouble"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If dictTarget.Exists(strKey) Then
        GetDictionaryDouble = CDbl(dictTarget(strKey))
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

Private Sub BuildAnalysisSheet( _
    ByVal wksAnalysis As Excel.Worksheet, ByVal arrOutputAll As Variant)

    Const METHOD_NAME As String = "BuildAnalysisSheet"
    Dim arrKeys As Variant
    Dim arrSummary() As Variant
    Dim chtObject As Excel.ChartObject
    Dim dictConfirmed As Object
    Dim dictRawMismatch As Object
    Dim dictReview As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngChartRows As Long
    Dim lngConfirmedTotal As Long
    Dim lngIndex As Long
    Dim lngRawMismatchTotal As Long
    Dim lngReviewTotal As Long
    Dim lngRow As Long
    Dim lngSummaryRows As Long
    Dim strGti As String
    Dim strResult As String
    Dim strScheduleMatch As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictRawMismatch = CreateObject("Scripting.Dictionary")
    Set dictConfirmed = CreateObject("Scripting.Dictionary")
    Set dictReview = CreateObject("Scripting.Dictionary")
    dictRawMismatch.CompareMode = vbTextCompare
    dictConfirmed.CompareMode = vbTextCompare
    dictReview.CompareMode = vbTextCompare

    For lngRow = 2 To UBound(arrOutputAll, 1)
        strGti = Trim$(CStr(arrOutputAll(lngRow, COL_GTI_NAME)))
        If Len(strGti) = 0 Then strGti = "(blank GTI NAME)"

        strScheduleMatch = UCase$(Trim$(CStr( _
            arrOutputAll(lngRow, COL_SCHEDULE_MATCH))))
        strResult = UCase$(Trim$(CStr( _
            arrOutputAll(lngRow, COL_VALIDATION_RESULT))))

        If strScheduleMatch = "FALSE" Then
            Call IncrementDictionaryCount(dictRawMismatch, strGti)
            lngRawMismatchTotal = lngRawMismatchTotal + 1
        End If

        If strResult = "LHS LAST COUPON MISMATCH" Then
            Call IncrementDictionaryCount(dictConfirmed, strGti)
            lngConfirmedTotal = lngConfirmedTotal + 1
        ElseIf Left$(strResult, 6) = "REVIEW" Then
            Call IncrementDictionaryCount(dictReview, strGti)
            lngReviewTotal = lngReviewTotal + 1
        End If
    Next lngRow

    wksAnalysis.Range("A1").Value2 = "Schedule validation overview"
    wksAnalysis.Range("A2").Value2 = "Raw schedule mismatches"
    wksAnalysis.Range("B2").Value2 = lngRawMismatchTotal
    wksAnalysis.Range("A3").Value2 = "Confirmed simple-product mismatches"
    wksAnalysis.Range("B3").Value2 = lngConfirmedTotal
    wksAnalysis.Range("A4").Value2 = "Review / complex / stub cases"
    wksAnalysis.Range("B4").Value2 = lngReviewTotal

    wksAnalysis.Range("A6").Value2 = "GTI NAME"
    wksAnalysis.Range("B6").Value2 = "Raw schedule mismatch count"
    wksAnalysis.Range("C6").Value2 = "Confirmed simple mismatch count"
    wksAnalysis.Range("D6").Value2 = "Review count"

    arrKeys = BuildCombinedThreeKeyArray(dictRawMismatch, dictConfirmed, dictReview)
    If Not IsEmpty(arrKeys) Then
        Call SortKeysByThreeCounts(arrKeys, dictRawMismatch, dictConfirmed, dictReview)
        lngSummaryRows = UBound(arrKeys) - LBound(arrKeys) + 1
        ReDim arrSummary(1 To lngSummaryRows, 1 To 4)

        For lngIndex = LBound(arrKeys) To UBound(arrKeys)
            arrSummary(lngIndex - LBound(arrKeys) + 1, 1) = CStr(arrKeys(lngIndex))
            arrSummary(lngIndex - LBound(arrKeys) + 1, 2) = _
                GetDictionaryCount(dictRawMismatch, CStr(arrKeys(lngIndex)))
            arrSummary(lngIndex - LBound(arrKeys) + 1, 3) = _
                GetDictionaryCount(dictConfirmed, CStr(arrKeys(lngIndex)))
            arrSummary(lngIndex - LBound(arrKeys) + 1, 4) = _
                GetDictionaryCount(dictReview, CStr(arrKeys(lngIndex)))
        Next lngIndex

        wksAnalysis.Range("A7").Resize(lngSummaryRows, 4).Value2 = arrSummary

        If lngRawMismatchTotal > 0 Then
            For lngIndex = LBound(arrKeys) To UBound(arrKeys)
                If GetDictionaryCount(dictRawMismatch, CStr(arrKeys(lngIndex))) > 0 Then
                    lngChartRows = lngChartRows + 1
                    If lngChartRows = 15 Then Exit For
                Else
                    Exit For
                End If
            Next lngIndex

            If lngChartRows > 0 Then
                Set chtObject = wksAnalysis.ChartObjects.Add( _
                    Left:=wksAnalysis.Range("F2").Left, _
                    Top:=wksAnalysis.Range("F2").Top, Width:=650, Height:=380)

                With chtObject.Chart
                    .ChartType = xlBarClustered
                    .SetSourceData Source:=wksAnalysis.Range( _
                        wksAnalysis.Cells(6, 1), _
                        wksAnalysis.Cells(6 + lngChartRows, 2))
                    .HasTitle = True
                    .ChartTitle.Text = "Raw schedule mismatches by GTI Name - top 15"
                    .HasLegend = False
                    .SeriesCollection(1).ApplyDataLabels
                End With
            End If
        End If
    End If

    wksAnalysis.Range("A1:D1").Font.Bold = True
    wksAnalysis.Range("A6:D6").Font.Bold = True

ExitPoint:
    Set chtObject = Nothing
    Set dictRawMismatch = Nothing
    Set dictConfirmed = Nothing
    Set dictReview = Nothing

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

Private Sub IncrementDictionaryCount(ByVal dictTarget As Object, ByVal strKey As String)
    Const METHOD_NAME As String = "IncrementDictionaryCount"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If dictTarget.Exists(strKey) Then
        dictTarget(strKey) = CLng(dictTarget(strKey)) + 1
    Else
        dictTarget.Add strKey, 1&
    End If

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

Private Function GetDictionaryCount(ByVal dictTarget As Object, ByVal strKey As String) As Long
    Const METHOD_NAME As String = "GetDictionaryCount"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If dictTarget.Exists(strKey) Then
        GetDictionaryCount = CLng(dictTarget(strKey))
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

Private Function BuildCombinedThreeKeyArray( _
    ByVal dictFirst As Object, ByVal dictSecond As Object, _
    ByVal dictThird As Object) As Variant

    Const METHOD_NAME As String = "BuildCombinedThreeKeyArray"
    Dim arrKeys() As Variant
    Dim dictAll As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngIndex As Long
    Dim vKey As Variant

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dictAll = CreateObject("Scripting.Dictionary")
    dictAll.CompareMode = vbTextCompare

    For Each vKey In dictFirst.Keys
        If Not dictAll.Exists(CStr(vKey)) Then dictAll.Add CStr(vKey), True
    Next vKey
    For Each vKey In dictSecond.Keys
        If Not dictAll.Exists(CStr(vKey)) Then dictAll.Add CStr(vKey), True
    Next vKey
    For Each vKey In dictThird.Keys
        If Not dictAll.Exists(CStr(vKey)) Then dictAll.Add CStr(vKey), True
    Next vKey

    If dictAll.Count > 0 Then
        ReDim arrKeys(0 To dictAll.Count - 1)
        For Each vKey In dictAll.Keys
            arrKeys(lngIndex) = CStr(vKey)
            lngIndex = lngIndex + 1
        Next vKey
        BuildCombinedThreeKeyArray = arrKeys
    End If

ExitPoint:
    Set dictAll = Nothing
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

Private Sub SortKeysByThreeCounts( _
    ByRef arrKeys As Variant, ByVal dictRawMismatch As Object, _
    ByVal dictConfirmed As Object, ByVal dictReview As Object)

    Const METHOD_NAME As String = "SortKeysByThreeCounts"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngI As Long
    Dim lngJ As Long
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim strTemp As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngI = LBound(arrKeys) To UBound(arrKeys) - 1
        For lngJ = lngI + 1 To UBound(arrKeys)
            lngLeft = GetDictionaryCount(dictRawMismatch, CStr(arrKeys(lngI))) * 1000000 + _
                GetDictionaryCount(dictConfirmed, CStr(arrKeys(lngI))) * 1000 + _
                GetDictionaryCount(dictReview, CStr(arrKeys(lngI)))
            lngRight = GetDictionaryCount(dictRawMismatch, CStr(arrKeys(lngJ))) * 1000000 + _
                GetDictionaryCount(dictConfirmed, CStr(arrKeys(lngJ))) * 1000 + _
                GetDictionaryCount(dictReview, CStr(arrKeys(lngJ)))

            If lngRight > lngLeft Then
                strTemp = CStr(arrKeys(lngI))
                arrKeys(lngI) = arrKeys(lngJ)
                arrKeys(lngJ) = strTemp
            End If
        Next lngJ
    Next lngI

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
    Const AUTOFIT_SAMPLE_ROWS As Long = 300
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngColumn As Long
    Dim lngLastColumn As Long
    Dim lngLastRow As Long
    Dim lngSampleLastRow As Long
    Dim rngSample As Excel.Range

    If Not DEV_MODE Then On Error GoTo ErrHandler

    lngLastColumn = wksTarget.Cells(1, wksTarget.Columns.Count).End(xlToLeft).Column
    lngLastRow = wksTarget.Cells(wksTarget.Rows.Count, 1).End(xlUp).Row
    lngSampleLastRow = lngLastRow
    If lngSampleLastRow > AUTOFIT_SAMPLE_ROWS Then lngSampleLastRow = AUTOFIT_SAMPLE_ROWS
    If lngSampleLastRow < 1 Then lngSampleLastRow = 1

    With wksTarget.Rows(1)
        .Font.Bold = True
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = wksTarget.StandardHeight * HEADER_HEIGHT_FACTOR
    End With

    ' AutoFit over a bounded sample only. AutoFit on a very large UsedRange can
    ' dominate runtime without materially improving widths because all widths are capped.
    Set rngSample = wksTarget.Range( _
        wksTarget.Cells(1, 1), wksTarget.Cells(lngSampleLastRow, lngLastColumn))
    rngSample.Columns.AutoFit

    For lngColumn = 1 To lngLastColumn
        If wksTarget.Columns(lngColumn).ColumnWidth > MAX_COLUMN_WIDTH Then
            wksTarget.Columns(lngColumn).ColumnWidth = MAX_COLUMN_WIDTH
        End If
    Next lngColumn

ExitPoint:
    Set rngSample = Nothing
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

