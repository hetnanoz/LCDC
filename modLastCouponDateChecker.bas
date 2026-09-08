Option Explicit

Private Const CLASS_NAME As String = "modLastCouponDateChecker"
Private Const ROOT_FONDSLISTE_NAME As String = "root_Fondsliste"
Private Const SAVE_PATH_NAME As String = "save_path"

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    None
' Returns:       None
' Description:   MASTER entry point. Loads both inputs, resolves Fund Code through
'                the latest Open-only Fondsliste and creates the staging output.
'-------------------------------------------------------------------------------
Public Sub RunLastCouponDateChecker()
    Const METHOD_NAME As String = "RunLastCouponDateChecker"
    Dim arrLHS As Variant
    Dim arrNeolink As Variant
    Dim blnCancelled As Boolean
    Dim blnOutputSaved As Boolean
    Dim blnStateCaptured As Boolean
    Dim blnPreviousScreenUpdating As Boolean
    Dim dictFund As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLHSRows As Long
    Dim lngNeolinkRows As Long
    Dim lngUnmatchedRows As Long
    Dim strFondslisteFile As String
    Dim strFondslisteFolder As String
    Dim strLHSFile As String
    Dim strNeolinkFile As String
    Dim strOutputPath As String
    Dim strSaveFolder As String
    Dim strSuccessMessage As String
    Dim vPreviousStatusBar As Variant
    Dim wkbOutput As Excel.Workbook

    If Not DEV_MODE Then On Error GoTo ErrHandler

    blnPreviousScreenUpdating = Application.ScreenUpdating
    vPreviousStatusBar = Application.StatusBar
    blnStateCaptured = True

    Application.ScreenUpdating = False
    Application.StatusBar = "Last Coupon Date Checker: reading configuration..."

    strFondslisteFolder = GetMappingPath(ROOT_FONDSLISTE_NAME)
    strSaveFolder = GetMappingPath(SAVE_PATH_NAME)

    Call ValidateFolderExists(strFondslisteFolder, "Fondsliste")
    Call ValidateFolderExists(strSaveFolder, "Output")

    strLHSFile = PickXlsxFile( _
        "Select LHS input - FAMOO-Full Inventory IFRS workbook")
    If Len(strLHSFile) = 0 Then
        blnCancelled = True
        GoTo ExitPoint
    End If

    strNeolinkFile = PickXlsxFile("Select Neolink input workbook")
    If Len(strNeolinkFile) = 0 Then
        blnCancelled = True
        GoTo ExitPoint
    End If

    Application.StatusBar = "Last Coupon Date Checker: loading latest Fondsliste..."
    strFondslisteFile = FindLatestFondsliste(strFondslisteFolder)
    Set dictFund = BuildOpenFundLookup(strFondslisteFile)

    Application.StatusBar = "Last Coupon Date Checker: reading LHS input..."
    arrLHS = LoadLHSInput(strLHSFile, lngLHSRows)

    Application.StatusBar = "Last Coupon Date Checker: reading Neolink input..."
    arrNeolink = LoadNeolinkInput( _
        strNeolinkFile, dictFund, lngNeolinkRows, lngUnmatchedRows)

    Application.StatusBar = "Last Coupon Date Checker: creating output..."
    Set wkbOutput = CreateOutputWorkbook(arrLHS, arrNeolink)
    strOutputPath = BuildOutputPath(strSaveFolder)
    Call SaveOutputWorkbook(wkbOutput, strOutputPath)
    blnOutputSaved = True

ExitPoint:
    If errNumber <> 0 And Not blnOutputSaved Then
        Call CloseWorkbookSafe(wkbOutput, False)
    End If

    Set wkbOutput = Nothing
    Set dictFund = Nothing

    If blnStateCaptured Then
        Application.StatusBar = vPreviousStatusBar
        Application.ScreenUpdating = blnPreviousScreenUpdating
    End If

    If errNumber <> 0 Then
        Call ErrorManager.display
    ElseIf Not blnCancelled And blnOutputSaved Then
        strSuccessMessage = _
            "Output created successfully:" & vbCrLf & strOutputPath & vbCrLf & vbCrLf & _
            "Input LHS rows: " & CStr(lngLHSRows) & vbCrLf & _
            "Input Neolink rows included: " & CStr(lngNeolinkRows) & vbCrLf & _
            "Neolink rows without Fondsliste match: " & CStr(lngUnmatchedRows)

        MsgBox strSuccessMessage, vbInformation, "Last Coupon Date Checker"
    End If

    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription)
    GoTo ExitPoint
End Sub
