Option Explicit

Private Const CLASS_NAME As String = "modOutput"
Private Const OUTPUT_BASE_NAME As String = "Last_Coupon_Date_Checker_"
Private Const ERR_OUTPUT As Long = vbObjectError + 6600

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    arrLHS - LHS output array; arrNeolink - Neolink output array
' Returns:       Excel.Workbook - newly created output workbook
' Description:   Creates Input LHS and Input Neolink sheets and writes arrays.
'-------------------------------------------------------------------------------
Public Function CreateOutputWorkbook( _
    ByVal arrLHS As Variant, ByVal arrNeolink As Variant) As Excel.Workbook

    Const METHOD_NAME As String = "CreateOutputWorkbook"
    Dim blnCompleted As Boolean
    Dim errDescription As String
    Dim errNumber As Long
    Dim wkbOutput As Excel.Workbook
    Dim wksLHS As Excel.Worksheet
    Dim wksNeolink As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set wkbOutput = Application.Workbooks.Add(xlWBATWorksheet)
    Set wksLHS = wkbOutput.Worksheets(1)
    wksLHS.Name = "Input LHS"

    Set wksNeolink = wkbOutput.Worksheets.Add(After:=wksLHS)
    wksNeolink.Name = "Input Neolink"

    wksLHS.Columns("A:B").NumberFormat = "@"
    wksLHS.Columns("C:D").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("G:G").NumberFormat = "dd/mm/yyyy"
    wksLHS.Columns("J:J").NumberFormat = "dd/mm/yyyy"

    wksNeolink.Columns("A:B").NumberFormat = "@"
    wksNeolink.Columns("F:G").NumberFormat = "dd/mm/yyyy"

    Call WriteArrayToWorksheet(wksLHS, arrLHS)
    Call WriteArrayToWorksheet(wksNeolink, arrNeolink)

    wksLHS.Rows(1).Font.Bold = True
    wksNeolink.Rows(1).Font.Bold = True
    wksLHS.UsedRange.Columns.AutoFit
    wksNeolink.UsedRange.Columns.AutoFit

    Set CreateOutputWorkbook = wkbOutput
    blnCompleted = True

ExitPoint:
    Set wksLHS = Nothing
    Set wksNeolink = Nothing

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
' Description:   Saves the macro-free output workbook as .xlsx.
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
