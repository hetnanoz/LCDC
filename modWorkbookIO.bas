Option Explicit

Private Const CLASS_NAME As String = "modWorkbookIO"
Private Const ERR_WORKBOOK_IO As Long = vbObjectError + 6200

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strRangeName - required named cell in this macro workbook
' Returns:       String - trimmed value stored in the named cell
' Description:   Reads a mandatory path from a named cell on any worksheet.
'                Workbook-level and worksheet-level names are supported.
'-------------------------------------------------------------------------------
Public Function GetNamedPath(ByVal strRangeName As String) As String
    Const METHOD_NAME As String = "GetNamedPath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngMatches As Long
    Dim nmeCandidate As Excel.Name
    Dim nmeFound As Excel.Name
    Dim rngNamed As Excel.Range
    Dim strLocalName As String
    Dim strValue As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For Each nmeCandidate In ThisWorkbook.Names
        strLocalName = GetUnqualifiedName(CStr(nmeCandidate.Name))

        If StrComp(strLocalName, strRangeName, vbTextCompare) = 0 Then
            lngMatches = lngMatches + 1
            Set nmeFound = nmeCandidate
        End If
    Next nmeCandidate

    If lngMatches = 0 Then
        Err.Raise ERR_WORKBOOK_IO, METHOD_NAME, _
                  "Named cell '" & strRangeName & "' was not found in the macro workbook."
    ElseIf lngMatches > 1 Then
        Err.Raise ERR_WORKBOOK_IO, METHOD_NAME, _
                  "Named cell '" & strRangeName & "' exists more than once. " & _
                  "Keep only one workbook/sheet-level definition with this name."
    End If

    Set rngNamed = nmeFound.RefersToRange

    If rngNamed.Cells.CountLarge <> 1 Then
        Err.Raise ERR_WORKBOOK_IO, METHOD_NAME, _
                  "Named item '" & strRangeName & "' must refer to exactly one cell."
    End If

    strValue = Trim$(CStr(rngNamed.Value2))

    If Len(strValue) = 0 Then
        Err.Raise ERR_WORKBOOK_IO, METHOD_NAME, _
                  "Named cell '" & strRangeName & "' is empty."
    End If

    GetNamedPath = strValue

ExitPoint:
    Set rngNamed = Nothing
    Set nmeFound = Nothing
    Set nmeCandidate = Nothing

    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strRangeName", strRangeName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strDefinedName - workbook/sheet-level Excel defined name
' Returns:       String - name without workbook/worksheet qualifier
' Description:   Normalizes Name.Name so sheet-level names can be found safely.
'-------------------------------------------------------------------------------
Private Function GetUnqualifiedName(ByVal strDefinedName As String) As String
    Const METHOD_NAME As String = "GetUnqualifiedName"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngBang As Long
    Dim strResult As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strResult = strDefinedName
    lngBang = InStrRev(strResult, "!")

    If lngBang > 0 Then
        strResult = Mid$(strResult, lngBang + 1)
    End If

    If Len(strResult) >= 2 Then
        If Left$(strResult, 1) = "'" And Right$(strResult, 1) = "'" Then
            strResult = Mid$(strResult, 2, Len(strResult) - 2)
        End If
    End If

    GetUnqualifiedName = strResult

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
        "strDefinedName", strDefinedName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strFolder - folder path; strDescription - business description
' Returns:       None
' Description:   Stops processing when a required folder does not exist.
'-------------------------------------------------------------------------------
Public Sub ValidateFolderExists( _
    ByVal strFolder As String, ByVal strDescription As String)

    Const METHOD_NAME As String = "ValidateFolderExists"
    Dim errDescription As String
    Dim errNumber As Long
    Dim fso As Object

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(strFolder) Then
        Err.Raise ERR_WORKBOOK_IO, METHOD_NAME, _
                  strDescription & " folder does not exist: " & strFolder
    End If

ExitPoint:
    Set fso = Nothing

    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strDescription", strDescription)
    GoTo ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strTitle - title displayed in the file picker
' Returns:       String - selected .xlsx path, or empty string when cancelled
' Description:   Lets the user select one Excel .xlsx input file.
'-------------------------------------------------------------------------------
Public Function PickXlsxFile(ByVal strTitle As String) As String
    Const METHOD_NAME As String = "PickXlsxFile"
    Dim dlg As Object
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set dlg = Application.FileDialog(3)

    With dlg
        .Title = strTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Workbook", "*.xlsx"

        If .Show = -1 Then
            PickXlsxFile = CStr(.SelectedItems(1))
        End If
    End With

ExitPoint:
    Set dlg = Nothing

    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strTitle", strTitle)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strFullPath - exact workbook path; blnOpenedByMacro - output flag
' Returns:       Excel.Workbook - existing open workbook or read-only workbook
' Description:   Reuses an already-open workbook by FullName, otherwise opens it.
'-------------------------------------------------------------------------------
Public Function GetOrOpenWorkbookReadOnly( _
    ByVal strFullPath As String, ByRef blnOpenedByMacro As Boolean) As Excel.Workbook

    Const METHOD_NAME As String = "GetOrOpenWorkbookReadOnly"
    Dim errDescription As String
    Dim errNumber As Long
    Dim wkbSource As Excel.Workbook

    If Not DEV_MODE Then On Error GoTo ErrHandler

    blnOpenedByMacro = False
    Set wkbSource = GetOpenWorkbookByFullPath(strFullPath)

    If wkbSource Is Nothing Then
        Set wkbSource = Application.Workbooks.Open( _
            Filename:=strFullPath, UpdateLinks:=0, ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False)
        blnOpenedByMacro = True
    End If

    Set GetOrOpenWorkbookReadOnly = wkbSource

ExitPoint:
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
' Parameters:    wkbSource - workbook; strSheetName - exact required sheet name
' Returns:       Excel.Worksheet
' Description:   Returns the requested worksheet or raises a controlled error.
'-------------------------------------------------------------------------------
Public Function GetWorksheetByName( _
    ByVal wkbSource As Excel.Workbook, ByVal strSheetName As String) As Excel.Worksheet

    Const METHOD_NAME As String = "GetWorksheetByName"
    Dim errDescription As String
    Dim errNumber As Long
    Dim wksCandidate As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For Each wksCandidate In wkbSource.Worksheets
        If StrComp(wksCandidate.Name, strSheetName, vbTextCompare) = 0 Then
            Set GetWorksheetByName = wksCandidate
            Exit For
        End If
    Next wksCandidate

    If GetWorksheetByName Is Nothing Then
        Err.Raise ERR_WORKBOOK_IO, METHOD_NAME, _
                  "Required worksheet '" & strSheetName & "' was not found."
    End If

ExitPoint:
    Set wksCandidate = Nothing

    If errNumber <> 0 Then
        Call VBA.Err.Raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    End If
    Exit Function

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "strSheetName", strSheetName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    wkbSource - workbook
' Returns:       Excel.Worksheet - first worksheet in the workbook
' Description:   Returns worksheet 1 for input files whose sheet name can vary.
'-------------------------------------------------------------------------------
Public Function GetFirstWorksheet(ByVal wkbSource As Excel.Workbook) As Excel.Worksheet
    Const METHOD_NAME As String = "GetFirstWorksheet"
    Dim errDescription As String
    Dim errNumber As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If wkbSource.Worksheets.Count = 0 Then
        Err.Raise ERR_WORKBOOK_IO, METHOD_NAME, "The workbook contains no worksheets."
    End If

    Set GetFirstWorksheet = wkbSource.Worksheets(1)

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
' Parameters:    wkbWorkbook - workbook; blnSaveChanges - Close SaveChanges flag
' Returns:       None
' Description:   Isolated cleanup helper. Any close failure is logged and swallowed.
'-------------------------------------------------------------------------------
Public Sub CloseWorkbookSafe( _
    ByVal wkbWorkbook As Excel.Workbook, ByVal blnSaveChanges As Boolean)

    Const METHOD_NAME As String = "CloseWorkbookSafe"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strWorkbookName As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    If wkbWorkbook Is Nothing Then GoTo ExitPoint

    strWorkbookName = wkbWorkbook.Name
    wkbWorkbook.Close SaveChanges:=blnSaveChanges

ExitPoint:
    Exit Sub

ErrHandler:
    errNumber = VBA.Err.Number
    errDescription = VBA.Err.Description
    Call ErrorManager.addError( _
        CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
        "workbook", strWorkbookName)
    Resume ExitPoint
End Sub

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strFullPath - exact workbook path
' Returns:       Excel.Workbook - already-open workbook, or Nothing
' Description:   Matches by FullName to avoid same-name workbook collisions.
'-------------------------------------------------------------------------------
Private Function GetOpenWorkbookByFullPath( _
    ByVal strFullPath As String) As Excel.Workbook

    Const METHOD_NAME As String = "GetOpenWorkbookByFullPath"
    Dim errDescription As String
    Dim errNumber As Long
    Dim wkbCandidate As Excel.Workbook

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For Each wkbCandidate In Application.Workbooks
        If StrComp(wkbCandidate.FullName, strFullPath, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByFullPath = wkbCandidate
            Exit For
        End If
    Next wkbCandidate

ExitPoint:
    Set wkbCandidate = Nothing

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
