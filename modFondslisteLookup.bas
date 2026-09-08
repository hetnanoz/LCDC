Option Explicit

Private Const CLASS_NAME As String = "modFondslisteLookup"

Private Const FOND_FILE_PATTERN As String = "Fondsliste*.xls*"
Private Const FOND_SHEET_NAME As String = "Liste"
Private Const FOND_STATUS_TEXT As String = "Open"
Private Const FOND_STATUS_COL As Long = 2
Private Const FOND_FUND_COL As Long = 4
Private Const FOND_MATCH_COL As Long = 14
Private Const FOND_LAST_READ_COL As Long = 14
Private Const FOND_FIRST_DATA_ROW As Long = 2

Private Const ERR_FONDSLISTE As Long = vbObjectError + 6300

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strFolder - folder containing Fondsliste files
' Returns:       String - full path of the newest matching Fondsliste
' Description:   Chooses the Fondsliste with the newest last 8-digit filename key.
'-------------------------------------------------------------------------------
Public Function FindLatestFondsliste(ByVal strFolder As String) As String
    Const METHOD_NAME As String = "FindLatestFondsliste"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strFile As String
    Dim strKey As String
    Dim strLatest As String
    Dim strLatestKey As String
    Dim strPath As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Call ValidateFolderExists(strFolder, "Fondsliste")

    strPath = Trim$(strFolder)
    If Right$(strPath, 1) <> Application.PathSeparator Then
        strPath = strPath & Application.PathSeparator
    End If

    strFile = Dir$(strPath & FOND_FILE_PATTERN)

    Do While Len(strFile) > 0
        strKey = FileDateKey(strFile)

        If Len(strKey) > 0 Then
            If Len(strLatest) = 0 _
               Or strKey > strLatestKey _
               Or (strKey = strLatestKey And _
                   StrComp(strFile, strLatest, vbBinaryCompare) > 0) Then
                strLatest = strFile
                strLatestKey = strKey
            End If
        End If

        strFile = Dir$()
    Loop

    If Len(strLatest) = 0 Then
        Err.Raise ERR_FONDSLISTE, METHOD_NAME, _
                  "No dated Fondsliste file matching '" & FOND_FILE_PATTERN & "' was found."
    End If

    FindLatestFondsliste = strPath & strLatest

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
        "strFolder", strFolder)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    strFullPath - exact Fondsliste path
' Returns:       Object - dictionary: normalized column N account -> column D Fund
' Description:   Reads Open rows only and reuses the workbook if already open.
'-------------------------------------------------------------------------------
Public Function BuildOpenFundLookup(ByVal strFullPath As String) As Object
    Const METHOD_NAME As String = "BuildOpenFundLookup"
    Dim arrSource As Variant
    Dim blnOpenedByMacro As Boolean
    Dim dictFund As Object
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngLastRow As Long
    Dim lngRow As Long
    Dim strKey As String
    Dim strStatus As String
    Dim wkbSource As Excel.Workbook
    Dim wksSource As Excel.Worksheet

    If Not DEV_MODE Then On Error GoTo ErrHandler

    Set wkbSource = GetOrOpenWorkbookReadOnly(strFullPath, blnOpenedByMacro)
    Set wksSource = GetWorksheetByName(wkbSource, FOND_SHEET_NAME)

    lngLastRow = wksSource.Cells(wksSource.Rows.Count, FOND_STATUS_COL).End(xlUp).Row
    If wksSource.Cells(wksSource.Rows.Count, FOND_MATCH_COL).End(xlUp).Row > lngLastRow Then
        lngLastRow = wksSource.Cells(wksSource.Rows.Count, FOND_MATCH_COL).End(xlUp).Row
    End If
    If wksSource.Cells(wksSource.Rows.Count, FOND_FUND_COL).End(xlUp).Row > lngLastRow Then
        lngLastRow = wksSource.Cells(wksSource.Rows.Count, FOND_FUND_COL).End(xlUp).Row
    End If

    If lngLastRow < FOND_FIRST_DATA_ROW Then
        Err.Raise ERR_FONDSLISTE, METHOD_NAME, "Fondsliste does not contain data rows."
    End If

    arrSource = wksSource.Range( _
        wksSource.Cells(1, 1), _
        wksSource.Cells(lngLastRow, FOND_LAST_READ_COL)).Value2

    Set dictFund = CreateObject("Scripting.Dictionary")
    dictFund.CompareMode = vbTextCompare

    For lngRow = FOND_FIRST_DATA_ROW To UBound(arrSource, 1)
        strStatus = Trim$(CStr(arrSource(lngRow, FOND_STATUS_COL)))

        If StrComp(strStatus, FOND_STATUS_TEXT, vbTextCompare) = 0 Then
            strKey = NormalizeFondslisteAccount(arrSource(lngRow, FOND_MATCH_COL))

            If Len(strKey) > 0 Then
                If Not dictFund.Exists(strKey) Then
                    dictFund.Add strKey, Trim$(CStr(arrSource(lngRow, FOND_FUND_COL)))
                End If
            End If
        End If
    Next lngRow

    If dictFund.Count = 0 Then
        Err.Raise ERR_FONDSLISTE, METHOD_NAME, _
                  "No Open Fondsliste rows produced an account lookup key."
    End If

    Set BuildOpenFundLookup = dictFund

ExitPoint:
    Set wksSource = Nothing

    If blnOpenedByMacro Then
        Call CloseWorkbookSafe(wkbSource, False)
    End If

    Set wkbSource = Nothing
    Set dictFund = Nothing

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
' Parameters:    strName - Fondsliste file name
' Returns:       String - last 8-digit group found in the file name
' Description:   Extracts the YYYYMMDD-style key used to select the newest file.
'-------------------------------------------------------------------------------
Private Function FileDateKey(ByVal strName As String) As String
    Const METHOD_NAME As String = "FileDateKey"
    Dim errDescription As String
    Dim errNumber As Long
    Dim lngIndex As Long

    If Not DEV_MODE Then On Error GoTo ErrHandler

    For lngIndex = 1 To Len(strName) - 7
        If Mid$(strName, lngIndex, 8) Like "########" Then
            FileDateKey = Mid$(strName, lngIndex, 8)
        End If
    Next lngIndex

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
        "strName", strName)
    GoTo ExitPoint
End Function

'-------------------------------------------------------------------------------
' Author:        Pawel Ligezka
' Creation date: 2026-09-08
' Parameters:    vValue - account value from Fondsliste column N
' Returns:       String - normalized account key
' Description:   Removes spaces, keeps the last 11 characters and upper-cases it.
'-------------------------------------------------------------------------------
Private Function NormalizeFondslisteAccount(ByVal vValue As Variant) As String
    Const METHOD_NAME As String = "NormalizeFondslisteAccount"
    Dim errDescription As String
    Dim errNumber As Long
    Dim strAccount As String

    If Not DEV_MODE Then On Error GoTo ErrHandler

    strAccount = Replace(Trim$(CStr(vValue)), " ", vbNullString)
    If Len(strAccount) >= 11 Then strAccount = Right$(strAccount, 11)

    NormalizeFondslisteAccount = UCase$(strAccount)

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
