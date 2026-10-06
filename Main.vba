Option Explicit

Sub SplitDataByScopeDaily()
    ' ==============================
    ' Configuration toggles
    ' ==============================
    Const DEDUPE_WITHIN_SCOPE As Boolean = False   ' False = global de-dup (A,B,C) before split; True = per-scope de-dup
    Const DELETE_SOURCE_FILE_IF_WE_OPENED_IT As Boolean = True   ' Only delete if this macro opened the source file

    ' ==============================
    ' Declarations
    ' ==============================
    Dim sourceWorkbook As Workbook, sourceSheet As Worksheet
    Dim targetWorkbook As Workbook, targetSheet As Worksheet
    Dim wb As Workbook
    Dim filePath As String, saveFolder As String
    Dim scopeColumn As Long, lastRow As Long, lastCol As Long
    Dim uniqueScopes As Object
    Dim cell As Range, lastCell As Range
    Dim scopeColFound As Boolean, wbWasOpen As Boolean
    Dim i As Long, rowCount As Long
    Dim key As Variant
    Dim scopeVal As String
    Dim displayName As String
    Dim fileName As String
    Dim safeName As String
    Dim badChars As Variant, ch As Variant

    ' Route runtime errors to the handler and disable Excel updates during processing.
    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' ==============================
    ' 1) Read source file path
    ' ==============================
    ' B2 in the macro workbook contains the path to the source file.
    filePath = CStr(ThisWorkbook.Sheets(1).Range("B2").Value)
    If Len(filePath) = 0 Or Dir(filePath) = "" Then
        MsgBox "File not found: " & filePath, vbExclamation
        GoTo CleanExit
    End If

    ' ==============================
    ' 2) Open or reuse the source workbook
    ' ==============================
    wbWasOpen = False
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
            wbWasOpen = True
            Exit For
        End If
    Next wb

    If wbWasOpen Then
        Set sourceWorkbook = wb
    Else
        Set sourceWorkbook = Workbooks.Open(filePath)
    End If

    ' Process the first sheet of the source workbook.
    Set sourceSheet = sourceWorkbook.Sheets(1)

    ' ==============================
    ' 3) Determine last row/col
    ' ==============================
    Set lastCell = sourceSheet.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

    If lastCell Is Nothing Then
        MsgBox "No data found in the source sheet.", vbExclamation
        GoTo CleanExit
    End If

    lastRow = lastCell.Row
    lastCol = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Or lastCol < 3 Then
        MsgBox "Insufficient data. Expecting headers in row 1 and at least columns A..C.", vbExclamation
        GoTo CleanExit
    End If

    ' ==============================
    ' 3a) Update G when L contains "Completed"
    ' ==============================
    ' Read columns L and G into arrays so they can be processed together.
    Dim vL As Variant, vG As Variant, r As Long

    If lastRow = 2 Then
        ' One data row: explicitly create two-dimensional arrays.
        ReDim vL(1 To 1, 1 To 1)
        ReDim vG(1 To 1, 1 To 1)

        vL(1, 1) = sourceSheet.Range("L2").Value2
        vG(1, 1) = sourceSheet.Range("G2").Value2
    Else
        vL = sourceSheet.Range("L2:L" & lastRow).Value2
        vG = sourceSheet.Range("G2:G" & lastRow).Value2
    End If

    ' Keep the original case-insensitive substring rule: Not Completed also matches.
    For r = 1 To UBound(vL, 1)
        If InStr(1, CStr(vL(r, 1)), "Completed", vbTextCompare) > 0 Then
            vG(r, 1) = "Completed"
        End If
    Next r

    'Setting the updated Value
    ' Write the entire status range back as values, as in the original macro.
    sourceSheet.Range("G2:G" & lastRow).Value2 = vG

    ' --- Explicit cleanup ---
    Erase vL
    Erase vG

    ' ==============================
    ' 4) Find "Scope" column
    ' ==============================
    scopeColFound = False

    For i = 1 To lastCol
        If StrComp(Trim$(CStr(sourceSheet.Cells(1, i).Value)), "Scope", vbTextCompare) = 0 Then
            scopeColumn = i
            scopeColFound = True
            Exit For
        End If
    Next i

    If Not scopeColFound Then
        MsgBox "Scope column not found in the header row.", vbExclamation
        GoTo CleanExit
    End If

    ' ============================
    ' 5) Global de-dup (A,B,C) BEFORE split (if configured)
    ' ============================
    ' The combined A/B/C key determines duplicates; the first occurrence is retained.
    If Not DEDUPE_WITHIN_SCOPE Then
        sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastRow, lastCol)) _
                   .RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes

        ' Refresh lastRow after de-dup
        Set lastCell = sourceSheet.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                             SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

        If lastCell Is Nothing Or lastCell.Row < 2 Then
            MsgBox "All rows removed as duplicates. Nothing to split.", vbInformation
            GoTo CleanExit
        End If

        lastRow = lastCell.Row
    End If

    ' ============================
    ' 6) Collect unique (non-empty) Scope values
    ' ============================
    ' Blank Scope values are excluded from the reports.
    Set uniqueScopes = CreateObject("Scripting.Dictionary")

    For Each cell In sourceSheet.Range(sourceSheet.Cells(2, scopeColumn), sourceSheet.Cells(lastRow, scopeColumn))
        scopeVal = Trim$(CStr(cell.Value))
        If Len(scopeVal) > 0 Then
            If Not uniqueScopes.Exists(scopeVal) Then uniqueScopes.Add scopeVal, True
        End If
    Next cell

    If uniqueScopes.Count = 0 Then
        MsgBox "No non-empty Scope values found.", vbInformation
        GoTo CleanExit
    End If

    ' ============================
    ' 7) Output folder (same as source)
    ' ============================
    saveFolder = Left$(filePath, InStrRev(filePath, "\"))
    If Len(saveFolder) = 0 Then saveFolder = CurDir$ & "\"

    ' ============================
    ' 8) Split per Scope
    ' ============================
    For Each key In uniqueScopes.Keys
        Set targetWorkbook = Workbooks.Add(xlWBATWorksheet)
        Set targetSheet = targetWorkbook.Sheets(1)

        ' Copy headers
        sourceSheet.Rows(1).Resize(1, lastCol).Copy Destination:=targetSheet.Rows(1)
        rowCount = 1

        ' Copy rows matching this scope
        For i = 2 To lastRow
            If Trim$(CStr(sourceSheet.Cells(i, scopeColumn).Value)) = CStr(key) Then
                sourceSheet.Rows(i).Resize(1, lastCol).Copy Destination:=targetSheet.Rows(rowCount + 1)
                rowCount = rowCount + 1
            End If
        Next i

        ' If no rows copied, skip save
        If rowCount = 1 Then
            targetWorkbook.Close SaveChanges:=False
            GoTo NextScope
        End If

        ' Optional: de-dup WITHIN this scope output (A,B,C)
        If DEDUPE_WITHIN_SCOPE Then
            Dim tgtLastRow As Long, tgtLastCol As Long
            tgtLastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
            tgtLastCol = targetSheet.Cells(1, targetSheet.Columns.Count).End(xlToLeft).Column

            targetSheet.Range(targetSheet.Cells(1, 1), targetSheet.Cells(tgtLastRow, tgtLastCol)) _
                       .RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
        End If

        ' Map scope code to display name (inline)
        displayName = Trim$(CStr(key))

        Select Case UCase$(displayName)
            Case "SA":   displayName = "San Angelo"
            Case "SL":   displayName = "San Lorenzo"
            Case "JUZ":  displayName = "ETHICON Juarez"
            ' --- Add more mappings here as needed ---
            Case Else
                ' leave as-is
        End Select

        ' Build a safe filename (inline sanitization)
        safeName = displayName & "_Daily_Outcome_Summary_Report_" & Format(Date, "DDMMYYYY") & ".xlsx"
        badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

        For Each ch In badChars
            safeName = Replace$(safeName, ch, " ")
        Next ch

        safeName = Trim$(safeName)

        ' Save and close this scope workbook
        ' The filename uses the run date. No date filtering is applied to source rows.
        targetWorkbook.SaveAs fileName:=saveFolder & safeName, FileFormat:=xlOpenXMLWorkbook
        targetWorkbook.Close SaveChanges:=False

NextScope:
    Next key

    ' ============================
    ' 9) Close & optionally delete source (only if we opened it)
    ' ============================
    ' Original behavior: delete the source only when this macro opened it and the flag is True.
    If Not wbWasOpen Then
        sourceWorkbook.Close SaveChanges:=False

        If DELETE_SOURCE_FILE_IF_WE_OPENED_IT Then
            If Len(Dir(filePath)) > 0 Then Kill filePath
        End If
    Else
        ' Leave user's workbook open as-is
    End If

' Restore Excel settings before leaving the macro.
CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    ' This footer was not visible in the screenshots; a minimal handler is supplied.
    ' Show the error, then restore Excel settings through the existing CleanExit block.
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanExit
End Sub