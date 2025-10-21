Attribute VB_Name = "PricingProcessorV2"
'====================================================================
' VBA Pricing Tool (Simplified Module)
' Provides clear/upload buttons with in-memory processing of pricing
' data and change tracking for Result Preview output.
'====================================================================
Option Explicit

' ========= USER CONFIG =========
Private Const SHEET_CONFIG As String = "Pricing Configurations"
Private Const SHEET_PREVIEW As String = "Result Preview"
Private Const DATA_START_ROW As Long = 2
Private Const CSV_TIMESTAMP_FORMAT As String = "yyyymmdd_hhnnss"
Private Const DISABLE_NOTES As Boolean = False
Private Const DEBUG_LOG As Boolean = False

' ========= COLUMN INDEX CACHE =========
Private Const COL_ASIN As Long = 3      ' Column C
Private Const COL_O As Long = 15        ' Column O
Private Const COL_R As Long = 18        ' Column R
Private Const COL_S As Long = 19        ' Column S
Private Const COL_T As Long = 20        ' Column T
Private Const COL_U As Long = 21        ' Column U
Private Const COL_V As Long = 22        ' Column V
Private Const COL_AH As Long = 34       ' Column AH
Private Const COL_AI As Long = 35       ' Column AI
Private Const COL_AJ As Long = 36       ' Column AJ
Private Const COL_AK As Long = 37       ' Column AK
Private Const COL_AL As Long = 38       ' Column AL
Private Const COL_AM As Long = 39       ' Column AM
Private Const COL_AO As Long = 41       ' Column AO
Private Const COL_AY As Long = 51       ' Column AY

' ========= BUTTON ENTRY POINTS =========
Public Sub Btn_ClearPricingDataV2()
    On Error GoTo EH
    OptimizeStart

    Dim wb As Workbook
    Set wb = ThisWorkbook

    ClearSheetData wb.Worksheets(SHEET_CONFIG)
    ClearSheetData wb.Worksheets(SHEET_PREVIEW)

    GoTo Finally
EH:
    MsgBox "Clear Pricing Data failed: " & Err.Description, vbExclamation
Finally:
    OptimizeEnd
End Sub

Public Sub Btn_UploadAndProcessV2()
    On Error GoTo EH
    OptimizeStart

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim wsConfig As Worksheet
    Dim wsPreview As Worksheet
    Set wsConfig = wb.Worksheets(SHEET_CONFIG)
    Set wsPreview = wb.Worksheets(SHEET_PREVIEW)

    ClearSheetData wsConfig
    ClearSheetData wsPreview

    Dim csvPath As String
    csvPath = PickCsvPath()
    If Len(csvPath) = 0 Then GoTo Finally

    Dim wbSrc As Workbook
    Set wbSrc = Workbooks.Open(Filename:=csvPath, ReadOnly:=True)

    Dim srcWs As Worksheet
    Set srcWs = wbSrc.Worksheets(1)

    Dim lastRow As Long, lastCol As Long
    lastRow = LastUsedRow(srcWs)
    lastCol = LastUsedColumn(srcWs)

    If lastRow <= 1 Or lastCol = 0 Then
        MsgBox "The selected CSV does not contain any data rows.", vbInformation
        GoTo Cleanup
    End If

    Dim sourceData As Variant
    sourceData = srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(lastRow, lastCol)).Value

Cleanup:
    If Not wbSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
    End If

    If Not IsArray(sourceData) Then GoTo Finally

    Dim rowCount As Long
    Dim colCount As Long
    rowCount = UBound(sourceData, 1)
    colCount = UBound(sourceData, 2)

    EnsureColumnsPresent colCount

    Dim computedData As Variant
    computedData = sourceData

    ApplyPricingRules computedData, rowCount, colCount

    If rowCount > 0 And colCount > 0 Then
        wsConfig.Cells(DATA_START_ROW, 1).Resize(rowCount, colCount).Value = computedData
    End If

    Dim changedRows As Collection
    Set changedRows = BuildResultPreview(wsPreview, wsConfig, sourceData, computedData, rowCount, colCount)

    If Not changedRows Is Nothing Then
        If changedRows.Count > 0 Then
            CreatePreviewCsv wsPreview
        Else
            MsgBox "No changes were required for the uploaded data.", vbInformation
        End If
    End If

    GoTo Finally
EH:
    MsgBox "Upload and process failed: " & Err.Description, vbExclamation
Finally:
    OptimizeEnd
End Sub

' ========= CORE LOGIC =========
Private Sub ApplyPricingRules(ByRef data As Variant, ByVal rowCount As Long, ByVal colCount As Long)
    If rowCount = 0 Then Exit Sub

    Dim asinMap As Object
    Set asinMap = CreateObject("Scripting.Dictionary")
    asinMap.CompareMode = vbTextCompare

    Dim r As Long
    For r = 1 To rowCount
        Dim key As String
        key = Trim$(CStr(data(r, COL_ASIN)))
        If Len(key) = 0 Then
            key = "#BLANK#" & Format$(r, "000000")
        End If
        If Not asinMap.Exists(key) Then
            Dim coll As Collection
            Set coll = New Collection
            coll.Add r
            asinMap.Add key, coll
        Else
            asinMap(key).Add r
        End If
    Next r

    Dim asinKey As Variant
    For Each asinKey In asinMap.Keys
        ApplyRulesToGroup data, asinMap(asinKey)
    Next asinKey
End Sub

Private Sub ApplyRulesToGroup(ByRef data As Variant, ByVal rows As Collection)
    If rows.Count = 0 Then Exit Sub

    Dim idx As Variant
    Dim hasSaleYes As Boolean
    Dim hasEndDate As Boolean
    Dim latestEndDate As Double
    Dim latestEndRow As Long
    Dim highestS As Double
    Dim hasS As Boolean
    Dim minAH As Double
    Dim minAHRow As Long

    For Each idx In rows
        Dim rowIndex As Long
        rowIndex = CLng(idx)

        If IsYesValue(data(rowIndex, COL_AI)) Then
            hasSaleYes = True
        End If

        Dim serialDate As Double
        If TryGetSerialDate(data(rowIndex, COL_AO), serialDate) Then
            If Not hasEndDate Or serialDate > latestEndDate Then
                latestEndDate = serialDate
                latestEndRow = rowIndex
            End If
            hasEndDate = True
        End If

        Dim sValue As Double
        If TryGetNumeric(data(rowIndex, COL_S), sValue) Then
            If Not hasS Or sValue > highestS Then
                highestS = sValue
                hasS = True
            End If
        End If

        Dim ahValue As Double
        If TryGetNumeric(data(rowIndex, COL_AH), ahValue) Then
            If minAHRow = 0 Or ahValue < minAH Then
                minAH = ahValue
                minAHRow = rowIndex
            End If
        End If
    Next idx

    Dim saleActive As Boolean
    saleActive = False
    If hasSaleYes And latestEndRow > 0 Then
        If latestEndDate > Date Then
            saleActive = True
        End If
    End If

    If saleActive Then
        ApplySalePath data, rows, latestEndRow
    Else
        ApplyBaselinePath data, rows, hasS, highestS, minAHRow
    End If
End Sub

Private Sub ApplySalePath(ByRef data As Variant, ByVal rows As Collection, ByVal donorRow As Long)
    Dim idx As Variant
    Dim tomorrowDate As Date
    tomorrowDate = DateAdd("d", 1, Date)

    For Each idx In rows
        Dim rowIndex As Long
        rowIndex = CLng(idx)
        data(rowIndex, COL_AI) = "Yes"
        data(rowIndex, COL_AO) = data(donorRow, COL_AO)
        data(rowIndex, COL_AM) = tomorrowDate
        CopyColumns data, donorRow, rowIndex, Array(COL_AJ, COL_AK, COL_AL, COL_O, COL_R, COL_S, COL_T, COL_U, COL_V)
    Next idx
End Sub

Private Sub ApplyBaselinePath(ByRef data As Variant, ByVal rows As Collection, _
                              ByVal hasS As Boolean, ByVal highestS As Double, ByVal donorRow As Long)
    Dim idx As Variant

    If hasS Then
        For Each idx In rows
            data(CLng(idx), COL_S) = highestS
        Next idx
    End If

    If donorRow > 0 Then
        For Each idx In rows
            Dim rowIndex As Long
            rowIndex = CLng(idx)
            CopyColumns data, donorRow, rowIndex, Array(COL_AJ, COL_AK, COL_AL, COL_O, COL_R, COL_T, COL_U, COL_V)
        Next idx
    End If
End Sub

Private Sub CopyColumns(ByRef data As Variant, ByVal fromRow As Long, ByVal toRow As Long, cols As Variant)
    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        data(toRow, CLng(cols(i))) = data(fromRow, CLng(cols(i)))
    Next i
End Sub

' ========= RESULT PREVIEW =========
Private Function BuildResultPreview(wsPreview As Worksheet, wsConfig As Worksheet, _
                                    originalData As Variant, computedData As Variant, _
                                    ByVal rowCount As Long, ByVal colCount As Long) As Collection
    Dim changes As New Collection
    Set BuildResultPreview = changes

    wsPreview.Cells.Clear

    If rowCount = 0 Or colCount = 0 Then
        wsConfig.Rows(1).Copy Destination:=wsPreview.Rows(1)
        Application.CutCopyMode = False
        Exit Function
    End If

    Dim rowHasChange() As Boolean
    ReDim rowHasChange(1 To rowCount) As Boolean

    Dim r As Long, c As Long
    For r = 1 To rowCount
        For c = 1 To colCount
            If Not ValuesEqual(originalData(r, c), computedData(r, c)) Then
                rowHasChange(r) = True
                changes.Add r
                Exit For
            End If
        Next c
    Next r

    wsConfig.Rows(1).Copy Destination:=wsPreview.Rows(1)
    Application.CutCopyMode = False
    wsConfig.Rows(1).Copy
    wsPreview.Rows(1).PasteSpecial xlPasteColumnWidths
    Application.CutCopyMode = False

    If rowCount > 0 Then
        wsConfig.Rows(DATA_START_ROW).Copy
        wsPreview.Rows(DATA_START_ROW).Resize(Application.Max(1, changes.Count)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

    If changes.Count = 0 Then Exit Function

    Dim outputData As Variant
    ReDim outputData(1 To changes.Count, 1 To colCount)

    Dim outIndex As Long
    outIndex = 0

    For r = 1 To rowCount
        If rowHasChange(r) Then
            outIndex = outIndex + 1
            For c = 1 To colCount
                outputData(outIndex, c) = computedData(r, c)
            Next c
        End If
    Next r

    wsPreview.Cells(DATA_START_ROW, 1).Resize(changes.Count, colCount).Value = outputData

    outIndex = 0
    For r = 1 To rowCount
        If rowHasChange(r) Then
            outIndex = outIndex + 1
            For c = 1 To colCount
                If Not ValuesEqual(originalData(r, c), computedData(r, c)) Then
                    Dim tgt As Range
                    Set tgt = wsPreview.Cells(DATA_START_ROW + outIndex - 1, c)
                    Dim noteText As String
                    noteText = "Previous value: " & FormatForNote(originalData(r, c))
                    AddCellNote tgt, noteText
                End If
            Next c
        End If
    Next r
End Function

Private Sub CreatePreviewCsv(wsPreview As Worksheet)
    Dim usedRange As Range
    Set usedRange = wsPreview.UsedRange
    If usedRange Is Nothing Then Exit Sub
    If usedRange.Rows.Count <= 1 Then Exit Sub

    Dim exportPath As String
    exportPath = ThisWorkbook.Path
    If Len(exportPath) = 0 Then exportPath = CurDir
    If Right$(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    exportPath = exportPath & Format(Now, CSV_TIMESTAMP_FORMAT) & "_result_preview.csv"

    wsPreview.Copy
    Dim wbExport As Workbook
    Set wbExport = ActiveWorkbook

    Application.DisplayAlerts = False
    wbExport.SaveAs Filename:=exportPath, FileFormat:=xlCSV, CreateBackup:=False
    Application.DisplayAlerts = True

    wbExport.Activate
    MsgBox "Results preview exported to:" & vbCrLf & exportPath & vbCrLf & _
           "The CSV has been left open for review.", vbInformation
End Sub

' ========= HELPERS =========
Private Sub EnsureColumnsPresent(ByVal colCount As Long)
    Dim requiredCols As Variant
    requiredCols = Array(COL_ASIN, COL_O, COL_R, COL_S, COL_T, COL_U, COL_V, COL_AH, COL_AI, COL_AJ, COL_AK, COL_AL, COL_AM, COL_AO)

    Dim i As Long
    For i = LBound(requiredCols) To UBound(requiredCols)
        If colCount < CLng(requiredCols(i)) Then
            Err.Raise vbObjectError + 2000 + i, "EnsureColumnsPresent", _
                      "Source data is missing required column '" & ColumnLetter(CLng(requiredCols(i))) & "'."
        End If
    Next i
End Sub

Private Function PickCsvPath() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select Pricing Configuration CSV"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        If .Show = -1 Then
            PickCsvPath = .SelectedItems(1)
        Else
            PickCsvPath = vbNullString
        End If
    End With
End Function

Private Function TryGetSerialDate(ByVal value As Variant, ByRef serial As Double) As Boolean
    On Error GoTo Fail
    If IsDate(value) Then
        serial = CDbl(CDate(value))
        TryGetSerialDate = True
        Exit Function
    End If
Fail:
    TryGetSerialDate = False
End Function

Private Function TryGetNumeric(ByVal value As Variant, ByRef numberOut As Double) As Boolean
    On Error GoTo Fail
    If IsNumeric(value) Then
        numberOut = CDbl(value)
        TryGetNumeric = True
        Exit Function
    End If
Fail:
    TryGetNumeric = False
End Function

Private Function IsYesValue(ByVal value As Variant) As Boolean
    IsYesValue = (UCase$(Trim$(CStr(value))) = "YES")
End Function

Private Function ValuesEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsMissingOrEmpty(a) And IsMissingOrEmpty(b) Then
        ValuesEqual = True
        Exit Function
    End If

    If TryCompareAsDate(a, b) Then
        ValuesEqual = True
        Exit Function
    End If

    If TryCompareAsNumber(a, b) Then
        ValuesEqual = True
        Exit Function
    End If

    ValuesEqual = (CStr(a) = CStr(b))
End Function

Private Function TryCompareAsDate(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim da As Double, db As Double
    If TryGetSerialDate(a, da) And TryGetSerialDate(b, db) Then
        TryCompareAsDate = (Abs(da - db) < 0.0000001)
    Else
        TryCompareAsDate = False
    End If
End Function

Private Function TryCompareAsNumber(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim na As Double, nb As Double
    If TryGetNumeric(a, na) And TryGetNumeric(b, nb) Then
        TryCompareAsNumber = (Abs(na - nb) < 0.0000001)
    Else
        TryCompareAsNumber = False
    End If
End Function

Private Function IsMissingOrEmpty(ByVal value As Variant) As Boolean
    If IsEmpty(value) Then
        IsMissingOrEmpty = True
    ElseIf IsNull(value) Then
        IsMissingOrEmpty = True
    ElseIf Trim$(CStr(value)) = "" Then
        IsMissingOrEmpty = True
    Else
        IsMissingOrEmpty = False
    End If
End Function

Private Function FormatForNote(ByVal value As Variant) As String
    If IsMissingOrEmpty(value) Then
        FormatForNote = "(blank)"
    ElseIf TryCompareAsDate(value, value) Then
        FormatForNote = Format$(CDate(value), "yyyy-mm-dd")
    ElseIf IsNumeric(value) Then
        FormatForNote = CStr(value)
    Else
        FormatForNote = CStr(value)
    End If
End Function

Private Function ColumnLetter(ByVal colNumber As Long) As String
    Dim n As Long
    n = colNumber
    Dim letters As String
    Do While n > 0
        letters = Chr$(((n - 1) Mod 26) + 65) & letters
        n = (n - 1) \\ 26
    Loop
    ColumnLetter = letters
End Function

Private Function LastUsedRow(ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If lastCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = lastCell.Row
    End If
End Function

Private Function LastUsedColumn(ws As Worksheet) As Long
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If lastCell Is Nothing Then
        LastUsedColumn = 0
    Else
        LastUsedColumn = lastCell.Column
    End If
End Function

Private Sub ClearSheetData(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    If lastRow >= DATA_START_ROW And lastCol > 0 Then
        ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(lastRow, lastCol)).ClearContents
        On Error Resume Next
        ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(lastRow, lastCol)).ClearComments
        On Error GoTo 0
    End If
End Sub

' ========= NOTES =========
Private Sub AddCellNote(ByVal tgt As Range, ByVal msg As String)
    If DISABLE_NOTES Then Exit Sub
    If tgt Is Nothing Then Exit Sub

    On Error Resume Next
    If Not tgt.Comment Is Nothing Then tgt.Comment.Delete
    On Error GoTo 0

    Dim firstErr As Long
    On Error Resume Next
    tgt.AddComment msg
    firstErr = Err.Number
    On Error GoTo 0

    If firstErr <> 0 Then
        Dim secondErr As Long, secondDesc As String
        On Error Resume Next
        tgt.AddCommentThreaded msg
        secondErr = Err.Number
        secondDesc = Err.Description
        On Error GoTo 0
        If secondErr <> 0 Then
            Err.Raise secondErr, "AddCellNote", "Failed to add note: " & secondDesc
        End If
    End If
End Sub

' ========= OPTIMIZATION =========
Private Sub OptimizeStart()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub OptimizeEnd()
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub

