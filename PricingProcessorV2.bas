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
Private Const TIMESTAMP_FORMAT As String = "yyyymmdd_hhnnss"
Private Const DISABLE_NOTES As Boolean = False
Private Const DEBUG_LOG As Boolean = False
Private Const OUTPUT_SHEET_NAME As String = "Pricing Configurations - MARGOL"

' ========= COLUMN INDEX CACHE =========
Private Const COL_ASIN As Long = 3              ' ASIN
Private Const COL_REPR As Long = 15             ' Reprice
Private Const COL_FLOOR As Long = 20            ' Floor price
Private Const COL_CEIL As Long = 22             ' Ceiling Price
Private Const COL_REPR_METH As Long = 23        ' Repricer Method
Private Const COL_REPR_STRAT As Long = 24       ' Repricer Strategy
Private Const COL_REPR_STR_VAL As Long = 25     ' Repricer Strategy Value
Private Const COL_SALE As Long = 40             ' On-Sale indicator
Private Const COL_SALE_METH As Long = 41        ' Sale Repricer Method
Private Const COL_SALE_STR As Long = 42         ' Sale Repricer Strategy
Private Const COL_SALE_STR_VAL As Long = 43     ' Sale Repricer Strategy Value
Private Const COL_SALE_START As Long = 44       ' Sale Start Date
Private Const COL_SALE_END As Long = 46         ' Sale End Date


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
    Dim ws As Worksheet
    Set srcWs = Nothing

    For Each ws In wbSrc.Worksheets
        If InStr(1, ws.Name, "Pricing Configurations", vbTextCompare) > 0 Then
            Set srcWs = ws
            Exit For
        End If
    Next ws
    
    If srcWs Is Nothing Then
        MsgBox "No worksheet with 'Pricing Configurations' found in the workbook.", vbExclamation
        wbSrc.Close False
        Exit Sub
    End If

    Dim lastRow As Long, lastCol As Long
    lastRow = LastUsedRow(srcWs)
    lastCol = LastUsedColumn(srcWs)

    If lastRow <= 1 Or lastCol = 0 Then
        MsgBox "The selected CSV does not contain any data rows.", vbInformation
        GoTo Cleanup
    End If

    Dim sourceData As Variant
    sourceData = srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(lastRow, lastCol)).value

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
        wsConfig.Cells(DATA_START_ROW, 1).Resize(rowCount, colCount).value = computedData
    End If

    Dim changedRows As Collection
    Set changedRows = BuildResultPreview(wsPreview, wsConfig, sourceData, computedData, rowCount, colCount)

    If Not changedRows Is Nothing Then
        If changedRows.Count > 0 Then
            CreatePreviewFile wsPreview
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
    For Each asinKey In asinMap.keys
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
    Dim minR As Double
    Dim minRRow As Long

    For Each idx In rows
        Dim rowIndex As Long
        rowIndex = CLng(idx)

        If IsYesValue(data(rowIndex, COL_SALE)) Then
            hasSaleYes = True
        End If

        Dim serialDate As Double
        If TryGetSerialDate(data(rowIndex, COL_SALE_END), serialDate) Then
            If Not hasEndDate Or serialDate > latestEndDate Then
                latestEndDate = serialDate
                latestEndRow = rowIndex
            End If
            hasEndDate = True
        End If

        Dim sValue As Double
        If TryGetNumeric(data(rowIndex, COL_CEIL), sValue) Then
            If Not hasS Or sValue > highestS Then
                highestS = sValue
                hasS = True
            End If
        End If

        Dim rValue As Double
        If TryGetNumeric(data(rowIndex, COL_FLOOR), rValue) Then
            If rValue > 0 Then
                If minRRow = 0 Or rValue < minR Then
                    minR = rValue
                    minRRow = rowIndex
                End If
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
        ApplyBaselinePath data, rows, hasS, highestS, minRRow
    End If
End Sub

Private Sub ApplySalePath(ByRef data As Variant, ByVal rows As Collection, ByVal donorRow As Long)
    Dim idx As Variant
    Dim tomorrowDate As Date
    tomorrowDate = DateAdd("d", 1, Date)

    For Each idx In rows
        Dim rowIndex As Long
        rowIndex = CLng(idx)
        data(rowIndex, COL_SALE) = "Yes"
        data(rowIndex, COL_SALE_END) = data(donorRow, COL_SALE_END)
        data(rowIndex, COL_SALE_START) = tomorrowDate
        CopyColumns data, donorRow, rowIndex, Array(COL_SALE_METH, COL_SALE_STR, COL_SALE_STR_VAL, COL_REPR, COL_FLOOR, COL_CEIL, COL_REPR_METH, COL_REPR_STRAT, COL_REPR_STR_VAL)
    Next idx
End Sub

Private Sub ApplyBaselinePath(ByRef data As Variant, ByVal rows As Collection, _
                              ByVal hasS As Boolean, ByVal highestS As Double, ByVal donorRow As Long)
    Dim idx As Variant

    If hasS Then
        For Each idx In rows
            data(CLng(idx), COL_CEIL) = highestS
        Next idx
    End If

    If donorRow > 0 Then
        For Each idx In rows
            Dim rowIndex As Long
            rowIndex = CLng(idx)
            CopyColumns data, donorRow, rowIndex, Array(COL_SALE_METH, COL_SALE_STR, COL_SALE_STR_VAL, COL_REPR, COL_FLOOR, COL_REPR_METH, COL_REPR_STRAT, COL_REPR_STR_VAL)
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
        wsConfig.rows(1).Copy Destination:=wsPreview.rows(1)
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

    wsConfig.rows(1).Copy Destination:=wsPreview.rows(1)
    Application.CutCopyMode = False
    wsConfig.rows(1).Copy
    wsPreview.rows(1).PasteSpecial xlPasteColumnWidths
    Application.CutCopyMode = False

    If rowCount > 0 Then
        wsConfig.rows(DATA_START_ROW).Copy
        wsPreview.rows(DATA_START_ROW).Resize(Application.Max(1, changes.Count)).PasteSpecial xlPasteFormats
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

    wsPreview.Cells(DATA_START_ROW, 1).Resize(changes.Count, colCount).value = outputData

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

Private Sub CreatePreviewFile(wsPreview As Worksheet)
    Dim usedRange As Range
    Set usedRange = wsPreview.usedRange
    If usedRange Is Nothing Then Exit Sub
    If usedRange.rows.Count <= 1 Then Exit Sub

    Dim exportPath As String
    exportPath = ThisWorkbook.Path
    If Len(exportPath) = 0 Then exportPath = CurDir
    If Right$(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    exportPath = exportPath & Format(Now, TIMESTAMP_FORMAT) & "_result_preview.xlsx"

    wsPreview.Copy
    Dim wbExport As Workbook
    Set wbExport = ActiveWorkbook
    
    Dim wsExport As Worksheet
    Set wsExport = wbExport.Sheets(1)
    wsExport.Name = OUTPUT_SHEET_NAME
    
    Application.DisplayAlerts = False
    wbExport.SaveAs Filename:=exportPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Application.DisplayAlerts = True

    wbExport.Activate
    MsgBox "Results preview exported to:" & vbCrLf & exportPath & vbCrLf & _
           "The file has been left open for review.", vbInformation
End Sub

' ========= HELPERS =========
Private Sub EnsureColumnsPresent(ByVal colCount As Long)
    Dim requiredCols As Variant
    requiredCols = Array(COL_ASIN, COL_REPR, COL_FLOOR, COL_CEIL, COL_REPR_METH, COL_REPR_STRAT, COL_REPR_STR_VAL, COL_SALE, COL_SALE_METH, COL_SALE_STR, COL_SALE_STR_VAL, COL_SALE_START, COL_SALE_END)

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
        .Title = "Select Pricing Configuration File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*"
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


