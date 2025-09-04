Attribute VB_Name = "Module1"
'====================================================================
' VBA Pricing Tool  Clear & Upload/Process (No extra worksheets)
' Using optimized ASIN struct computation for A:P
'====================================================================
Option Explicit

' ========= USER CONFIG =========
Private Const TOOL_SHEET_NAME As String = "Pricing Configurations"
Private Const PASTE_START_CELL As String = "Q1"
Private Const FILTER_COL_LETTER As String = "O"
Private Const EXPORT_SHEET_NAME As String = "Pricing Configurations"

Private Const DISABLE_NOTES As Boolean = False
Private Const USE_VBA_COMPUTE As Boolean = True

Private Function MAPPING_PAIRS() As Variant
    MAPPING_PAIRS = Array( _
        Array("A", "O"), _
        Array("B", "T"), _
        Array("C", "V"), _
        Array("D", "W"), _
        Array("E", "X"), _
        Array("F", "Y"), _
        Array("G", "AL"), _
        Array("H", "AM"), _
        Array("I", "AN"), _
        Array("J", "AO"), _
        Array("K", "AP"), _
        Array("L", "AQ"), _
        Array("M", "AR"), _
        Array("N", "AS") _
    )
End Function

' ========= BUTTON ENTRY POINTS =========
Public Sub Btn_ClearPricingData()
    On Error GoTo EH
    OptimizeStart
    ClearPricingData ThisWorkbook.Worksheets(TOOL_SHEET_NAME)
    GoTo Finally
EH:
    MsgBox "Clear failed: " & Err.Description, vbExclamation
Finally:
    OptimizeEnd
End Sub

Public Sub Btn_UploadAndProcess()
    On Error GoTo EH
    OptimizeStart
    Dim wsTool As Worksheet
    Set wsTool = ThisWorkbook.Worksheets(TOOL_SHEET_NAME)
    ClearPricingData wsTool

    Dim srcPath As String
    srcPath = PickSourceWorkbookPath()
    If Len(srcPath) = 0 Then GoTo Finally

    Dim wbSrc As Workbook
    Set wbSrc = Workbooks.Open(Filename:=srcPath, ReadOnly:=True)
    ImportAllPricingConfigurationSheets wbSrc, wsTool, wsTool.Range(PASTE_START_CELL)
    wbSrc.Close SaveChanges:=False

    Dim lastRow As Long, lastCol As Long
    UsedBounds wsTool, lastRow, lastCol

    If USE_VBA_COMPUTE Then
        ComputeDerivedColumns_AP wsTool, lastRow
    End If

    BuildFilteredExport wsTool, PASTE_START_CELL
    GoTo Finally
EH:
    MsgBox "Upload/Process failed: " & Err.Description, vbExclamation
Finally:
    OptimizeEnd
End Sub

' ========= CORE LOGIC =========
Private Sub ClearPricingData(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    UsedBounds ws, lastRow, lastCol
    If lastRow < 3 Then Exit Sub
    ws.Range(ws.Rows(3), ws.Rows(lastRow)).ClearContents
End Sub

Private Function PickSourceWorkbookPath() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select source Excel file"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xlsb;*.xls"
        If .Show = -1 Then
            PickSourceWorkbookPath = .SelectedItems(1)
        Else
            PickSourceWorkbookPath = vbNullString
        End If
    End With
End Function

Private Sub ImportAllPricingConfigurationSheets(wbSrc As Workbook, wsTool As Worksheet, pasteStart As Range)
    Dim nextPasteRow As Long
    nextPasteRow = pasteStart.Row
    Dim sh As Worksheet
    For Each sh In wbSrc.Worksheets
        If InStr(1, sh.Name, "Pricing Configurations", vbTextCompare) > 0 Then
            Dim rng As Range
            Set rng = SheetDataRange(sh)
            If Not rng Is Nothing Then
                wsTool.Cells(nextPasteRow, pasteStart.Column).Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value
                nextPasteRow = nextPasteRow + rng.Rows.Count
            End If
        End If
    Next sh
End Sub

Private Function SheetDataRange(ws As Worksheet) As Range
    On Error GoTo EH
    Dim ur As Range
    Set ur = ws.UsedRange
    If ur Is Nothing Then Exit Function
    Set SheetDataRange = ur
    Exit Function
EH:
    Set SheetDataRange = Nothing
    On Error GoTo 0
    Err.Raise Err.Number, "SheetDataRange", _
              "Failed to obtain UsedRange for '" & ws.Name & "': " & Err.Description
End Function

' ========= ASIN STRUCT COMPUTE =========
Private Sub ComputeDerivedColumns_AP(ws As Worksheet, lastRow As Long)
    If lastRow < 2 Then Exit Sub

    Dim n As Long: n = lastRow - 1

    Dim vS As Variant, vAE As Variant, vAJ As Variant, vAL As Variant
    Dim vAM As Variant, vAN As Variant, vAO As Variant, vBB As Variant
    Dim vBC As Variant, vBD As Variant, vBE As Variant, vBF As Variant
    Dim vBG As Variant, vBH As Variant, vBI As Variant

    vS = ws.Range("S2:S" & lastRow).Value2
    vAE = ws.Range("AE2:AE" & lastRow).Value2
    vAJ = ws.Range("AJ2:AJ" & lastRow).Value2
    vAL = ws.Range("AL2:AL" & lastRow).Value2
    vAM = ws.Range("AM2:AM" & lastRow).Value2
    vAN = ws.Range("AN2:AN" & lastRow).Value2
    vAO = ws.Range("AO2:AO" & lastRow).Value2
    vBB = ws.Range("BB2:BB" & lastRow).Value2
    vBC = ws.Range("BC2:BC" & lastRow).Value2
    vBD = ws.Range("BD2:BD" & lastRow).Value2
    vBE = ws.Range("BE2:BE" & lastRow).Value2
    vBF = ws.Range("BF2:BF" & lastRow).Value2
    vBG = ws.Range("BG2:BG" & lastRow).Value2
    vBH = ws.Range("BH2:BH" & lastRow).Value2
    vBI = ws.Range("BI2:BI" & lastRow).Value2

    Dim asinIdx As Object
    Dim cnt() As Long, firstRow() As Long, minAJ() As Double, minAJRow() As Long
    Dim maxAL() As Double, donorRow() As Long, minBEff() As Double, minBEffRow() As Long
    Dim dAE() As Boolean, dAJ() As Boolean, dAL() As Boolean, dAM() As Boolean
    Dim dAN() As Boolean, dAO() As Boolean, dBB() As Boolean, dBC() As Boolean
    Dim dBD() As Boolean, dBE() As Boolean, dBF() As Boolean, dBG() As Boolean
    Dim dBH() As Boolean, dBI() As Boolean

    BuildAsinAggregates n, vS, vAE, vAJ, vAL, vAM, vAN, vAO, vBB, vBC, vBD, vBE, vBF, vBG, vBH, vBI, _
                        asinIdx, cnt, firstRow, minAJ, minAJRow, maxAL, donorRow, minBEff, minBEffRow, _
                        dAE, dAJ, dAL, dAM, dAN, dAO, dBB, dBC, dBD, dBE, dBF, dBG, dBH, dBI

    Dim colP As Long: colP = ColLetterToNum("P")
    Dim outAP As Variant: ReDim outAP(1 To n, 1 To colP)
    Dim i As Long
    For i = 1 To n
        PopulateOutputRow i, outAP, asinIdx, cnt, firstRow, minAJ, minAJRow, maxAL, donorRow, minBEff, minBEffRow, _
                          dAE, dAJ, dAL, dAM, dAN, dAO, dBB, dBC, dBD, dBE, dBF, dBG, dBH, dBI, _
                          vS, vAE, vAJ, vAL, vAM, vAN, vAO, vBB, vBC, vBD, vBE, vBF, vBG, vBH, vBI
    Next i

    ws.Range("A2", ws.Cells(lastRow, colP)).Value = outAP
End Sub

Private Sub BuildAsinAggregates(ByVal n As Long, vS As Variant, vAE As Variant, vAJ As Variant, vAL As Variant, _
                                vAM As Variant, vAN As Variant, vAO As Variant, vBB As Variant, vBC As Variant, vBD As Variant, _
                                vBE As Variant, vBF As Variant, vBG As Variant, vBH As Variant, vBI As Variant, _
                                ByRef asinIdx As Object, ByRef cnt() As Long, ByRef firstRow() As Long, _
                                ByRef minAJ() As Double, ByRef minAJRow() As Long, ByRef maxAL() As Double, _
                                ByRef donorRow() As Long, ByRef minBEff() As Double, ByRef minBEffRow() As Long, _
                                ByRef dAE() As Boolean, ByRef dAJ() As Boolean, ByRef dAL() As Boolean, ByRef dAM() As Boolean, _
                                ByRef dAN() As Boolean, ByRef dAO() As Boolean, ByRef dBB() As Boolean, ByRef dBC() As Boolean, _
                                ByRef dBD() As Boolean, ByRef dBE() As Boolean, ByRef dBF() As Boolean, ByRef dBG() As Boolean, _
                                ByRef dBH() As Boolean, ByRef dBI() As Boolean)

    Set asinIdx = CreateObject("Scripting.Dictionary")
    asinIdx.CompareMode = vbTextCompare

    Dim cap As Long: cap = 0
    Dim k As Long: k = 0
    Dim i As Long

    Dim fAE() As Variant, fAJ() As Variant, fAL() As Variant, fAM() As Variant
    Dim fAN() As Variant, fAO() As Variant, fBB() As Variant, fBC() As Variant
    Dim fBD() As Variant, fBE() As Variant, fBF() As Variant, fBG() As Variant
    Dim fBH() As Variant, fBI() As Variant

    For i = 1 To n
        Dim s As String: s = CStr(vS(i, 1))
        If Not asinIdx.Exists(s) Then
            k = k + 1
            If k > cap Then
                cap = IIf(cap = 0, 256, cap * 2)
                ReDim Preserve cnt(1 To cap), firstRow(1 To cap), minAJ(1 To cap), minAJRow(1 To cap), maxAL(1 To cap), _
                                donorRow(1 To cap), minBEff(1 To cap), minBEffRow(1 To cap)
                ReDim Preserve fAE(1 To cap), dAE(1 To cap), fAJ(1 To cap), dAJ(1 To cap), _
                                fAL(1 To cap), dAL(1 To cap), fAM(1 To cap), dAM(1 To cap), _
                                fAN(1 To cap), dAN(1 To cap), fAO(1 To cap), dAO(1 To cap), _
                                fBB(1 To cap), dBB(1 To cap), fBC(1 To cap), dBC(1 To cap), _
                                fBD(1 To cap), dBD(1 To cap), fBE(1 To cap), dBE(1 To cap), _
                                fBF(1 To cap), dBF(1 To cap), fBG(1 To cap), dBG(1 To cap), _
                                fBH(1 To cap), dBH(1 To cap), fBI(1 To cap), dBI(1 To cap)
            End If
            asinIdx(s) = k
            cnt(k) = 0
            firstRow(k) = i
            minAJ(k) = 1E+308
            minAJRow(k) = 0
            maxAL(k) = -1E+308
            minBEff(k) = 1E+308: minBEffRow(k) = 0
            donorRow(k) = 0
            fAE(k) = Empty: fAJ(k) = Empty: fAL(k) = Empty
            fAM(k) = Empty: fAN(k) = Empty: fAO(k) = Empty
            fBB(k) = Empty: fBC(k) = Empty: fBD(k) = Empty
            fBE(k) = Empty: fBF(k) = Empty: fBG(k) = Empty
            fBH(k) = Empty: fBI(k) = Empty
        End If

        Dim idx As Long: idx = CLng(asinIdx(s))
        cnt(idx) = cnt(idx) + 1

        UpdateDistinct fAE(idx), dAE(idx), vAE(i, 1)
        UpdateDistinct fAJ(idx), dAJ(idx), vAJ(i, 1)
        UpdateDistinct fAL(idx), dAL(idx), vAL(i, 1)
        UpdateDistinct fAM(idx), dAM(idx), vAM(i, 1)
        UpdateDistinct fAN(idx), dAN(idx), vAN(i, 1)
        UpdateDistinct fAO(idx), dAO(idx), vAO(i, 1)
        UpdateDistinct fBB(idx), dBB(idx), vBB(i, 1)
        UpdateDistinct fBC(idx), dBC(idx), vBC(i, 1)
        UpdateDistinct fBD(idx), dBD(idx), vBD(i, 1)
        UpdateDistinct fBE(idx), dBE(idx), vBE(i, 1)
        UpdateDistinct fBF(idx), dBF(idx), vBF(i, 1)
        UpdateDistinct fBG(idx), dBG(idx), vBG(i, 1)
        UpdateDistinct fBH(idx), dBH(idx), vBH(i, 1)
        UpdateDistinct fBI(idx), dBI(idx), vBI(i, 1)

        If IsNumeric(vAJ(i, 1)) Then
            Dim aj As Double: aj = CDbl(vAJ(i, 1))
            If aj < minAJ(idx) Then
                minAJ(idx) = aj
                minAJRow(idx) = i
            End If
            Dim ajEff As Double: ajEff = IIf(aj = 0, 99999, aj)
            If ajEff < minBEff(idx) Then
                minBEff(idx) = ajEff
                minBEffRow(idx) = i
            End If
        End If
        If IsNumeric(vAL(i, 1)) Then
            Dim alv As Double: alv = CDbl(vAL(i, 1))
            If alv > maxAL(idx) Then maxAL(idx) = alv
        End If
        If donorRow(idx) = 0 Then
            If UCase$(Trim$(CStr(vBB(i, 1)))) = "YES" Then donorRow(idx) = i
        End If
    Next i

    ReDim Preserve cnt(1 To k), firstRow(1 To k), minAJ(1 To k), minAJRow(1 To k), _
                    maxAL(1 To k), donorRow(1 To k), minBEff(1 To k), minBEffRow(1 To k), _
                    dAE(1 To k), dAJ(1 To k), dAL(1 To k), dAM(1 To k), dAN(1 To k), _
                    dAO(1 To k), dBB(1 To k), dBC(1 To k), dBD(1 To k), dBE(1 To k), _
                    dBF(1 To k), dBG(1 To k), dBH(1 To k), dBI(1 To k)
End Sub

Private Sub PopulateOutputRow(ByVal i As Long, ByRef outAP As Variant, asinIdx As Object, _
                              cnt() As Long, firstRow() As Long, minAJ() As Double, minAJRow() As Long, _
                              maxAL() As Double, donorRow() As Long, minBEff() As Double, minBEffRow() As Long, _
                              dAE() As Boolean, dAJ() As Boolean, dAL() As Boolean, dAM() As Boolean, _
                              dAN() As Boolean, dAO() As Boolean, dBB() As Boolean, dBC() As Boolean, _
                              dBD() As Boolean, dBE() As Boolean, dBF() As Boolean, dBG() As Boolean, _
                              dBH() As Boolean, dBI() As Boolean, vS As Variant, vAE As Variant, vAJ As Variant, _
                              vAL As Variant, vAM As Variant, vAN As Variant, vAO As Variant, vBB As Variant, _
                              vBC As Variant, vBD As Variant, vBE As Variant, vBF As Variant, vBG As Variant, _
                              vBH As Variant, vBI As Variant)

    Dim colA As Long, colB As Long, colC As Long, colD As Long, colE As Long, colF As Long, colG As Long
    Dim colH As Long, colI As Long, colJ As Long, colK As Long, colL As Long, colM As Long, colN As Long, colO As Long, colP As Long
    colA = ColLetterToNum("A")
    colB = ColLetterToNum("B")
    colC = ColLetterToNum("C")
    colD = ColLetterToNum("D")
    colE = ColLetterToNum("E")
    colF = ColLetterToNum("F")
    colG = ColLetterToNum("G")
    colH = ColLetterToNum("H")
    colI = ColLetterToNum("I")
    colJ = ColLetterToNum("J")
    colK = ColLetterToNum("K")
    colL = ColLetterToNum("L")
    colM = ColLetterToNum("M")
    colN = ColLetterToNum("N")
    colO = ColLetterToNum("O")
    colP = ColLetterToNum("P")

    Dim asin2 As String: asin2 = CStr(vS(i, 1))
    Dim id As Long: id = CLng(asinIdx(asin2))
    Dim pcount As Long: pcount = cnt(id)
    outAP(i, colP) = pcount

    Dim uniqAE As Boolean: uniqAE = dAE(id)
    Dim uniqAJ As Boolean: uniqAJ = dAJ(id)
    Dim uniqAL As Boolean: uniqAL = dAL(id)
    Dim uniqAM As Boolean: uniqAM = dAM(id)
    Dim uniqAN As Boolean: uniqAN = dAN(id)
    Dim uniqAO As Boolean: uniqAO = dAO(id)
    Dim uniqBB As Boolean: uniqBB = dBB(id)
    Dim uniqBC As Boolean: uniqBC = dBC(id)
    Dim uniqBD As Boolean: uniqBD = dBD(id)
    Dim uniqBE As Boolean: uniqBE = dBE(id)
    Dim uniqBF As Boolean: uniqBF = dBF(id)
    Dim uniqBG As Boolean: uniqBG = dBG(id)
    Dim uniqBH As Boolean: uniqBH = dBH(id)
    Dim uniqBI As Boolean: uniqBI = dBI(id)

    Dim Brow As Variant: Brow = "SKIP"
    If pcount > 1 And uniqAJ Then
        If minBEffRow(id) > 0 Then
            Brow = minBEff(id)
        End If
    End If
    outAP(i, colB) = Brow

    If pcount > 1 Then
        If uniqAE Then
            If UCase$(CStr(vAE(i, 1))) <> "YES" Then outAP(i, colA) = "Yes" Else outAP(i, colA) = "SKIP"
        Else
            outAP(i, colA) = "SKIP"
        End If
    Else
        outAP(i, colA) = "SKIP"
    End If

    If pcount > 1 Then
        If uniqAL Then
            If maxAL(id) > -1E+307 Then outAP(i, colC) = maxAL(id) Else outAP(i, colC) = "SKIP"
        Else
            outAP(i, colC) = "SKIP"
        End If
    Else
        outAP(i, colC) = "SKIP"
    End If

    Dim keyRow As Long: keyRow = IIf(minBEffRow(id) > 0, minBEffRow(id), minAJRow(id))
    If pcount > 1 Then
        If uniqAM Then
            If CStr(Brow) = "SKIP" Then
                outAP(i, colD) = "Product Sphere"
            Else
                outAP(i, colD) = vAM(keyRow, 1)
            End If
        Else
            outAP(i, colD) = "SKIP"
        End If
    Else
        outAP(i, colD) = "SKIP"
    End If

    If pcount > 1 Then
        If uniqAN Then
            If CStr(Brow) = "SKIP" Then
                outAP(i, colE) = "Increase Margin Maintain Unit Sales"
            Else
                outAP(i, colE) = vAN(keyRow, 1)
            End If
        Else
            outAP(i, colE) = "SKIP"
        End If
    Else
        outAP(i, colE) = "SKIP"
    End If

    If pcount > 1 Then
        If uniqAO Then
            If CStr(Brow) = "SKIP" Then
                outAP(i, colF) = ""
            Else
                outAP(i, colF) = vAO(keyRow, 1)
            End If
        Else
            outAP(i, colF) = "SKIP"
        End If
    Else
        outAP(i, colF) = "SKIP"
    End If

    If pcount > 1 Then
        If uniqBB Then
            If UCase$(CStr(vBB(i, 1))) <> "YES" Then outAP(i, colG) = "Yes" Else outAP(i, colG) = "SKIP"
        Else
            outAP(i, colG) = "SKIP"
        End If
    Else
        outAP(i, colG) = "SKIP"
    End If

    If pcount > 1 Then
        Dim fr As Long: fr = firstRow(id)
        Dim hVal As Variant, iVal As Variant, jVal As Variant, kVal As Variant
        Dim lVal As Variant, mVal As Variant, nVal As Variant
        If CStr(Brow) = "SKIP" Then
            hVal = IIf(uniqBC, vBC(fr, 1), "SKIP")
            iVal = IIf(uniqBD, vBD(fr, 1), "SKIP")
            jVal = IIf(uniqBE, vBE(fr, 1), "SKIP")
            kVal = IIf(uniqBF, vBF(fr, 1), "SKIP")
            lVal = IIf(uniqBG, vBG(fr, 1), "SKIP")
            mVal = IIf(uniqBH, vBH(fr, 1), "SKIP")
            nVal = IIf(uniqBI, vBI(fr, 1), "SKIP")
        Else
            hVal = IIf(uniqBC, vBC(keyRow, 1), "SKIP")
            iVal = IIf(uniqBD, vBD(keyRow, 1), "SKIP")
            jVal = IIf(uniqBE, vBE(keyRow, 1), "SKIP")
            kVal = IIf(uniqBF, vBF(keyRow, 1), "SKIP")
            lVal = IIf(uniqBG, vBG(keyRow, 1), "SKIP")
            mVal = IIf(uniqBH, vBH(keyRow, 1), "SKIP")
            nVal = IIf(uniqBI, vBI(keyRow, 1), "SKIP")
        End If
        If (pcount > 1) And uniqBB And (UCase$(CStr(vBB(i, 1))) <> "YES") And (donorRow(id) > 0) Then
            hVal = vBC(donorRow(id), 1)
            iVal = vBD(donorRow(id), 1)
            jVal = vBE(donorRow(id), 1)
            kVal = vBF(donorRow(id), 1)
            lVal = vBG(donorRow(id), 1)
            mVal = vBH(donorRow(id), 1)
            nVal = vBI(donorRow(id), 1)
        End If
        outAP(i, colH) = hVal
        outAP(i, colI) = iVal
        outAP(i, colJ) = jVal
        outAP(i, colK) = kVal
        outAP(i, colL) = lVal
        outAP(i, colM) = mVal
        outAP(i, colN) = nVal
    Else
        outAP(i, colH) = "SKIP"
        outAP(i, colI) = "SKIP"
        outAP(i, colJ) = "SKIP"
        outAP(i, colK) = "SKIP"
        outAP(i, colL) = "SKIP"
        outAP(i, colM) = "SKIP"
        outAP(i, colN) = "SKIP"
    End If

    Dim skipCount As Long: skipCount = 0
    Dim c As Long
    For c = colA To colN
        If UCase$(CStr(outAP(i, c))) = "SKIP" Or IsEmpty(outAP(i, c)) Then skipCount = skipCount + 1
    Next c
    If skipCount = (colN - colA + 1) Then
        outAP(i, colO) = "SKIP"
    Else
        outAP(i, colO) = "FILTER"
    End If
End Sub

Private Sub UpdateDistinct(ByRef firstVal As Variant, ByRef seenDifferent As Boolean, ByVal newVal As Variant)
    If IsEmpty(firstVal) Then
        firstVal = newVal
        Exit Sub
    End If
    If CStr(firstVal) <> CStr(newVal) Then seenDifferent = True
End Sub

' ========= EXPORT =========
Private Sub BuildFilteredExport(wsTool As Worksheet, pasteStartCellAddress As String)
    Dim startCell As Range: Set startCell = wsTool.Range(pasteStartCellAddress)

    Dim lastRow As Long, lastCol As Long
    UsedBounds wsTool, lastRow, lastCol
    If lastRow < 2 Then Exit Sub

    Dim colN As Long, colS As Long, colBB As Long, colAL As Long
    colN = ColLetterToNum("N")
    colS = ColLetterToNum("S")
    colBB = ColLetterToNum("BB")
    colAL = ColLetterToNum("AL")

    Dim dataFirstCol As Long: dataFirstCol = startCell.Column  ' Q
    Dim dataLastCol As Long:  dataLastCol = lastCol            ' rightmost used in tool
    Dim width As Long: width = dataLastCol - dataFirstCol + 1

    ' Create export workbook/sheet
    Dim wbOut As Workbook, wsOut As Worksheet
    Set wbOut = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)
    On Error Resume Next: wsOut.Name = EXPORT_SHEET_NAME: On Error GoTo 0

    ' Headers
    wsOut.Cells(1, 1).Resize(1, width).Value = wsTool.Cells(1, dataFirstCol).Resize(1, width).Value
    Dim maps As Variant: maps = MAPPING_PAIRS
    EnsureMappedHeadersFromTool wsTool, wsOut, maps

    Dim mapInfo As Variant
    ReDim mapInfo(LBound(maps) To UBound(maps))
    Dim mi As Long
    For mi = LBound(maps) To UBound(maps)
        mapInfo(mi) = Array( _
            CStr(maps(mi)(0)), CStr(maps(mi)(1)), _
            ColLetterToNum(CStr(maps(mi)(0))), _
            ColLetterToNum(CStr(maps(mi)(1))) _
        )
    Next mi

    Dim pairLetters As Variant
    pairLetters = Array( _
        Array("BC", "AM"), Array("BD", "AN"), Array("BE", "AO"), _
        Array("BF", "AP"), Array("BG", "AQ"), Array("BH", "AR"), Array("BI", "AS") _
    )
    Dim pairSrcIdx() As Long, pairDstIdx() As Long
    ReDim pairSrcIdx(LBound(pairLetters) To UBound(pairLetters))
    ReDim pairDstIdx(LBound(pairLetters) To UBound(pairLetters))
    For mi = LBound(pairLetters) To UBound(pairLetters)
        pairSrcIdx(mi) = ColLetterToNum(CStr(pairLetters(mi)(0)))
        pairDstIdx(mi) = ColLetterToNum(CStr(pairLetters(mi)(1)))
    Next mi

    ' Preload tool blocks
    Dim toolVals As Variant, filterVals As Variant
    toolVals = wsTool.Range("A2", wsTool.Cells(lastRow, colN)).Value2
    filterVals = wsTool.Range(FILTER_COL_LETTER & 2, FILTER_COL_LETTER & lastRow).Value2

    ' Donor map per ASIN (first BB=="Yes")
    Dim vS As Variant, vBB As Variant
    vS = wsTool.Range("S2", wsTool.Cells(lastRow, colS)).Value2
    vBB = wsTool.Range("BB2", wsTool.Cells(lastRow, colBB)).Value2
    Dim donorByAsin As Object: Set donorByAsin = CreateObject("Scripting.Dictionary"): donorByAsin.CompareMode = vbTextCompare
    Dim i As Long
    For i = 2 To lastRow
        Dim asin As String: asin = CStr(vS(i - 1, 1))
        If Len(asin) > 0 And Not donorByAsin.Exists(asin) Then
            If UCase$(Trim$(CStr(vBB(i - 1, 1)))) = "YES" Then donorByAsin(asin) = i
        End If
    Next i

    Dim r As Long, outRow As Long: outRow = 2
    For r = 2 To lastRow
        If UCase$(Trim$(CStr(filterVals(r - 1, 1)))) = "FILTER" Then
            ' 1) Copy Q:Last as-is
            wsOut.Cells(outRow, 1).Resize(1, width).Value = wsTool.Cells(r, dataFirstCol).Resize(1, width).Value

            ' 2) Overlay mapped values from tool A:N into destination columns unless SKIP (and add notes if changed)
            Dim m As Long
            For m = LBound(mapInfo) To UBound(mapInfo)
                Dim v As Variant: v = toolVals(r - 1, mapInfo(m)(2))
                If Not IsSkipValue(v) Then
                    Dim dc As Long: dc = mapInfo(m)(3)
                    Dim oldv As Variant: oldv = wsOut.Cells(outRow, dc).Value
                    If CStr(oldv) <> CStr(v) Then
                        wsOut.Cells(outRow, dc).Value = v
                        NoteReplace wsOut.Cells(outRow, dc), oldv, v, _
                                   "Source: Tool " & mapInfo(m)(0) & " ? Export " & mapInfo(m)(1)
                    Else
                        wsOut.Cells(outRow, dc).Value = v
                    End If
                End If
            Next m

            ' 3) If AL (mapped from G) is "Yes", force AM:AS from donor row's H..N
            If UCase$(Trim$(CStr(wsOut.Cells(outRow, colAL).Value))) = "YES" Then
                Dim asinCurr As String: asinCurr = CStr(wsTool.Cells(r, colS).Value)
                If Len(asinCurr) > 0 And donorByAsin.Exists(asinCurr) Then
                    Dim dRow As Long: dRow = CLng(donorByAsin(asinCurr))
                    Dim u As Long
                    For u = LBound(pairSrcIdx) To UBound(pairSrcIdx)
                        Dim dstC As Long: dstC = pairDstIdx(u)
                        Dim newVal As Variant: newVal = wsTool.Cells(dRow, pairSrcIdx(u)).Value
                        Dim prevVal As Variant: prevVal = wsOut.Cells(outRow, dstC).Value
                        If CStr(prevVal) <> CStr(newVal) Then
                            wsOut.Cells(outRow, dstC).Value = newVal
                            NoteReplace wsOut.Cells(outRow, dstC), prevVal, newVal, _
                                       "Source: Donor " & pairLetters(u)(0) & " ? Export " & pairLetters(u)(1) & " (AL=Yes)"
                        Else
                            wsOut.Cells(outRow, dstC).Value = newVal
                        End If
                    Next u

                End If
            End If

            outRow = outRow + 1
        End If
    Next r

    wsOut.Cells.EntireColumn.AutoFit

    Dim suggested As String: suggested = "PricingExport_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    With Application.FileDialog(msoFileDialogSaveAs)
        .InitialFileName = suggested
        If .Show = -1 Then
            wbOut.SaveAs .SelectedItems(1), FileFormat:=xlOpenXMLWorkbook
            MsgBox "Export saved: " & .SelectedItems(1), vbInformation
        Else
            MsgBox "Export left unsaved (workbook remains open).", vbInformation
        End If
    End With
End Sub

Private Sub EnsureMappedHeadersFromTool(wsTool As Worksheet, wsOut As Worksheet, mapPairs As Variant)
    Dim i As Long
    For i = LBound(mapPairs) To UBound(mapPairs)
        Dim toolCol As String: toolCol = CStr(mapPairs(i)(0))
        Dim destCol As String:  destCol = CStr(mapPairs(i)(1))
        Dim hdr As String: hdr = CStr(wsTool.Cells(1, ColLetterToNum(toolCol)).Value)
        If Len(hdr) > 0 Then wsOut.Cells(1, ColLetterToNum(destCol)).Value = hdr
    Next i
End Sub

Private Function IsSkipValue(v As Variant) As Boolean
    On Error GoTo SafeFalse
    If IsError(v) Then Go To SafeFalse
    IsSkipValue = (UCase$(Trim$(CStr(v))) = "SKIP")
    Exit Function
SafeFalse:
    IsSkipValue = False
End Function

' ========= NOTES (optional) =========
Private Sub AddCellNote(ByVal tgt As Range, ByVal msg As String)
    If DISABLE_NOTES Then Exit Sub

    ' Delete any existing legacy comment
    If Not tgt.Comment Is Nothing Then tgt.Comment.Delete

    ' Try legacy comment first
    Dim firstErr As Long
    On Error Resume Next
    tgt.AddComment msg
    firstErr = Err.Number
    On Error GoTo 0

    If firstErr <> 0 Then
        ' Fallback to threaded comment (newer Excel)
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

Private Sub NoteReplace(ByVal tgt As Range, ByVal oldVal As Variant, ByVal newVal As Variant, ByVal reason As String)
    If DISABLE_NOTES Then Exit Sub
    Dim o As String, n As String
    o = CStr(oldVal): n = CStr(newVal)
    Dim msg As String
    msg = "Replaced value" & vbCrLf & _
          "Old: " & o & vbCrLf & _
          "New: " & n & vbCrLf & _
          reason
    AddCellNote tgt, msg
End Sub

' ========= UTILITIES =========
Private Sub UsedBounds(ws As Worksheet, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim ur As Range
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then
        lastRow = 1: lastCol = 1
        Exit Sub
    End If
    lastRow = ur.Row + ur.Rows.Count - 1
    lastCol = ur.Column + ur.Columns.Count - 1
End Sub

Private Function ColLetterToNum(ByVal colLetter As String) As Long
    Dim i As Long, n As Long
    colLetter = UCase$(Trim$(colLetter))
    For i = 1 To Len(colLetter)
        n = n * 26 + (Asc(Mid$(colLetter, i, 1)) - 64)
    Next i
    ColLetterToNum = n
End Function

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

