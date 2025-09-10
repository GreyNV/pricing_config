Attribute VB_Name = "Module1"
'====================================================================
' VBA Pricing Tool  Clear & Upload/Process (No extra worksheets)
' Using optimized ASIN struct computation for A:P
'====================================================================
Option Explicit

' ========= COLUMN INDEX CACHE =========
Private Const COL_A_IDX As Long = 1
Private Const COL_B_IDX As Long = 2
Private Const COL_C_IDX As Long = 3
Private Const COL_D_IDX As Long = 4
Private Const COL_E_IDX As Long = 5
Private Const COL_F_IDX As Long = 6
Private Const COL_G_IDX As Long = 7
Private Const COL_H_IDX As Long = 8
Private Const COL_I_IDX As Long = 9
Private Const COL_J_IDX As Long = 10
Private Const COL_K_IDX As Long = 11
Private Const COL_L_IDX As Long = 12
Private Const COL_M_IDX As Long = 13
Private Const COL_N_IDX As Long = 14
Private Const COL_O_IDX As Long = 15
Private Const COL_P_IDX As Long = 16
Private Const COL_S_IDX As Long = 19
Private Const COL_T_IDX As Long = 20
Private Const COL_V_IDX As Long = 22
Private Const COL_W_IDX As Long = 23
Private Const COL_X_IDX As Long = 24
Private Const COL_Y_IDX As Long = 25
Private Const COL_AL_IDX As Long = 38
Private Const COL_AM_IDX As Long = 39
Private Const COL_AN_IDX As Long = 40
Private Const COL_AO_IDX As Long = 41
Private Const COL_AP_IDX As Long = 42
Private Const COL_AQ_IDX As Long = 43
Private Const COL_AR_IDX As Long = 44
Private Const COL_AS_IDX As Long = 45
Private Const COL_BB_IDX As Long = 54
Private Const COL_BC_IDX As Long = 55
Private Const COL_BD_IDX As Long = 56
Private Const COL_BE_IDX As Long = 57
Private Const COL_BF_IDX As Long = 58
Private Const COL_BG_IDX As Long = 59
Private Const COL_BH_IDX As Long = 60
Private Const COL_BI_IDX As Long = 61

' ========= USER CONFIG =========
Private Const TOOL_SHEET_NAME As String = "Pricing Configurations"
Private Const PASTE_START_CELL As String = "Q1"
Private Const FILTER_COL_LETTER As String = "O"
Private Const EXPORT_SHEET_NAME As String = "Pricing Configurations"

Private Const DISABLE_NOTES As Boolean = False
Private Const USE_VBA_COMPUTE As Boolean = True
Private Const DEBUG_LOG As Boolean = True

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
    If DEBUG_LOG Then Debug.Print "Btn_ClearPricingData: start"
    ClearPricingData ThisWorkbook.Worksheets(TOOL_SHEET_NAME)
    If DEBUG_LOG Then Debug.Print "Btn_ClearPricingData: after Clear"
    GoTo Finally
EH:
    MsgBox "Clear failed: " & Err.Description, vbExclamation
Finally:
    OptimizeEnd
    If DEBUG_LOG Then Debug.Print "Btn_ClearPricingData: finished"
End Sub

Public Sub Btn_UploadAndProcess()
    On Error GoTo EH
    OptimizeStart
    Dim wsTool As Worksheet
    Set wsTool = ThisWorkbook.Worksheets(TOOL_SHEET_NAME)
    If DEBUG_LOG Then Debug.Print "Btn_UploadAndProcess: start"
    ClearPricingData wsTool
    If DEBUG_LOG Then Debug.Print "Btn_UploadAndProcess: after clear"

    Dim srcPath As String
    srcPath = PickSourceWorkbookPath()
    Dim wbSrc As Workbook
    Set wbSrc = Nothing
    If Len(srcPath) = 0 Then GoTo Finally

    Set wbSrc = Workbooks.Open(Filename:=srcPath, ReadOnly:=True)
    If DEBUG_LOG Then Debug.Print "Btn_UploadAndProcess: source opened"
    ImportAllPricingConfigurationSheets wbSrc, wsTool, wsTool.Range(PASTE_START_CELL)
    If DEBUG_LOG Then Debug.Print "Btn_UploadAndProcess: import complete"

    Dim lastRow As Long, lastCol As Long
    Dim ub As Variant
    ub = UsedBounds(wsTool)
    lastRow = ub(0): lastCol = ub(1)

    If USE_VBA_COMPUTE Then
        If DEBUG_LOG Then Debug.Print "Btn_UploadAndProcess: compute start"
        ComputeDerivedColumns_AP wsTool, lastRow
        If DEBUG_LOG Then Debug.Print "Btn_UploadAndProcess: compute done"
    End If

    If DEBUG_LOG Then Debug.Print "Btn_UploadAndProcess: export start"
    BuildFilteredExport wsTool, PASTE_START_CELL
    If DEBUG_LOG Then Debug.Print "Btn_UploadAndProcess: export done"
    GoTo Finally
EH:
    MsgBox "Upload/Process failed: " & Err.Description, vbExclamation
Finally:
    If Not wbSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
    End If
    OptimizeEnd
    If DEBUG_LOG Then Debug.Print "Btn_UploadAndProcess: finished"
End Sub

' ========= CORE LOGIC =========
Private Sub ClearPricingData(ws As Worksheet)
    If DEBUG_LOG Then Debug.Print "ClearPricingData: ws=" & ws.Name
    Dim lastRow As Long, lastCol As Long
    Dim ub As Variant
    ub = UsedBounds(ws)
    lastRow = ub(0): lastCol = ub(1)
    If DEBUG_LOG Then Debug.Print "ClearPricingData: lastRow=" & lastRow & " lastCol=" & lastCol
    If lastRow < 3 Then Exit Sub
    ws.Range(ws.Rows(3), ws.Rows(lastRow)).ClearContents
    If DEBUG_LOG Then Debug.Print "ClearPricingData: cleared through row " & lastRow
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
    If DEBUG_LOG Then Debug.Print "ImportAllPricingConfigurationSheets: start"
    Dim nextPasteRow As Long
    nextPasteRow = pasteStart.Row
    Dim sh As Worksheet
    For Each sh In wbSrc.Worksheets
        If DEBUG_LOG Then Debug.Print "Import check: " & sh.Name
        If InStr(1, sh.Name, "Pricing Configurations", vbTextCompare) > 0 Then
            Dim rng As Range
            Set rng = SheetDataRange(sh)
            If Not rng Is Nothing Then
                wsTool.Cells(nextPasteRow, pasteStart.Column).Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value
                If DEBUG_LOG Then Debug.Print "Imported sheet " & sh.Name & " rows=" & rng.Rows.Count
                nextPasteRow = nextPasteRow + rng.Rows.Count
            End If
        End If
    Next sh
    If DEBUG_LOG Then Debug.Print "ImportAllPricingConfigurationSheets: done"
End Sub

Private Function SheetDataRange(ws As Worksheet, Optional lastRow As Long = 0, Optional lastCol As Long = 0) As Range
    On Error GoTo EH
    If lastRow = 0 Or lastCol = 0 Then
        Dim ub As Variant
        ub = UsedBounds(ws)
        lastRow = ub(0): lastCol = ub(1)
    End If
    If DEBUG_LOG Then Debug.Print "SheetDataRange: ws=" & ws.Name & " lastRow=" & lastRow & " lastCol=" & lastCol
    If lastRow = 1 And lastCol = 1 Then
        If Len(ws.Cells(1, 1).Value) = 0 Then Exit Function
    End If
    Set SheetDataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    Exit Function
EH:
    Set SheetDataRange = Nothing
    On Error GoTo 0
    Err.Raise Err.Number, "SheetDataRange", _
              "Failed to obtain data range for '" & ws.Name & "': " & Err.Description
End Function

' ========= ASIN STRUCT COMPUTE =========
Private Sub ComputeDerivedColumns_AP(ws As Worksheet, lastRow As Long)
    If lastRow < 2 Then Exit Sub
    If DEBUG_LOG Then Debug.Print "ComputeDerivedColumns_AP: lastRow=" & lastRow

    Dim n As Long: n = lastRow - 1
    If DEBUG_LOG Then Debug.Print "ComputeDerivedColumns_AP: n=" & n

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
    Dim keyAM() As Variant, keyAN() As Variant, keyAO() As Variant
    ' True when column values differ across rows for the ASIN
    Dim hasVarAE() As Boolean, hasVarAJ() As Boolean, hasVarAL() As Boolean, hasVarAM() As Boolean
    Dim hasVarAN() As Boolean, hasVarAO() As Boolean, hasVarBB() As Boolean, hasVarBC() As Boolean
    Dim hasVarBD() As Boolean, hasVarBE() As Boolean, hasVarBF() As Boolean, hasVarBG() As Boolean
    Dim hasVarBH() As Boolean, hasVarBI() As Boolean

    BuildAsinAggregates n, vS, vAE, vAJ, vAL, vAM, vAN, vAO, vBB, vBC, vBD, vBE, vBF, vBG, vBH, vBI, _
                        asinIdx, cnt, firstRow, minAJ, minAJRow, maxAL, donorRow, minBEff, minBEffRow, keyAM, keyAN, keyAO, _
                        hasVarAE, hasVarAJ, hasVarAL, hasVarAM, hasVarAN, hasVarAO, hasVarBB, hasVarBC, hasVarBD, hasVarBE, hasVarBF, hasVarBG, hasVarBH, hasVarBI
    If DEBUG_LOG Then Debug.Print "ComputeDerivedColumns_AP: aggregates built"

    Dim outAP As Variant: ReDim outAP(1 To n, 1 To COL_P_IDX)
    Dim i As Long
    For i = 1 To n
        If DEBUG_LOG Then Debug.Print "PopulateOutputRow: " & i
        PopulateOutputRow i, outAP, asinIdx, cnt, firstRow, minAJ, minAJRow, maxAL, donorRow, minBEff, minBEffRow, _
                          keyAM, keyAN, keyAO, _
                          hasVarAE, hasVarAJ, hasVarAL, hasVarAM, hasVarAN, hasVarAO, hasVarBB, hasVarBC, hasVarBD, hasVarBE, hasVarBF, hasVarBG, hasVarBH, hasVarBI, _
                          vS, vAE, vAJ, vAL, vAM, vAN, vAO, vBB, vBC, vBD, vBE, vBF, vBG, vBH, vBI
    Next i
    ws.Range("A2", ws.Cells(lastRow, COL_P_IDX)).Value = outAP
    If DEBUG_LOG Then Debug.Print "ComputeDerivedColumns_AP: output written"
End Sub

Private Sub BuildAsinAggregates(ByVal n As Long, vS As Variant, vAE As Variant, vAJ As Variant, vAL As Variant, _
                                vAM As Variant, vAN As Variant, vAO As Variant, vBB As Variant, vBC As Variant, vBD As Variant, _
                                vBE As Variant, vBF As Variant, vBG As Variant, vBH As Variant, vBI As Variant, _
                                ByRef asinIdx As Object, ByRef cnt() As Long, ByRef firstRow() As Long, _
                                ByRef minAJ() As Double, ByRef minAJRow() As Long, ByRef maxAL() As Double, _
                                ByRef donorRow() As Long, ByRef minBEff() As Double, ByRef minBEffRow() As Long, _
                                ByRef keyAM() As Variant, ByRef keyAN() As Variant, ByRef keyAO() As Variant, _
                                ByRef hasVarAE() As Boolean, ByRef hasVarAJ() As Boolean, ByRef hasVarAL() As Boolean, ByRef hasVarAM() As Boolean, _
                                ByRef hasVarAN() As Boolean, ByRef hasVarAO() As Boolean, ByRef hasVarBB() As Boolean, ByRef hasVarBC() As Boolean, _
                                ByRef hasVarBD() As Boolean, ByRef hasVarBE() As Boolean, ByRef hasVarBF() As Boolean, ByRef hasVarBG() As Boolean, _
                                ByRef hasVarBH() As Boolean, ByRef hasVarBI() As Boolean)

    Set asinIdx = CreateObject("Scripting.Dictionary")
    asinIdx.CompareMode = vbTextCompare

    If DEBUG_LOG Then Debug.Print "BuildAsinAggregates: start n=" & n

    Dim cap As Long: cap = 0
    Dim k As Long: k = 0
    Dim i As Long

    Dim fAE() As Variant, fAJ() As Variant, fAL() As Variant, fAM() As Variant
    Dim fAN() As Variant, fAO() As Variant, fBB() As Variant, fBC() As Variant
    Dim fBD() As Variant, fBE() As Variant, fBF() As Variant, fBG() As Variant
    Dim fBH() As Variant, fBI() As Variant

    For i = 1 To n
        If DEBUG_LOG Then Debug.Print "BuildAsinAggregates: row=" & i
        Dim s As String: s = CStr(vS(i, 1))
        If Not asinIdx.Exists(s) Then
            k = k + 1
            If k > cap Then
                cap = IIf(cap = 0, 256, cap * 2)
                ReDim Preserve cnt(1 To cap), firstRow(1 To cap), minAJ(1 To cap), minAJRow(1 To cap), maxAL(1 To cap), _
                                donorRow(1 To cap), minBEff(1 To cap), minBEffRow(1 To cap), _
                                keyAM(1 To cap), keyAN(1 To cap), keyAO(1 To cap)
                ReDim Preserve fAE(1 To cap), hasVarAE(1 To cap), fAJ(1 To cap), hasVarAJ(1 To cap), _
                                fAL(1 To cap), hasVarAL(1 To cap), fAM(1 To cap), hasVarAM(1 To cap), _
                                fAN(1 To cap), hasVarAN(1 To cap), fAO(1 To cap), hasVarAO(1 To cap), _
                                fBB(1 To cap), hasVarBB(1 To cap), fBC(1 To cap), hasVarBC(1 To cap), _
                                fBD(1 To cap), hasVarBD(1 To cap), fBE(1 To cap), hasVarBE(1 To cap), _
                                fBF(1 To cap), hasVarBF(1 To cap), fBG(1 To cap), hasVarBG(1 To cap), _
                                fBH(1 To cap), hasVarBH(1 To cap), fBI(1 To cap), hasVarBI(1 To cap)
            End If
            asinIdx(s) = k
            cnt(k) = 0
            firstRow(k) = i
            minAJ(k) = 1E+308
            minAJRow(k) = 0
            maxAL(k) = -1E+308
            minBEff(k) = 1E+308: minBEffRow(k) = 0
            keyAM(k) = Empty: keyAN(k) = Empty: keyAO(k) = Empty
            donorRow(k) = 0
            fAE(k) = Empty: fAJ(k) = Empty: fAL(k) = Empty
            fAM(k) = Empty: fAN(k) = Empty: fAO(k) = Empty
            fBB(k) = Empty: fBC(k) = Empty: fBD(k) = Empty
            fBE(k) = Empty: fBF(k) = Empty: fBG(k) = Empty
            fBH(k) = Empty: fBI(k) = Empty
        End If

        Dim idx As Long: idx = CLng(asinIdx(s))
        cnt(idx) = cnt(idx) + 1

        UpdateDistinct fAE(idx), hasVarAE(idx), vAE(i, 1)
        UpdateDistinct fAJ(idx), hasVarAJ(idx), vAJ(i, 1)
        UpdateDistinct fAL(idx), hasVarAL(idx), vAL(i, 1)
        UpdateDistinct fAM(idx), hasVarAM(idx), vAM(i, 1)
        UpdateDistinct fAN(idx), hasVarAN(idx), vAN(i, 1)
        UpdateDistinct fAO(idx), hasVarAO(idx), vAO(i, 1)
        UpdateDistinct fBB(idx), hasVarBB(idx), vBB(i, 1)
        UpdateDistinct fBC(idx), hasVarBC(idx), vBC(i, 1)
        UpdateDistinct fBD(idx), hasVarBD(idx), vBD(i, 1)
        UpdateDistinct fBE(idx), hasVarBE(idx), vBE(i, 1)
        UpdateDistinct fBF(idx), hasVarBF(idx), vBF(i, 1)
        UpdateDistinct fBG(idx), hasVarBG(idx), vBG(i, 1)
        UpdateDistinct fBH(idx), hasVarBH(idx), vBH(i, 1)
        UpdateDistinct fBI(idx), hasVarBI(idx), vBI(i, 1)

        If IsNumeric(vAJ(i, 1)) Then
            Dim aj As Double: aj = CDbl(vAJ(i, 1))
            If aj < minAJ(idx) Then
                minAJ(idx) = aj
                minAJRow(idx) = i
                If minBEffRow(idx) = 0 Then
                    keyAM(idx) = vAM(i, 1)
                    keyAN(idx) = vAN(i, 1)
                    keyAO(idx) = vAO(i, 1)
                End If
            End If
            Dim ajEff As Double: ajEff = IIf(aj = 0, 99999, aj)
            If ajEff < minBEff(idx) Then
                minBEff(idx) = ajEff
                minBEffRow(idx) = i
                keyAM(idx) = vAM(i, 1)
                keyAN(idx) = vAN(i, 1)
                keyAO(idx) = vAO(i, 1)
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
                    keyAM(1 To k), keyAN(1 To k), keyAO(1 To k), _
                    hasVarAE(1 To k), hasVarAJ(1 To k), hasVarAL(1 To k), hasVarAM(1 To k), hasVarAN(1 To k), _
                    hasVarAO(1 To k), hasVarBB(1 To k), hasVarBC(1 To k), hasVarBD(1 To k), hasVarBE(1 To k), _
                    hasVarBF(1 To k), hasVarBG(1 To k), hasVarBH(1 To k), hasVarBI(1 To k)
End Sub

Private Sub PopulateOutputRow(ByVal i As Long, ByRef outAP As Variant, asinIdx As Object, _
                              cnt() As Long, firstRow() As Long, minAJ() As Double, minAJRow() As Long, _
                              maxAL() As Double, donorRow() As Long, minBEff() As Double, minBEffRow() As Long, _
                              keyAM() As Variant, keyAN() As Variant, keyAO() As Variant, _
                              hasVarAE() As Boolean, hasVarAJ() As Boolean, hasVarAL() As Boolean, hasVarAM() As Boolean, _
                              hasVarAN() As Boolean, hasVarAO() As Boolean, hasVarBB() As Boolean, hasVarBC() As Boolean, _
                              hasVarBD() As Boolean, hasVarBE() As Boolean, hasVarBF() As Boolean, hasVarBG() As Boolean, _
                              hasVarBH() As Boolean, hasVarBI() As Boolean, vS As Variant, vAE As Variant, vAJ As Variant, _
                              vAL As Variant, vAM As Variant, vAN As Variant, vAO As Variant, vBB As Variant, vBC As Variant, _
                              vBD As Variant, vBE As Variant, vBF As Variant, vBG As Variant, vBH As Variant, vBI As Variant)

    Dim colA As Long, colB As Long, colC As Long, colD As Long, colE As Long, colF As Long, colG As Long
    Dim colH As Long, colI As Long, colJ As Long, colK As Long, colL As Long, colM As Long, colN As Long, colO As Long, colP As Long
    colA = COL_A_IDX
    colB = COL_B_IDX
    colC = COL_C_IDX
    colD = COL_D_IDX
    colE = COL_E_IDX
    colF = COL_F_IDX
    colG = COL_G_IDX
    colH = COL_H_IDX
    colI = COL_I_IDX
    colJ = COL_J_IDX
    colK = COL_K_IDX
    colL = COL_L_IDX
    colM = COL_M_IDX
    colN = COL_N_IDX
    colO = COL_O_IDX
    colP = COL_P_IDX

    Dim asin2 As String: asin2 = CStr(vS(i, 1))
    Dim id As Long: id = CLng(asinIdx(asin2))
    Dim pcount As Long: pcount = cnt(id)
    If DEBUG_LOG Then Debug.Print "PopulateOutputRow internal: row=" & i & " asin=" & asin2 & " count=" & pcount
    outAP(i, colP) = pcount

    If pcount = 1 Then
        Dim c As Long
        For c = colA To colN
            outAP(i, c) = "SKIP"
        Next c
        outAP(i, colO) = "SKIP"
        Exit Sub
    End If

    ' hasVarXX arrays are True when column values differ across rows for the ASIN

    Dim Brow As Variant: Brow = "SKIP"
    If hasVarAJ(id) Then
        If minBEffRow(id) > 0 Then
            Brow = minBEff(id)
        End If
    End If
    outAP(i, colB) = Brow

    Dim keyRow As Long
    keyRow = IIf(minBEffRow(id) > 0, minBEffRow(id), minAJRow(id))

    If hasVarAE(id) Then
        If UCase$(CStr(vAE(i, 1))) <> "YES" Then outAP(i, colA) = "Yes" Else outAP(i, colA) = "SKIP"
    Else
        outAP(i, colA) = "SKIP"
    End If

    If hasVarAL(id) Then
        If maxAL(id) > -1E+307 Then outAP(i, colC) = maxAL(id) Else outAP(i, colC) = "SKIP"
    Else
        outAP(i, colC) = "SKIP"
    End If

    If hasVarAM(id) Then
        If CStr(Brow) = "SKIP" Then
            Dim dVal As Variant: dVal = vAM(keyRow, 1)
            If Len(Trim$(CStr(dVal))) > 0 And UCase$(CStr(dVal)) <> "SKIP" Then
                outAP(i, colD) = dVal
            Else
                outAP(i, colD) = "Product Sphere"
            End If
        Else
            outAP(i, colD) = keyAM(id)
        End If
    Else
        outAP(i, colD) = "SKIP"
    End If

    ' Columns E and F depend on the final decision for D.
    If hasVarAN(id) Then
        If CStr(Brow) = "SKIP" Then
            Dim eVal As Variant: eVal = vAN(keyRow, 1)
            If Len(Trim$(CStr(eVal))) > 0 And UCase$(CStr(eVal)) <> "SKIP" Then
                outAP(i, colE) = eVal
            Else
                outAP(i, colE) = "Increase Margin Maintain Unit Sales"
            End If
        Else
            outAP(i, colE) = keyAN(id)
        End If
    Else
        outAP(i, colE) = "SKIP"
    End If

    If hasVarAO(id) Then
        If CStr(Brow) = "SKIP" Then
            Dim fVal As Variant: fVal = vAO(keyRow, 1)
            If Len(Trim$(CStr(fVal))) > 0 And UCase$(CStr(fVal)) <> "SKIP" Then
                outAP(i, colF) = fVal
            Else
                outAP(i, colF) = ""
            End If
        Else
            outAP(i, colF) = keyAO(id)
        End If
    Else
        outAP(i, colF) = "SKIP"
    End If

    Dim validDonor As Boolean
    validDonor = False
    If donorRow(id) > 0 Then
        Dim donorDate As Variant: donorDate = vBH(donorRow(id), 1)
        If IsFutureDate(donorDate) Then validDonor = True
    End If

    Dim fr As Long: fr = firstRow(id)
    Dim hVal As Variant, iVal As Variant, jVal As Variant, kVal As Variant
    Dim lVal As Variant, mVal As Variant, nVal As Variant
    If CStr(Brow) = "SKIP" Then
        hVal = IIf(hasVarBC(id), vBC(fr, 1), "SKIP")
        iVal = IIf(hasVarBD(id), vBD(fr, 1), "SKIP")
        jVal = IIf(hasVarBE(id), vBE(fr, 1), "SKIP")
        kVal = IIf(hasVarBF(id), vBF(fr, 1), "SKIP")
        lVal = IIf(hasVarBG(id), vBG(fr, 1), "SKIP")
        mVal = IIf(hasVarBH(id), vBH(fr, 1), "SKIP")
        nVal = IIf(hasVarBI(id), vBI(fr, 1), "SKIP")
    Else
        hVal = IIf(hasVarBC(id), vBC(keyRow, 1), "SKIP")
        iVal = IIf(hasVarBD(id), vBD(keyRow, 1), "SKIP")
        jVal = IIf(hasVarBE(id), vBE(keyRow, 1), "SKIP")
        kVal = IIf(hasVarBF(id), vBF(keyRow, 1), "SKIP")
        lVal = IIf(hasVarBG(id), vBG(keyRow, 1), "SKIP")
        mVal = IIf(hasVarBH(id), vBH(keyRow, 1), "SKIP")
        nVal = IIf(hasVarBI(id), vBI(keyRow, 1), "SKIP")
    End If

    If hasVarBB(id) And validDonor And (UCase$(CStr(vBB(i, 1))) <> "YES") Then
        hVal = vBC(donorRow(id), 1)
        iVal = vBD(donorRow(id), 1)
        jVal = vBE(donorRow(id), 1)
        kVal = Date + 1
        lVal = vBG(donorRow(id), 1)
        Dim saleEnd As Variant: saleEnd = vBH(donorRow(id), 1)
        If IsFutureDate(saleEnd) Then
            mVal = saleEnd
        Else
            mVal = "SKIP"
        End If
        nVal = vBI(donorRow(id), 1)
    End If
    outAP(i, colG) = "SKIP"
    outAP(i, colH) = hVal
    outAP(i, colI) = iVal
    outAP(i, colJ) = jVal
    outAP(i, colK) = kVal
    outAP(i, colL) = lVal
    outAP(i, colM) = mVal
    outAP(i, colN) = nVal

    Dim skipCount As Long: skipCount = 0
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
    If DEBUG_LOG Then Debug.Print "BuildFilteredExport: start"
    Dim startCell As Range: Set startCell = wsTool.Range(pasteStartCellAddress)

    Dim lastRow As Long, lastCol As Long
    Dim ub As Variant
    ub = UsedBounds(wsTool)
    lastRow = ub(0): lastCol = ub(1)
    If DEBUG_LOG Then Debug.Print "BuildFilteredExport: lastRow=" & lastRow & " lastCol=" & lastCol
    If lastRow < 2 Then Exit Sub

    Dim colN As Long, colS As Long, colBB As Long, colAL As Long
    Dim colBC As Long, colBI As Long, colFilter As Long
    colN = COL_N_IDX
    colS = COL_S_IDX
    colBB = COL_BB_IDX
    colAL = COL_AL_IDX
    colBC = COL_BC_IDX
    colBI = COL_BI_IDX
    colFilter = ColIndex(FILTER_COL_LETTER)

    Dim firstCol As Long: firstCol = startCell.Column  ' Q
    Dim dataLastCol As Long: dataLastCol = Application.Max(lastCol, COL_AS_IDX)
    Dim maps As Variant: maps = MAPPING_PAIRS
    Dim mapInfo As Variant
    ReDim mapInfo(LBound(maps) To UBound(maps))
    Dim mi As Long
    For mi = LBound(maps) To UBound(maps)
        mapInfo(mi) = Array( _
            CStr(maps(mi)(0)), CStr(maps(mi)(1)), _
            ColIndex(CStr(maps(mi)(0))), _
            ColIndex(CStr(maps(mi)(1))) _
        )
    Next mi

    Dim pairLetters As Variant
    pairLetters = Array( _
        Array("BC", "AM"), Array("BD", "AN"), Array("BE", "AO"), _
        Array("BF", "AP"), Array("BG", "AQ"), Array("BH", "AR"), Array("BI", "AS") _
    )
    Dim pairSrcIdx As Variant, pairDstIdx As Variant, pairSrcOffset() As Long
    Dim bhIdx As Long
    pairSrcIdx = Array(COL_BC_IDX, COL_BD_IDX, COL_BE_IDX, COL_BF_IDX, COL_BG_IDX, COL_BH_IDX, COL_BI_IDX)
    pairDstIdx = Array(COL_AM_IDX, COL_AN_IDX, COL_AO_IDX, COL_AP_IDX, COL_AQ_IDX, COL_AR_IDX, COL_AS_IDX)
    ReDim pairSrcOffset(LBound(pairSrcIdx) To UBound(pairSrcIdx))
    For mi = LBound(pairSrcIdx) To UBound(pairSrcIdx)
        pairSrcOffset(mi) = pairSrcIdx(mi) - colBC + 1
        If pairSrcIdx(mi) = COL_BH_IDX Then bhIdx = mi
    Next mi

    Dim maxDestCol As Long: maxDestCol = dataLastCol
    For mi = LBound(mapInfo) To UBound(mapInfo)
        maxDestCol = Application.Max(maxDestCol, CLng(mapInfo(mi)(3)))
    Next mi
    For mi = LBound(pairDstIdx) To UBound(pairDstIdx)
        maxDestCol = Application.Max(maxDestCol, CLng(pairDstIdx(mi)))
    Next mi
    If maxDestCol > dataLastCol Then dataLastCol = maxDestCol
    Dim width As Long
    width = dataLastCol - firstCol + 1
    If DEBUG_LOG Then Debug.Print "BuildFilteredExport: width=" & width
    If width < 1 Then Exit Sub

    Dim wbOut As Workbook, wsOut As Worksheet
    Set wbOut = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)
    On Error Resume Next: wsOut.Name = EXPORT_SHEET_NAME: On Error GoTo 0

    wsOut.Cells(1, 1).Resize(1, width).Value = wsTool.Cells(1, firstCol).Resize(1, width).Value
    FillMappedHeaderBlanksFromTool wsTool, wsOut, maps, firstCol, width

    ' Preload tool blocks
    Dim toolVals As Variant, filterVals As Variant, tailVals As Variant
    toolVals = wsTool.Range("A2", wsTool.Cells(lastRow, colN)).Value2
    filterVals = wsTool.Range(wsTool.Cells(2, colFilter), wsTool.Cells(lastRow, colFilter)).Value2
    tailVals = wsTool.Range(wsTool.Cells(2, firstCol), wsTool.Cells(lastRow, dataLastCol)).Value2

    ' Donor map per ASIN (first BB=="Yes")
    Dim vS As Variant, vBB As Variant, donorSrcVals As Variant
    vS = wsTool.Range("S2", wsTool.Cells(lastRow, colS)).Value2
    vBB = wsTool.Range("BB2", wsTool.Cells(lastRow, colBB)).Value2
    donorSrcVals = wsTool.Range(wsTool.Cells(2, colBC), wsTool.Cells(lastRow, colBI)).Value2
    Dim donorByAsin As Object: Set donorByAsin = CreateObject("Scripting.Dictionary"): donorByAsin.CompareMode = vbTextCompare
    Dim i As Long
    For i = 1 To UBound(vS, 1)
        Dim asin As String: asin = CStr(vS(i, 1))
        If Len(asin) > 0 And Not donorByAsin.Exists(asin) Then
            If UCase$(Trim$(CStr(vBB(i, 1)))) = "YES" Then donorByAsin(asin) = i
        End If
    Next i

    Dim outArr As Variant: ReDim outArr(1 To lastRow - 1, 1 To width)
    Dim notes As Collection: Set notes = New Collection
    Dim errLog As Collection: Set errLog = New Collection

    Dim r As Long, outIdx As Long
    Dim alRel As Long: alRel = colAL - firstCol + 1
    Dim asinCurr As String
    For r = 1 To UBound(filterVals, 1)
        asinCurr = CStr(vS(r, 1))
        On Error GoTo RowErr
        If DEBUG_LOG Then Debug.Print "BuildFilteredExport: processing row " & r
        If UCase$(Trim$(CStr(filterVals(r, 1)))) = "FILTER" Then
            outIdx = outIdx + 1
            If DEBUG_LOG Then Debug.Print "BuildFilteredExport: exporting row " & r

            ' 1) Copy base columns as-is
            For i = firstCol To dataLastCol
                outArr(outIdx, i - firstCol + 1) = tailVals(r, i - firstCol + 1)
            Next i

            ' 2) Overlay mapped values from tool A:N into destination columns unless SKIP (and add notes if changed)
            Dim m As Long
            For m = LBound(mapInfo) To UBound(mapInfo)
                Dim scRel As Long: scRel = mapInfo(m)(2)
                If scRel >= 1 And scRel <= colN Then
                    Dim v As Variant: v = toolVals(r, scRel)
                    If Not IsSkipValue(v) Then
                        Dim dcRel As Long: dcRel = mapInfo(m)(3) - firstCol + 1
                        If dcRel >= 1 And dcRel <= width Then
                            Dim oldv As Variant: oldv = outArr(outIdx, dcRel)
                            If CStr(oldv) <> CStr(v) Then
                                outArr(outIdx, dcRel) = v
                                notes.Add Array(outIdx + 1, dcRel, oldv, v, _
                                                "Source: Tool " & mapInfo(m)(0) & " ? Export " & mapInfo(m)(1))
                            Else
                                outArr(outIdx, dcRel) = v
                            End If
                        End If
                    End If
                End If
            Next m

            ' 3) If AL (mapped from G) is "Yes", force AM:AS from donor row's H..N
            If alRel >= 1 And alRel <= width Then
                If UCase$(Trim$(CStr(outArr(outIdx, alRel)))) = "YES" Then
                    If Len(asinCurr) > 0 And donorByAsin.Exists(asinCurr) Then
                        Dim dIdx As Long: dIdx = CLng(donorByAsin(asinCurr))
                        Dim donorEnd As Variant: donorEnd = donorSrcVals(dIdx, pairSrcOffset(bhIdx))
                        Dim validDonor As Boolean: validDonor = IsFutureDate(donorEnd)
                        If validDonor Then
                            Dim u As Long
                            For u = LBound(pairSrcIdx) To UBound(pairSrcIdx)
                                Dim dstC As Long: dstC = pairDstIdx(u)
                                Dim dstRel As Long: dstRel = dstC - firstCol + 1
                                Dim newVal As Variant
                                If pairSrcIdx(u) = COL_BF_IDX Then
                                    newVal = Date + 1
                                ElseIf pairSrcIdx(u) = COL_BH_IDX Then
                                    newVal = donorEnd
                                Else
                                    newVal = donorSrcVals(dIdx, pairSrcOffset(u))
                                End If
                                If dstRel >= 1 And dstRel <= width Then
                                    Dim prevVal As Variant: prevVal = outArr(outIdx, dstRel)
                                    If CStr(prevVal) <> CStr(newVal) Then
                                        outArr(outIdx, dstRel) = newVal
                                        notes.Add Array(outIdx + 1, dstRel, prevVal, newVal, _
                                                        "Source: Donor " & pairLetters(u)(0) & " ? Export " & pairLetters(u)(1) & " (AL=Yes)")
                                    Else
                                        outArr(outIdx, dstRel) = newVal
                                    End If
                                End If
                            Next u
                        End If
                    End If
                End If
            End If
        End If ' Filter
RowNext:
        On Error GoTo 0
    Next r

    If DEBUG_LOG Then Debug.Print "BuildFilteredExport: total exported=" & outIdx
    If outIdx = 0 Then GoTo Finish
    Dim finalArr() As Variant
    ReDim finalArr(1 To outIdx, 1 To width)
    Dim c As Long
    For r = 1 To outIdx
        For c = 1 To width
            finalArr(r, c) = outArr(r, c)
        Next c
    Next r
    wsOut.Range("A2").Resize(outIdx, width).Value = finalArr

    Dim ni As Variant
    For Each ni In notes
        NoteReplace wsOut.Cells(ni(0), ni(1)), ni(2), ni(3), CStr(ni(4))
    Next ni

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

Finish:
    If errLog.Count > 0 Then
        Dim errItem As Variant, msg As String
        msg = "Rows skipped due to errors:" & vbCrLf
        For Each errItem In errLog
            msg = msg & errItem & vbCrLf
        Next errItem
        MsgBox msg, vbExclamation, "BuildFilteredExport"
    End If
    Exit Sub

RowErr:
    errLog.Add "ASIN " & asinCurr & ": " & Err.Description
    Err.Clear
    Resume RowNext
End Sub

Private Sub FillMappedHeaderBlanksFromTool(wsTool As Worksheet, wsOut As Worksheet, mapPairs As Variant, firstCol As Long, width As Long)
    Dim i As Long
    For i = LBound(mapPairs) To UBound(mapPairs)
        Dim toolCol As String: toolCol = CStr(mapPairs(i)(0))
        Dim destCol As String: destCol = CStr(mapPairs(i)(1))
        Dim destRel As Long: destRel = ColIndex(destCol) - firstCol + 1
        If destRel >= 1 And destRel <= width Then
            If Len(wsOut.Cells(1, destRel).Value) = 0 Then
                Dim hdr As String: hdr = CStr(wsTool.Cells(1, ColIndex(toolCol)).Value)
                If Len(hdr) > 0 Then wsOut.Cells(1, destRel).Value = hdr
            End If
        End If
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
    If DEBUG_LOG Then Debug.Print "AddCellNote: " & tgt.Address

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
    If DEBUG_LOG Then Debug.Print "NoteReplace: " & tgt.Address
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
Private Function UsedBounds(ws As Worksheet) As Variant
    If DEBUG_LOG Then Debug.Print "UsedBounds: ws=" & ws.Name
    Dim lastCell As Range
    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlValues, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If lastCell Is Nothing Then
        If DEBUG_LOG Then Debug.Print "UsedBounds: empty"
        UsedBounds = Array(1, 1)
        Exit Function
    End If
    Dim lastRow As Long, lastCol As Long
    lastRow = lastCell.Row
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlValues, _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then
        lastCol = 1
    Else
        lastCol = lastCell.Column
    End If
    If DEBUG_LOG Then Debug.Print "UsedBounds: lastRow=" & lastRow & " lastCol=" & lastCol
    UsedBounds = Array(lastRow, lastCol)
End Function

Private Function ColLetterToNum(ByVal colLetter As String) As Long
    Dim i As Long, n As Long
    colLetter = UCase$(Trim$(colLetter))
    For i = 1 To Len(colLetter)
        n = n * 26 + (Asc(Mid$(colLetter, i, 1)) - 64)
    Next i
    ColLetterToNum = n
End Function

Private Function ColIndex(ByVal colLetter As String) As Long
    Select Case UCase$(colLetter)
        Case "A": ColIndex = COL_A_IDX
        Case "B": ColIndex = COL_B_IDX
        Case "C": ColIndex = COL_C_IDX
        Case "D": ColIndex = COL_D_IDX
        Case "E": ColIndex = COL_E_IDX
        Case "F": ColIndex = COL_F_IDX
        Case "G": ColIndex = COL_G_IDX
        Case "H": ColIndex = COL_H_IDX
        Case "I": ColIndex = COL_I_IDX
        Case "J": ColIndex = COL_J_IDX
        Case "K": ColIndex = COL_K_IDX
        Case "L": ColIndex = COL_L_IDX
        Case "M": ColIndex = COL_M_IDX
        Case "N": ColIndex = COL_N_IDX
        Case "O": ColIndex = COL_O_IDX
        Case "P": ColIndex = COL_P_IDX
        Case "S": ColIndex = COL_S_IDX
        Case "T": ColIndex = COL_T_IDX
        Case "V": ColIndex = COL_V_IDX
        Case "W": ColIndex = COL_W_IDX
        Case "X": ColIndex = COL_X_IDX
        Case "Y": ColIndex = COL_Y_IDX
        Case "AL": ColIndex = COL_AL_IDX
        Case "AM": ColIndex = COL_AM_IDX
        Case "AN": ColIndex = COL_AN_IDX
        Case "AO": ColIndex = COL_AO_IDX
        Case "AP": ColIndex = COL_AP_IDX
        Case "AQ": ColIndex = COL_AQ_IDX
        Case "AR": ColIndex = COL_AR_IDX
        Case "AS": ColIndex = COL_AS_IDX
        Case "BB": ColIndex = COL_BB_IDX
        Case "BC": ColIndex = COL_BC_IDX
        Case "BD": ColIndex = COL_BD_IDX
        Case "BE": ColIndex = COL_BE_IDX
        Case "BF": ColIndex = COL_BF_IDX
        Case "BG": ColIndex = COL_BG_IDX
        Case "BH": ColIndex = COL_BH_IDX
        Case "BI": ColIndex = COL_BI_IDX
        Case Else
            ColIndex = ColLetterToNum(colLetter)
    End Select
End Function

Private Function IsFutureDate(v As Variant) As Boolean
    IsFutureDate = IsDate(v) And CDate(v) > Date
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

