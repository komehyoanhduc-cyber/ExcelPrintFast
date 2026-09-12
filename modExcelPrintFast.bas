Attribute VB_Name = "modExcelPrintFast"
Option Explicit

Public Type TPrintOptions
    useSelection As Boolean
    PaperSizeName As String
    OrientationMode As String
    BlackAndWhite As Boolean
    CenterHorizontally As Boolean
    RepeatHeader As Boolean
    ExpandRowHeight As Boolean
    KeepSignatureWithData As Boolean
    SignatureDataRows As Long
    FooterEnabled As Boolean
    FooterLeft As String
    FooterCenter As String
    footerRight As String
    IncludePrintDate As Boolean
    PrintGridlines As Boolean
    PrintHeadings As Boolean
    PrintComments As Boolean
    SkipHiddenSheets As Boolean
    PreviewAfterSetup As Boolean
    MarginMode As String
    KeepManualBreaks As Boolean
    AutoLastPageBreak As Boolean
End Type

'FIXED: Ham kiem tra xem value co phai la so khong
Public Function IsNumeric(ByVal value As Variant) As Boolean
    On Error Resume Next
    IsNumeric = (value = CDbl(value))
    On Error GoTo 0
End Function

Public Sub ShowExcelPrintFast()
    frmExcelPrintFast.Show
End Sub

Public Sub SetDefaultOptions(ByRef opt As TPrintOptions)
    With opt
        .useSelection = False
        .PaperSizeName = "A4"
        .OrientationMode = "Auto"
        .BlackAndWhite = True
        .CenterHorizontally = True
        .RepeatHeader = True
        .ExpandRowHeight = True
        .KeepSignatureWithData = True
        .SignatureDataRows = 3
        .FooterEnabled = True
        .FooterLeft = "&F"
        .FooterCenter = "&A"
        .footerRight = "Trang &P/&N"
        .IncludePrintDate = False
        .PrintGridlines = False
        .PrintHeadings = False
        .PrintComments = False
        .SkipHiddenSheets = True
        .PreviewAfterSetup = True
        .MarginMode = "Standard"
        .KeepManualBreaks = True
        .AutoLastPageBreak = True
    End With
End Sub

Public Sub ProcessSelectedSheets(ByRef opt As TPrintOptions, ByVal sheetNames As Collection, ByVal showPreview As Boolean)
    Dim wb As Workbook, item As Variant, ws As Worksheet
    Dim result As String, successCount As Long, skippedCount As Long
    Dim oldCalculation As XlCalculation
    Dim totalSheets As Long, currentCount As Long
    Dim pct As Double
    
    If sheetNames Is Nothing Then Exit Sub
    If sheetNames.Count = 0 Then
        MsgBox "Hay chon it nhat mot sheet.", vbExclamation, "Thiet lap in nhanh"
        Exit Sub
    End If
    
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    
    totalSheets = sheetNames.Count
    currentCount = 0
    
    oldCalculation = Application.Calculation
    On Error GoTo SafeExit
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    For Each item In sheetNames
        currentCount = currentCount + 1
        pct = currentCount / totalSheets
        
        Application.StatusBar = "Dang xu ly thiet lap in: " & VBA.Format(pct, "0%") & " (" & currentCount & "/" & totalSheets & " sheet)..."
        DoEvents
        
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(item))
        On Error GoTo SafeExit
        
        If Not ws Is Nothing Then
            If opt.SkipHiddenSheets And ws.Visible <> xlSheetVisible Then
                skippedCount = skippedCount + 1
            Else
                result = result & ProcessOneSheet(ws, opt) & vbCrLf
                successCount = successCount + 1
            End If
        End If
    Next item
    
SafeExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "Khong the hoan tat: " & Err.Description, vbExclamation, "Thiet lap in nhanh"
        Err.Clear
        Exit Sub
    End If
    
    If showPreview And successCount > 0 Then
        On Error Resume Next
        wb.Worksheets(CStr(sheetNames(1))).Activate
        ActiveSheet.PrintPreview
        On Error GoTo 0
    End If
    
    If skippedCount > 0 Then result = result & "Bo qua " & skippedCount & " sheet an." & vbCrLf
    MsgBox "Da thiet lap " & successCount & " sheet.", vbInformation, "Thiet lap in nhanh"
End Sub

'FIXED: Wrapper sub de ProcessOneSheetDirect co the duoc goi tu form
Public Sub ProcessOneSheetDirect(ByVal ws As Worksheet, ByRef opt As TPrintOptions)
    ProcessOneSheet ws, opt
End Sub

Private Function ProcessOneSheet(ByVal ws As Worksheet, ByRef opt As TPrintOptions) As String
    Dim tableRange As Range, headerRow As Long, signatureRow As Long
    Set tableRange = DetectMainTableRange(ws, opt.useSelection)
    If tableRange Is Nothing Then
        ProcessOneSheet = ws.Name & ": khong tim thay bang du lieu."
        Exit Function
    End If
    
    SaveCurrentPrintState ws
    headerRow = DetectHeaderRow(tableRange)
    signatureRow = DetectSignatureStart(tableRange)
    If opt.ExpandRowHeight Then ExpandRowsSafely tableRange
    ApplyPageSetup ws, tableRange, headerRow, signatureRow, opt
    
    ProcessOneSheet = ws.Name & ": " & tableRange.Address(False, False) & _
        IIf(headerRow > 0, ", header dong " & headerRow & " den " & (headerRow + 1), ", khong nhan dien header") & _
        IIf(signatureRow > 0, ", co khoi ky", "")
End Function

Public Sub RestorePreviousPrintSetup()
    Dim wb As Workbook, backup As Worksheet, lastRow As Long, r As Long, ws As Worksheet
    Dim restored As Long, breakRows As Variant, breakRow As Variant
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set backup = wb.Worksheets("__PrintFast_LastState")
    On Error GoTo 0
    If backup Is Nothing Then
        MsgBox "Chua co thiet lap nao de khoi phuc.", vbInformation, "Thiet lap in nhanh"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    lastRow = backup.Cells(backup.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(backup.Cells(r, 1).Value))
        On Error GoTo 0
        If Not ws Is Nothing Then
            With ws.PageSetup
                .PrintArea = CStr(backup.Cells(r, 2).Value)
                .PrintTitleRows = CStr(backup.Cells(r, 3).Value)
                .Orientation = CLng(backup.Cells(r, 4).Value)
                .BlackAndWhite = CBool(backup.Cells(r, 5).Value)
                .CenterHorizontally = CBool(backup.Cells(r, 6).Value)
                .Zoom = backup.Cells(r, 7).Value
                .FitToPagesWide = backup.Cells(r, 8).Value
                .FitToPagesTall = backup.Cells(r, 9).Value
                .LeftMargin = CDbl(backup.Cells(r, 10).Value)
                .RightMargin = CDbl(backup.Cells(r, 11).Value)
                .TopMargin = CDbl(backup.Cells(r, 12).Value)
                .BottomMargin = CDbl(backup.Cells(r, 13).Value)
                .LeftFooter = CStr(backup.Cells(r, 14).Value)
                .CenterFooter = CStr(backup.Cells(r, 15).Value)
                .RightFooter = CStr(backup.Cells(r, 16).Value)
            End With
            ws.ResetAllPageBreaks
            If Len(CStr(backup.Cells(r, 17).Value)) > 0 Then
                breakRows = Split(CStr(backup.Cells(r, 17).Value), ",")
                For Each breakRow In breakRows
                    If IsNumeric(breakRow) Then ws.HPageBreaks.Add Before:=ws.Rows(CLng(breakRow))
                Next breakRow
            End If
            restored = restored + 1
        End If
    Next r
    Application.ScreenUpdating = True
    MsgBox "Da khoi phuc thiet lap in cho " & restored & " sheet.", vbInformation, "Thiet lap in nhanh"
End Sub

Private Sub SaveCurrentPrintState(ByVal ws As Worksheet)
    Dim backup As Worksheet, r As Long, manualBreakRows As String, pb As HPageBreak
    Set backup = GetBackupSheet(True)
    If backup Is Nothing Then Exit Sub
    On Error Resume Next
    backup.Rows(backup.Columns(1).Find(What:=ws.Name, LookIn:=xlValues, LookAt:=xlWhole).Row).Delete
    On Error GoTo 0
    r = backup.Cells(backup.Rows.Count, 1).End(xlUp).Row + 1
    For Each pb In ws.HPageBreaks
        If pb.Type = xlPageBreakManual Then manualBreakRows = manualBreakRows & IIf(Len(manualBreakRows) > 0, ",", "") & pb.Location.Row
    Next pb
    With ws.PageSetup
        backup.Cells(r, 1).Resize(1, 17).Value = Array(ws.Name, .PrintArea, .PrintTitleRows, .Orientation, .BlackAndWhite, .CenterHorizontally, .Zoom, .FitToPagesWide, .FitToPagesTall, .LeftMargin, .RightMargin, .TopMargin, .BottomMargin, .LeftFooter, .CenterFooter, .RightFooter, manualBreakRows)
    End With
    backup.Visible = xlSheetVeryHidden
End Sub

Private Function DetectMainTableRange(ByVal ws As Worksheet, ByVal useSelection As Boolean) As Range
    Dim lastRow As Long, endCol As Long
    Dim cell As Range, i As Long
    Dim titleFound As Boolean
    
    If useSelection Then
        If TypeName(Selection) = "Range" Then
            If Selection.Worksheet Is ws Then
                If Application.WorksheetFunction.CountA(Selection) > 0 Then
                    Set DetectMainTableRange = Selection
                    Exit Function
                End If
            End If
        End If
    End If
    
    On Error Resume Next
    lastRow = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    If lastRow < 1 Then Exit Function
    
    endCol = 1
    titleFound = False
    
    For i = 1 To IIf(lastRow < 15, lastRow, 15)
        For Each cell In ws.Rows(i).Cells
            If cell.Column > 40 Then Exit For
            If cell.MergeCells Then
                If cell.MergeArea.Columns.Count > 1 Then
                    Dim mEndCol As Long
                    mEndCol = cell.MergeArea.Column + cell.MergeArea.Columns.Count - 1
                    If mEndCol > endCol Then
                        endCol = mEndCol
                        titleFound = True
                    End If
                End If
            End If
        Next cell
    Next i
    
    If Not titleFound Or endCol = 1 Then
        On Error Resume Next
        endCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        On Error GoTo 0
        If endCol < 1 Then endCol = 1
    End If
    
    Set DetectMainTableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, endCol))
End Function

Public Function DetectHeaderRow(ByVal rng As Range) As Long
    Dim r As Long, c As Long
    Dim cellText As String
    Dim foundHeader As Boolean
    
    For r = 1 To Application.Min(15, rng.Rows.Count)
        foundHeader = False
        For c = 1 To Application.Min(10, rng.Columns.Count)
            cellText = Trim(CStr(rng.Cells(r, c).Value))
            
            If InStr(1, cellText, "STT", vbTextCompare) > 0 Or _
               InStr(1, cellText, "Noi dung", vbTextCompare) > 0 Or _
               InStr(1, cellText, "Don vi", vbTextCompare) > 0 Or _
               InStr(1, cellText, "Khoi luong", vbTextCompare) > 0 Or _
               InStr(1, cellText, "Thanh tien", vbTextCompare) > 0 Or _
               (Left(cellText, 1) = "(" And Right(cellText, 1) = ")" And IsNumeric(Mid(cellText, 2, Len(cellText) - 2))) Then
                foundHeader = True
                Exit For
            End If
        Next c
        
        If foundHeader Then
            If r > 1 Then
                Dim prevText As String
                prevText = Trim(CStr(rng.Cells(r - 1, 1).Value))
                If InStr(1, prevText, "STT", vbTextCompare) > 0 Or _
                   InStr(1, prevText, "Noi dung", vbTextCompare) > 0 Or _
                   InStr(1, prevText, "Don vi", vbTextCompare) > 0 Or _
                   InStr(1, prevText, "Khoi luong", vbTextCompare) > 0 Then
                    DetectHeaderRow = r - 1
                    Exit Function
                End If
            End If
            
            DetectHeaderRow = r
            Exit Function
        End If
    Next r
    
    DetectHeaderRow = 10
End Function

Private Function DetectSignatureStart(ByVal tableRange As Range) As Long
    Dim ws As Worksheet, r As Long, c As Long, startRow As Long, lastRow As Long
    Dim textValue As String, terms As Variant, term As Variant
    terms = Array("nguoi lap", "nguoi kiem", "ke toan", "giam doc", "truong phong", "thu truong", "ky ten", "xac nhan")
    Set ws = tableRange.Worksheet
    lastRow = tableRange.Row + tableRange.Rows.Count - 1
    startRow = WorksheetFunction.Max(tableRange.Row + 2, lastRow - 22)
    For r = startRow To lastRow
        For c = tableRange.Column To tableRange.Column + tableRange.Columns.Count - 1
            textValue = LCase$(Trim$(CStr(ws.Cells(r, c).Value2)))
            If Len(textValue) > 0 Then
                For Each term In terms
                    If InStr(1, textValue, CStr(term), vbTextCompare) > 0 Then
                        DetectSignatureStart = r
                        Exit Function
                    End If
                Next term
            End If
        Next c
    Next r
End Function

Private Sub ExpandRowsSafely(ByVal tableRange As Range)
    Dim ws As Worksheet, r As Long, oldHeight As Double
    Set ws = tableRange.Worksheet
    For r = tableRange.Row To tableRange.Row + tableRange.Rows.Count - 1
        If Not ws.Rows(r).Hidden Then
            oldHeight = ws.Rows(r).RowHeight
            If RowHasWrappedText(ws, r, tableRange.Column, tableRange.Columns.Count) Then
                ws.Rows(r).AutoFit
                If ws.Rows(r).RowHeight < oldHeight Then ws.Rows(r).RowHeight = oldHeight
            End If
        End If
    Next r
End Sub

Private Function RowHasWrappedText(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal firstColumn As Long, ByVal columnCount As Long) As Boolean
    Dim c As Long
    For c = firstColumn To firstColumn + columnCount - 1
        If ws.Cells(rowNumber, c).WrapText Then
            RowHasWrappedText = True
            Exit Function
        End If
    Next c
End Function

Private Sub ApplyPageSetup(ByVal ws As Worksheet, ByVal tableRange As Range, ByVal headerRow As Long, ByVal signatureRow As Long, ByRef opt As TPrintOptions)
    Dim ps As PageSetup, footerRight As String, breakRow As Long
    Dim endHeaderRow As Long
    Set ps = ws.PageSetup
    
    On Error Resume Next
    Application.PrintCommunication = False
    
    ps.PrintArea = tableRange.Address
    ps.PaperSize = PaperSizeFromName(opt.PaperSizeName)
    ps.BlackAndWhite = opt.BlackAndWhite
    ps.CenterHorizontally = opt.CenterHorizontally
    ps.PrintGridlines = opt.PrintGridlines
    ps.PrintHeadings = opt.PrintHeadings
    ps.Order = xlDownThenOver
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = False
    ps.Orientation = OrientationFromOption(opt.OrientationMode, tableRange)
    
    If opt.MarginMode = "Standard" Then
        ps.LeftMargin = Application.InchesToPoints(0.5)
        ps.RightMargin = Application.InchesToPoints(0.5)
        ps.TopMargin = Application.InchesToPoints(0.65)
        ps.BottomMargin = Application.InchesToPoints(0.65)
        ps.HeaderMargin = Application.InchesToPoints(0.3)
        ps.FooterMargin = Application.InchesToPoints(0.3)
    ElseIf opt.MarginMode = "Narrow" Then
        ps.LeftMargin = Application.InchesToPoints(0.25)
        ps.RightMargin = Application.InchesToPoints(0.25)
        ps.TopMargin = Application.InchesToPoints(0.35)
        ps.BottomMargin = Application.InchesToPoints(0.35)
    End If
    
    If opt.RepeatHeader And headerRow > 0 Then
        endHeaderRow = headerRow + 1
        If endHeaderRow > tableRange.Row + tableRange.Rows.Count - 1 Then
            endHeaderRow = tableRange.Row + tableRange.Rows.Count - 1
        End If
        ps.PrintTitleRows = "$" & headerRow & ":$" & endHeaderRow
    End If
    
    If opt.FooterEnabled Then
        footerRight = opt.footerRight
        If opt.IncludePrintDate Then footerRight = footerRight & "  &D"
        ps.LeftFooter = opt.FooterLeft
        ps.CenterFooter = opt.FooterCenter
        ps.RightFooter = footerRight
    Else
        ps.LeftFooter = vbNullString: ps.CenterFooter = vbNullString: ps.RightFooter = vbNullString
    End If
    
    Application.PrintCommunication = True
    On Error GoTo 0
    
    If opt.KeepSignatureWithData And signatureRow > 0 And opt.AutoLastPageBreak Then
        If Not (opt.KeepManualBreaks And ws.HPageBreaks.Count > 0) Then
            If Not opt.KeepManualBreaks Then ws.ResetAllPageBreaks
            breakRow = signatureRow - WorksheetFunction.Max(1, opt.SignatureDataRows)
            If breakRow > tableRange.Row + 1 Then
                On Error Resume Next
                ws.HPageBreaks.Add Before:=ws.Rows(breakRow)
                On Error GoTo 0
            End If
        End If
    End If
End Sub

Private Function PaperSizeFromName(ByVal paperName As String) As XlPaperSize
    Select Case UCase$(paperName)
        Case "A3": PaperSizeFromName = xlPaperA3
        Case "A5": PaperSizeFromName = xlPaperA5
        Case "LETTER": PaperSizeFromName = xlPaperLetter
        Case Else: PaperSizeFromName = xlPaperA4
    End Select
End Function

Private Function OrientationFromOption(ByVal mode As String, ByVal tableRange As Range) As XlPageOrientation
    Dim widthEstimate As Double, c As Long
    If UCase$(mode) = "DOC" Then
        OrientationFromOption = xlPortrait
    ElseIf UCase$(mode) = "NGANG" Then
        OrientationFromOption = xlLandscape
    Else
        For c = tableRange.Column To tableRange.Column + tableRange.Columns.Count - 1
            widthEstimate = widthEstimate + tableRange.Worksheet.Columns(c).ColumnWidth
        Next c
        If widthEstimate > 85 Or tableRange.Columns.Count > 8 Then
            OrientationFromOption = xlLandscape
        Else
            OrientationFromOption = xlPortrait
        End If
    End If
End Function

Public Sub SaveDefaultOptions(ByRef opt As TPrintOptions)
    Dim ws As Worksheet, r As Long
    Set ws = GetConfigSheet(True)
    If ws Is Nothing Then Exit Sub
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("Option", "Value")
    r = 2
    PutConfig ws, r, "UseSelection", opt.useSelection
    PutConfig ws, r, "PaperSizeName", opt.PaperSizeName
    PutConfig ws, r, "OrientationMode", opt.OrientationMode
    PutConfig ws, r, "BlackAndWhite", opt.BlackAndWhite
    PutConfig ws, r, "CenterHorizontally", opt.CenterHorizontally
    PutConfig ws, r, "RepeatHeader", opt.RepeatHeader
    PutConfig ws, r, "ExpandRowHeight", opt.ExpandRowHeight
    PutConfig ws, r, "KeepSignatureWithData", opt.KeepSignatureWithData
    PutConfig ws, r, "SignatureDataRows", opt.SignatureDataRows
    PutConfig ws, r, "FooterEnabled", opt.FooterEnabled
    PutConfig ws, r, "FooterLeft", opt.FooterLeft
    PutConfig ws, r, "FooterCenter", opt.FooterCenter
    PutConfig ws, r, "FooterRight", opt.footerRight
    PutConfig ws, r, "IncludePrintDate", opt.IncludePrintDate
    PutConfig ws, r, "PrintGridlines", opt.PrintGridlines
    PutConfig ws, r, "PrintHeadings", opt.PrintHeadings
    PutConfig ws, r, "PrintComments", opt.PrintComments
    PutConfig ws, r, "SkipHiddenSheets", opt.SkipHiddenSheets
    PutConfig ws, r, "PreviewAfterSetup", opt.PreviewAfterSetup
    PutConfig ws, r, "MarginMode", opt.MarginMode
    PutConfig ws, r, "KeepManualBreaks", opt.KeepManualBreaks
    PutConfig ws, r, "AutoLastPageBreak", opt.AutoLastPageBreak
    ws.Visible = xlSheetVeryHidden
    MsgBox "Da luu thiet lap mac dinh cho workbook nay.", vbInformation, "Thiet lap in nhanh"
End Sub

Public Sub LoadDefaultOptions(ByRef opt As TPrintOptions)
    Dim ws As Worksheet
    SetDefaultOptions opt
    Set ws = GetConfigSheet(False)
    If ws Is Nothing Then Exit Sub
    opt.useSelection = GetConfig(ws, "UseSelection", opt.useSelection)
    opt.PaperSizeName = GetConfig(ws, "PaperSizeName", opt.PaperSizeName)
    opt.OrientationMode = GetConfig(ws, "OrientationMode", opt.OrientationMode)
    opt.BlackAndWhite = GetConfig(ws, "BlackAndWhite", opt.BlackAndWhite)
    opt.CenterHorizontally = GetConfig(ws, "CenterHorizontally", opt.CenterHorizontally)
    opt.RepeatHeader = GetConfig(ws, "RepeatHeader", opt.RepeatHeader)
    opt.ExpandRowHeight = GetConfig(ws, "ExpandRowHeight", opt.ExpandRowHeight)
    opt.KeepSignatureWithData = GetConfig(ws, "KeepSignatureWithData", opt.KeepSignatureWithData)
    opt.SignatureDataRows = CLng(GetConfig(ws, "SignatureDataRows", opt.SignatureDataRows))
    opt.FooterEnabled = GetConfig(ws, "FooterEnabled", opt.FooterEnabled)
    opt.FooterLeft = GetConfig(ws, "FooterLeft", opt.FooterLeft)
    opt.FooterCenter = GetConfig(ws, "FooterCenter", opt.FooterCenter)
    opt.footerRight = GetConfig(ws, "FooterRight", opt.footerRight)
    opt.IncludePrintDate = GetConfig(ws, "IncludePrintDate", opt.IncludePrintDate)
    opt.PrintGridlines = GetConfig(ws, "PrintGridlines", opt.PrintGridlines)
    opt.PrintHeadings = GetConfig(ws, "PrintHeadings", opt.PrintHeadings)
    opt.PrintComments = GetConfig(ws, "PrintComments", opt.PrintComments)
    opt.SkipHiddenSheets = GetConfig(ws, "SkipHiddenSheets", opt.SkipHiddenSheets)
    opt.PreviewAfterSetup = GetConfig(ws, "PreviewAfterSetup", opt.PreviewAfterSetup)
    opt.MarginMode = GetConfig(ws, "MarginMode", opt.MarginMode)
    opt.KeepManualBreaks = GetConfig(ws, "KeepManualBreaks", opt.KeepManualBreaks)
    opt.AutoLastPageBreak = GetConfig(ws, "AutoLastPageBreak", opt.AutoLastPageBreak)
End Sub

Private Function GetConfigSheet(ByVal createIfMissing As Boolean) As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function
    
    On Error Resume Next
    Set GetConfigSheet = wb.Worksheets("__PrintFast_Config")
    On Error GoTo 0
    
    If GetConfigSheet Is Nothing And createIfMissing Then
        On Error Resume Next
        Set GetConfigSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        If Not GetConfigSheet Is Nothing Then GetConfigSheet.Name = "__PrintFast_Config"
        On Error GoTo 0
    End If
End Function

Private Function GetBackupSheet(ByVal createIfMissing As Boolean) As Worksheet
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function
    
    On Error Resume Next
    Set GetBackupSheet = wb.Worksheets("__PrintFast_LastState")
    On Error GoTo 0
    
    If GetBackupSheet Is Nothing And createIfMissing Then
        On Error Resume Next
        Set GetBackupSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        If Not GetBackupSheet Is Nothing Then
            GetBackupSheet.Name = "__PrintFast_LastState"
            GetBackupSheet.Range("A1:Q1").Value = Array("Sheet", "PrintArea", "PrintTitleRows", "Orientation", "BlackWhite", "CenterHorizontal", "Zoom", "FitWide", "FitTall", "LeftMargin", "RightMargin", "TopMargin", "BottomMargin", "LeftFooter", "CenterFooter", "RightFooter", "ManualBreakRows")
            GetBackupSheet.Visible = xlSheetVeryHidden
        End If
        On Error GoTo 0
    End If
End Function

Private Sub PutConfig(ByVal ws As Worksheet, ByRef r As Long, ByVal key As String, ByVal value As Variant)
    ws.Cells(r, 1).Value = key
    ws.Cells(r, 2).Value = value
    r = r + 1
End Sub

Private Function GetConfig(ByVal ws As Worksheet, ByVal key As String, ByVal defaultValue As Variant) As Variant
    Dim hit As Range
    Set hit = ws.Columns(1).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)
    If hit Is Nothing Then
        GetConfig = defaultValue
    Else
        GetConfig = hit.Offset(0, 1).Value
    End If
End Function
