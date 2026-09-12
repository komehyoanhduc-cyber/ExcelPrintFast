VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExcelPrintFast 
   Caption         =   "Excel Print Fast - Cau Hinh In Nhanh"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8265.001
   OleObjectBlob   =   "frmExcelPrintFast.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExcelPrintFast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mOptions As TPrintOptions

Private Function U(ByVal txt As String) As String
    Dim i As Long, res As String: res = "": i = 1
    Do While i <= Len(txt)
        If Mid(txt, i, 2) = "\u" And i + 5 <= Len(txt) Then
            res = res & ChrW(CLng("&H" & Mid(txt, i + 2, 4))): i = i + 6
        Else: res = res & Mid(txt, i, 1): i = i + 1: End If
    Loop
    U = res
End Function

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim currentPrinter As String
    Dim printerList As Variant
    
    LoadDefaultOptions mOptions
    BuildSheetChecklist
    BuildComboLists
    ApplyOptionsToForm
    
    On Error Resume Next
    currentPrinter = Application.ActivePrinter
    cboPrinters.Clear
    printerList = GetInstalledPrinters()
    
    If IsArray(printerList) Then
        For i = LBound(printerList) To UBound(printerList)
            cboPrinters.AddItem printerList(i)
            If printerList(i) = currentPrinter Then
                cboPrinters.ListIndex = cboPrinters.ListCount - 1
            End If
        Next i
    End If
    
    If cboPrinters.ListIndex = -1 And cboPrinters.ListCount > 0 Then
        cboPrinters.ListIndex = 0
    End If
    On Error GoTo 0
    
    If Not lblProgress Is Nothing Then lblProgress.Caption = "San sang thiet lap..."
End Sub

Private Sub BuildSheetChecklist()
    Dim ws As Worksheet, chk As MSForms.CheckBox, topPos As Single, maxTop As Single
    topPos = 10
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> "__PrintFast_Config" And ws.Name <> "__PrintFast_LastState" Then
            Set chk = Me.fraSheets.Controls.Add("Forms.CheckBox.1", "sheet_" & CStr(ws.Index), True)
            chk.Caption = ws.Name: chk.Tag = ws.Name
            chk.Left = 10: chk.Top = topPos: chk.Width = 115: chk.Height = 18
            chk.Value = (ws.Visible = xlSheetVisible)
            topPos = topPos + 20
        End If
    Next ws
    maxTop = topPos + 10
    If maxTop > Me.fraSheets.Height Then Me.fraSheets.ScrollHeight = maxTop
End Sub

Private Sub BuildComboLists()
    cboPaperSize.Clear: cboPaperSize.AddItem "A4": cboPaperSize.AddItem "A3": cboPaperSize.AddItem "A5": cboPaperSize.AddItem "Letter"
    cboMargins.Clear
    cboMargins.AddItem U("Tieu chuan")
    cboMargins.AddItem U("Hep")
    cboMargins.AddItem U("Giu hien tai")
    FillFooterCombo cboFooterLeft: FillFooterCombo cboFooterCenter: FillFooterCombo cboFooterRight
End Sub

Private Sub FillFooterCombo(ByVal cbo As MSForms.ComboBox)
    cbo.Clear
    cbo.AddItem U("Khong hien thi")
    cbo.AddItem U("Ten file")
    cbo.AddItem U("Ten sheet")
    cbo.AddItem U("Trang / Tong trang")
    cbo.AddItem U("Ngay in")
End Sub

Private Sub ApplyOptionsToForm()
    cboPaperSize.Value = mOptions.PaperSizeName
    Select Case mOptions.MarginMode
        Case "Standard": cboMargins.ListIndex = 0
        Case "Narrow": cboMargins.ListIndex = 1
        Case Else: cboMargins.ListIndex = 2
    End Select
    optAutoOrientation.Value = (mOptions.OrientationMode = "Auto")
    optPortrait.Value = (mOptions.OrientationMode = "Doc")
    optLandscape.Value = (mOptions.OrientationMode = "Ngang")
    chkBlackWhite.Value = mOptions.BlackAndWhite
    chkCenterHorizontally.Value = mOptions.CenterHorizontally
    chkRepeatHeader.Value = mOptions.RepeatHeader
    chkExpandRows.Value = mOptions.ExpandRowHeight
    chkKeepSignature.Value = mOptions.KeepSignatureWithData
    chkFooter.Value = mOptions.FooterEnabled
    cboFooterLeft.ListIndex = FooterTextToIndex(mOptions.FooterLeft)
    cboFooterCenter.ListIndex = FooterTextToIndex(mOptions.FooterCenter)
    cboFooterRight.ListIndex = FooterTextToIndex(mOptions.footerRight)
    chkIncludeDate.Value = mOptions.IncludePrintDate
    chkGridlines.Value = mOptions.PrintGridlines
    chkPrintHeadings.Value = mOptions.PrintHeadings
    chkSkipHidden.Value = mOptions.SkipHiddenSheets
    chkPreview.Value = mOptions.PreviewAfterSetup
    chkKeepManualBreaks.Value = mOptions.KeepManualBreaks
    txtSignatureRows.Value = CStr(mOptions.SignatureDataRows)
    chkFooter_Click
End Sub

Private Function FooterIndexToText(ByVal idx As Long) As String
    Select Case idx
        Case 1: FooterIndexToText = "&F"
        Case 2: FooterIndexToText = "&A"
        Case 3: FooterIndexToText = "Trang &P/&N"
        Case 4: FooterIndexToText = "&D"
        Case Else: FooterIndexToText = vbNullString
    End Select
End Function

Private Function FooterTextToIndex(ByVal txt As String) As Long
    Select Case txt
        Case "&F": FooterTextToIndex = 1
        Case "&A": FooterTextToIndex = 2
        Case "Trang &P/&N": FooterTextToIndex = 3
        Case "&D": FooterTextToIndex = 4
        Case Else: FooterTextToIndex = 0
    End Select
End Function

Private Function ReadOptionsFromForm() As Boolean
    If Not IsNumeric(txtSignatureRows.Value) Then
        MsgBox U("So dong du lieu cung khoi ky phai la so."), vbExclamation: Exit Function
    End If
    If CLng(txtSignatureRows.Value) < 1 Or CLng(txtSignatureRows.Value) > 20 Then
        MsgBox U("So dong du lieu cung khoi ky phai tu 1 den 20."), vbExclamation: Exit Function
    End If
    mOptions.PaperSizeName = cboPaperSize.Value
    If optPortrait.Value Then
        mOptions.OrientationMode = "Doc"
    ElseIf optLandscape.Value Then
        mOptions.OrientationMode = "Ngang"
    Else
        mOptions.OrientationMode = "Auto"
    End If
    mOptions.BlackAndWhite = chkBlackWhite.Value
    mOptions.CenterHorizontally = chkCenterHorizontally.Value
    mOptions.RepeatHeader = chkRepeatHeader.Value
    mOptions.ExpandRowHeight = chkExpandRows.Value
    mOptions.KeepSignatureWithData = chkKeepSignature.Value
    mOptions.AutoLastPageBreak = chkKeepSignature.Value
    mOptions.SignatureDataRows = CLng(txtSignatureRows.Value)
    mOptions.FooterEnabled = chkFooter.Value
    mOptions.FooterLeft = FooterIndexToText(cboFooterLeft.ListIndex)
    mOptions.FooterCenter = FooterIndexToText(cboFooterCenter.ListIndex)
    mOptions.footerRight = FooterIndexToText(cboFooterRight.ListIndex)
    mOptions.IncludePrintDate = chkIncludeDate.Value
    mOptions.PrintGridlines = chkGridlines.Value
    mOptions.PrintHeadings = chkPrintHeadings.Value
    mOptions.SkipHiddenSheets = chkSkipHidden.Value
    mOptions.PreviewAfterSetup = chkPreview.Value
    mOptions.KeepManualBreaks = chkKeepManualBreaks.Value
    Select Case cboMargins.ListIndex
        Case 0: mOptions.MarginMode = "Standard"
        Case 1: mOptions.MarginMode = "Narrow"
        Case Else: mOptions.MarginMode = "Keep"
    End Select
    ReadOptionsFromForm = True
End Function

Private Function SelectedSheetNames() As Collection
    Dim ctl As control, selected As New Collection
    For Each ctl In fraSheets.Controls
        If TypeName(ctl) = "CheckBox" Then
            If ctl.Value = True Then selected.Add ctl.Tag
        End If
    Next ctl
    Set SelectedSheetNames = selected
End Function

Private Sub cmdSelectAllSheets_Click()
    Dim ctl As control
    For Each ctl In fraSheets.Controls
        If TypeName(ctl) = "CheckBox" Then ctl.Value = True
    Next ctl
End Sub

Private Sub cmdClearSheets_Click()
    Dim ctl As control
    For Each ctl In fraSheets.Controls
        If TypeName(ctl) = "CheckBox" Then ctl.Value = False
    Next ctl
End Sub

Private Sub cmdSaveDefaults_Click()
    If ReadOptionsFromForm Then SaveDefaultOptions mOptions
End Sub

Private Sub cmdRestore_Click()
    Me.Hide: RestorePreviousPrintSetup: Unload Me
End Sub

Private Sub cmdRun_Click()
    Dim sheets As Collection
    Dim wsName As Variant
    Dim ws As Worksheet
    Dim totalSheets As Long, currentCount As Long
    Dim percent As Long
    Dim oldCalc As XlCalculation
    
    If Not ReadOptionsFromForm Then Exit Sub
    Set sheets = SelectedSheetNames
    If sheets.Count = 0 Then
        MsgBox U("Hay chon it nhat mot sheet."), vbExclamation: Exit Sub
    End If
    
    If cboPrinters.Value <> "" Then
        On Error Resume Next
        Application.ActivePrinter = cboPrinters.Value
        On Error GoTo 0
    End If
    
    totalSheets = sheets.Count
    currentCount = 0
    
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    For Each wsName In sheets
        currentCount = currentCount + 1
        percent = CLng((currentCount / totalSheets) * 100)
        
        If Not lblProgress Is Nothing Then
            lblProgress.Caption = U("Dang xu ly: ") & wsName & " (" & percent & "%) - [" & currentCount & "/" & totalSheets & "]"
        End If
        DoEvents
        
        Set ws = Nothing
        On Error Resume Next
        Set ws = ActiveWorkbook.Worksheets(CStr(wsName))
        On Error GoTo ErrorHandler
        
        If Not ws Is Nothing Then
            ProcessOneSheetDirect ws, mOptions
        End If
    Next wsName

ErrorHandler:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox U("Co loi xay ra: ") & Err.Description, vbExclamation, U("Loi")
    Else
        If Not lblProgress Is Nothing Then lblProgress.Caption = U("Hoan tat xu ly thanh cong!")
        
        If mOptions.PreviewAfterSetup And totalSheets > 0 Then
            On Error Resume Next
            ActiveWorkbook.Worksheets(CStr(sheets(1))).PrintPreview
            On Error GoTo 0
        End If
        
        MsgBox U("Da thiet lap in nhanh thanh cong cho ") & totalSheets & U(" sheet!"), vbInformation, U("Hoan tat")
    End If
    
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub chkFooter_Click()
    cboFooterLeft.Enabled = chkFooter.Value
    cboFooterCenter.Enabled = chkFooter.Value
    cboFooterRight.Enabled = chkFooter.Value
    chkIncludeDate.Enabled = chkFooter.Value
End Sub

Private Function GetInstalledPrinters() As Variant
    Dim objWmi As Object, colPrinters As Object, objPrinter As Object
    Dim printerArr() As String
    Dim count As Long
    
    On Error Resume Next
    Set objWmi = GetObject("winmgmts:\\.\root\cimv2")
    Set colPrinters = objWmi.ExecQuery("Select * from Win32_Printer")
    
    If Not colPrinters Is Nothing Then
        ReDim printerArr(0 To colPrinters.Count - 1)
        count = 0
        For Each objPrinter In colPrinters
            printerArr(count) = GetFullPrinterName(objPrinter.Name)
            count = count + 1
        Next objPrinter
        GetInstalledPrinters = printerArr
    Else
        GetInstalledPrinters = Array(Application.ActivePrinter)
    End If
End Function

Private Function GetFullPrinterName(ByVal printerName As String) As String
    Dim subKey As String
    Dim wsh As Object
    Dim fullPrinterName As String
    Dim i As Long
    
    Set wsh = CreateObject("WScript.Shell")
    For i = 0 To 99
        On Error Resume Next
        subKey = "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Devices\" & printerName
        fullPrinterName = wsh.RegRead(subKey)
        On Error GoTo 0
        
        If fullPrinterName <> "" Then
            Dim parts() As String
            parts = Split(fullPrinterName, ",")
            If UBound(parts) >= 1 Then
                GetFullPrinterName = printerName & " on " & parts(1)
                Exit Function
            End If
        End If
    Next i
    GetFullPrinterName = printerName
End Function
