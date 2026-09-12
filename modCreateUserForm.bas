Attribute VB_Name = "modCreateUserForm"
Option Explicit

'Tao UserForm tu dong
Public Sub CreatePrintFastUserForm()
    Dim VBProj As Object
    Dim VBComp As Object
    Dim CodeModule As Object
    Dim frm As Object
    Dim i As Integer
    
    On Error Resume Next
    Set VBProj = ThisWorkbook.VBProject
    
    'Xoa form cu neu ton tai
    For Each VBComp In VBProj.VBComponents
        If VBComp.Name = "frmExcelPrintFast" Then
            VBProj.VBComponents.Remove VBComp
            Exit For
        End If
    Next VBComp
    
    'Tao UserForm moi
    Set frm = VBProj.VBComponents.Add(3) '3 = vbext_ct_MSForm
    frm.Name = "frmExcelPrintFast"
    frm.Properties("Caption").Value = "Excel Print Fast - Cau Hinh In Nhanh"
    frm.Properties("Width").Value = 600
    frm.Properties("Height").Value = 550
    frm.Properties("StartUpPosition").Value = 1 'CenterOwner
    
    'Add Controls
    AddControlsToForm frm
    
    'Add Code
    Set CodeModule = frm.CodeModule
    AddFormCode CodeModule
    
    MsgBox "UserForm da tao thanh cong!" & vbCrLf & "Chay: ShowExcelPrintFast()", vbInformation
End Sub

Private Sub AddControlsToForm(frm As Object)
    Dim ctrl As Object
    Dim topPos As Integer
    
    topPos = 10
    
    '=== LABEL: Chon Sheet ==="
    Set ctrl = frm.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblSheets"
        .Caption = "Chon Sheet:"
        .Left = 10
        .Top = topPos
        .Width = 80
        .Height = 20
        .Font.Bold = True
    End With
    topPos = topPos + 25
    
    '=== FRAME: Sheet Checklist ==="
    Set ctrl = frm.Controls.Add("Forms.Frame.1")
    With ctrl
        .Name = "fraSheets"
        .Caption = ""
        .Left = 10
        .Top = topPos
        .Width = 280
        .Height = 120
        .BorderStyle = 1
    End With
    topPos = topPos + 130
    
    '=== BUTTONS: Select/Clear ==="
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "cmdSelectAllSheets"
        .Caption = "Chon tat ca"
        .Left = 10
        .Top = topPos
        .Width = 80
        .Height = 20
    End With
    
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "cmdClearSheets"
        .Caption = "Bo chon"
        .Left = 100
        .Top = topPos
        .Width = 80
        .Height = 20
    End With
    topPos = topPos + 30
    
    '=== LABEL: Cau Hinh ==="
    Set ctrl = frm.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblSettings"
        .Caption = "Cau hinh in:"
        .Left = 10
        .Top = topPos
        .Width = 80
        .Height = 20
        .Font.Bold = True
    End With
    topPos = topPos + 25
    
    '=== Paper Size ==="
    Set ctrl = frm.Controls.Add("Forms.Label.1")
    With ctrl
        .Caption = "Co giay:"
        .Left = 10
        .Top = topPos
        .Width = 60
        .Height = 18
    End With
    
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboPaperSize"
        .Left = 80
        .Top = topPos
        .Width = 120
        .Height = 20
        .AddItem "A4"
        .AddItem "A3"
        .AddItem "A5"
        .AddItem "Letter"
        .Value = "A4"
    End With
    topPos = topPos + 25
    
    '=== Orientation ==="
    Set ctrl = frm.Controls.Add("Forms.Label.1")
    With ctrl
        .Caption = "Huong:"
        .Left = 10
        .Top = topPos
        .Width = 60
        .Height = 18
    End With
    
    Set ctrl = frm.Controls.Add("Forms.OptionButton.1")
    With ctrl
        .Name = "optPortrait"
        .Caption = "Doc"
        .Left = 80
        .Top = topPos
        .Width = 50
        .Height = 18
    End With
    
    Set ctrl = frm.Controls.Add("Forms.OptionButton.1")
    With ctrl
        .Name = "optLandscape"
        .Caption = "Ngang"
        .Left = 140
        .Top = topPos
        .Width = 60
        .Height = 18
    End With
    
    Set ctrl = frm.Controls.Add("Forms.OptionButton.1")
    With ctrl
        .Name = "optAutoOrientation"
        .Caption = "Tu dong"
        .Left = 210
        .Top = topPos
        .Width = 70
        .Height = 18
        .Value = True
    End With
    topPos = topPos + 25
    
    '=== Checkboxes ==="
    Set ctrl = frm.Controls.Add("Forms.CheckBox.1")
    With ctrl
        .Name = "chkBlackWhite"
        .Caption = "Den trang"
        .Left = 10
        .Top = topPos
        .Width = 80
        .Height = 18
        .Value = True
    End With
    topPos = topPos + 20
    
    Set ctrl = frm.Controls.Add("Forms.CheckBox.1")
    With ctrl
        .Name = "chkCenterHorizontally"
        .Caption = "Giua ngang"
        .Left = 10
        .Top = topPos
        .Width = 80
        .Height = 18
        .Value = True
    End With
    topPos = topPos + 20
    
    Set ctrl = frm.Controls.Add("Forms.CheckBox.1")
    With ctrl
        .Name = "chkRepeatHeader"
        .Caption = "Lap header"
        .Left = 10
        .Top = topPos
        .Width = 80
        .Height = 18
        .Value = True
    End With
    topPos = topPos + 20
    
    Set ctrl = frm.Controls.Add("Forms.CheckBox.1")
    With ctrl
        .Name = "chkGridlines"
        .Caption = "Luoi"
        .Left = 10
        .Top = topPos
        .Width = 80
        .Height = 18
    End With
    topPos = topPos + 20
    
    Set ctrl = frm.Controls.Add("Forms.CheckBox.1")
    With ctrl
        .Name = "chkPreview"
        .Caption = "Xem truoc"
        .Left = 10
        .Top = topPos
        .Width = 80
        .Height = 18
        .Value = True
    End With
    topPos = topPos + 30
    
    '=== LABEL: Progress ==="
    Set ctrl = frm.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblProgress"
        .Caption = "San sang thiet lap..."
        .Left = 10
        .Top = topPos
        .Width = 280
        .Height = 30
        .WordWrap = True
        .BackColor = RGB(240, 240, 240)
    End With
    topPos = topPos + 35
    
    '=== BUTTONS: Run/Close ==="
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "cmdRun"
        .Caption = "Chay"
        .Left = 10
        .Top = topPos
        .Width = 80
        .Height = 25
        .BackColor = RGB(0, 176, 80)
        .ForeColor = RGB(255, 255, 255)
    End With
    
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "cmdClose"
        .Caption = "Dong"
        .Left = 100
        .Top = topPos
        .Width = 80
        .Height = 25
    End With
    
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "cmdSaveDefaults"
        .Caption = "Luu mac dinh"
        .Left = 190
        .Top = topPos
        .Width = 100
        .Height = 25
    End With
End Sub

Private Sub AddFormCode(CodeModule As Object)
    Dim formCode As String
    
    formCode = "Option Explicit" & vbCrLf & vbCrLf & _
    "Private mOptions As TPrintOptions" & vbCrLf & vbCrLf & _
    "Private Sub UserForm_Initialize()" & vbCrLf & _
    "    LoadDefaultOptions mOptions" & vbCrLf & _
    "    BuildSheetChecklist" & vbCrLf & _
    "    ApplyOptionsToForm" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub BuildSheetChecklist()" & vbCrLf & _
    "    Dim ws As Worksheet, chk As MSForms.CheckBox, topPos As Single" & vbCrLf & _
    "    topPos = 10" & vbCrLf & _
    "    For Each ws In ActiveWorkbook.Worksheets" & vbCrLf & _
    "        If ws.Name <> \"__PrintFast_Config\" And ws.Name <> \"__PrintFast_LastState\" Then" & vbCrLf & _
    "            Set chk = Me.fraSheets.Controls.Add(\"Forms.CheckBox.1\", \"sheet_\" & CStr(ws.Index), True)" & vbCrLf & _
    "            chk.Caption = ws.Name: chk.Tag = ws.Name" & vbCrLf & _
    "            chk.Left = 10: chk.Top = topPos: chk.Width = 260: chk.Height = 18" & vbCrLf & _
    "            chk.Value = (ws.Visible = xlSheetVisible)" & vbCrLf & _
    "            topPos = topPos + 20" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    Next ws" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub ApplyOptionsToForm()" & vbCrLf & _
    "    cboPaperSize.Value = mOptions.PaperSizeName" & vbCrLf & _
    "    optAutoOrientation.Value = (mOptions.OrientationMode = \"Auto\")" & vbCrLf & _
    "    optPortrait.Value = (mOptions.OrientationMode = \"Doc\")" & vbCrLf & _
    "    optLandscape.Value = (mOptions.OrientationMode = \"Ngang\")" & vbCrLf & _
    "    chkBlackWhite.Value = mOptions.BlackAndWhite" & vbCrLf & _
    "    chkCenterHorizontally.Value = mOptions.CenterHorizontally" & vbCrLf & _
    "    chkRepeatHeader.Value = mOptions.RepeatHeader" & vbCrLf & _
    "    chkGridlines.Value = mOptions.PrintGridlines" & vbCrLf & _
    "    chkPreview.Value = mOptions.PreviewAfterSetup" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Function SelectedSheetNames() As Collection" & vbCrLf & _
    "    Dim ctl As control, selected As New Collection" & vbCrLf & _
    "    For Each ctl In fraSheets.Controls" & vbCrLf & _
    "        If TypeName(ctl) = \"CheckBox\" Then" & vbCrLf & _
    "            If ctl.Value = True Then selected.Add ctl.Tag" & vbCrLf & _
    "        End If" & vbCrLf & _
    "    Next ctl" & vbCrLf & _
    "    Set SelectedSheetNames = selected" & vbCrLf & _
    "End Function" & vbCrLf & vbCrLf & _
    "Private Sub cmdSelectAllSheets_Click()" & vbCrLf & _
    "    Dim ctl As control" & vbCrLf & _
    "    For Each ctl In fraSheets.Controls" & vbCrLf & _
    "        If TypeName(ctl) = \"CheckBox\" Then ctl.Value = True" & vbCrLf & _
    "    Next ctl" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub cmdClearSheets_Click()" & vbCrLf & _
    "    Dim ctl As control" & vbCrLf & _
    "    For Each ctl In fraSheets.Controls" & vbCrLf & _
    "        If TypeName(ctl) = \"CheckBox\" Then ctl.Value = False" & vbCrLf & _
    "    Next ctl" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub cmdRun_Click()" & vbCrLf & _
    "    Dim sheets As Collection" & vbCrLf & _
    "    Dim wsName As Variant, ws As Worksheet" & vbCrLf & _
    "    Dim totalSheets As Long, currentCount As Long, percent As Long" & vbCrLf & _
    "    Dim oldCalc As XlCalculation" & vbCrLf & vbCrLf & _
    "    Set sheets = SelectedSheetNames" & vbCrLf & _
    "    If sheets.Count = 0 Then" & vbCrLf & _
    "        MsgBox \"Hay chon it nhat mot sheet.\", vbExclamation: Exit Sub" & vbCrLf & _
    "    End If" & vbCrLf & vbCrLf & _
    "    mOptions.PaperSizeName = cboPaperSize.Value" & vbCrLf & _
    "    mOptions.BlackAndWhite = chkBlackWhite.Value" & vbCrLf & _
    "    mOptions.CenterHorizontally = chkCenterHorizontally.Value" & vbCrLf & _
    "    mOptions.RepeatHeader = chkRepeatHeader.Value" & vbCrLf & _
    "    mOptions.PrintGridlines = chkGridlines.Value" & vbCrLf & _
    "    mOptions.PreviewAfterSetup = chkPreview.Value" & vbCrLf & vbCrLf & _
    "    If optPortrait.Value Then mOptions.OrientationMode = \"Doc\"" & vbCrLf & _
    "    If optLandscape.Value Then mOptions.OrientationMode = \"Ngang\"" & vbCrLf & _
    "    If optAutoOrientation.Value Then mOptions.OrientationMode = \"Auto\"" & vbCrLf & vbCrLf & _
    "    totalSheets = sheets.Count" & vbCrLf & _
    "    currentCount = 0" & vbCrLf & _
    "    oldCalc = Application.Calculation" & vbCrLf & _
    "    Application.ScreenUpdating = False" & vbCrLf & _
    "    Application.Calculation = xlCalculationManual" & vbCrLf & vbCrLf & _
    "    On Error GoTo ErrorHandler" & vbCrLf & _
    "    For Each wsName In sheets" & vbCrLf & _
    "        currentCount = currentCount + 1" & vbCrLf & _
    "        percent = CLng((currentCount / totalSheets) * 100)" & vbCrLf & _
    "        lblProgress.Caption = \"Dang xu ly: \" & wsName & \" (\" & percent & \"%)\"" & vbCrLf & _
    "        DoEvents" & vbCrLf & _
    "        Set ws = Nothing" & vbCrLf & _
    "        On Error Resume Next" & vbCrLf & _
    "        Set ws = ActiveWorkbook.Worksheets(CStr(wsName))" & vbCrLf & _
    "        On Error GoTo ErrorHandler" & vbCrLf & _
    "        If Not ws Is Nothing Then ProcessOneSheetDirect ws, mOptions" & vbCrLf & _
    "    Next wsName" & vbCrLf & vbCrLf & _
    "ErrorHandler:" & vbCrLf & _
    "    Application.Calculation = oldCalc" & vbCrLf & _
    "    Application.ScreenUpdating = True" & vbCrLf & _
    "    If Err.Number <> 0 Then" & vbCrLf & _
    "        MsgBox \"Co loi: \" & Err.Description, vbExclamation" & vbCrLf & _
    "    Else" & vbCrLf & _
    "        lblProgress.Caption = \"Hoan tat!\"" & vbCrLf & _
    "        MsgBox \"Da thiet lap in cho \" & totalSheets & \" sheet!\", vbInformation" & vbCrLf & _
    "    End If" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub cmdClose_Click()" & vbCrLf & _
    "    Unload Me" & vbCrLf & _
    "End Sub" & vbCrLf & vbCrLf & _
    "Private Sub cmdSaveDefaults_Click()" & vbCrLf & _
    "    SaveDefaultOptions mOptions" & vbCrLf & _
    "    MsgBox \"Da luu mac dinh!\", vbInformation" & vbCrLf & _
    "End Sub"
    
    CodeModule.AddFromString formCode
End Sub

'Goi ham nay de tao form
Public Sub ShowExcelPrintFast()
    On Error Resume Next
    frmExcelPrintFast.Show
    If Err.Number <> 0 Then
        CreatePrintFastUserForm
        frmExcelPrintFast.Show
    End If
    On Error GoTo 0
End Sub
