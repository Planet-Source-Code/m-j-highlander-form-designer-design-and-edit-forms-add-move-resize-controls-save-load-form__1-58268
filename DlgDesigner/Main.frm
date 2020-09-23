VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dialog Designer"
   ClientHeight    =   570
   ClientLeft      =   1635
   ClientTop       =   1935
   ClientWidth     =   9150
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   9150
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   926
      ButtonWidth     =   847
      ButtonHeight    =   820
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "label"
            Object.ToolTipText     =   "Label"
            Object.Tag             =   ""
            ImageKey        =   "label"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "text"
            Object.ToolTipText     =   "Text Box"
            Object.Tag             =   ""
            ImageKey        =   "text"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "edit"
            Object.ToolTipText     =   "Edit Box"
            Object.Tag             =   ""
            ImageKey        =   "edit"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "check"
            Object.ToolTipText     =   "Check Box"
            Object.Tag             =   ""
            ImageKey        =   "check"
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "option"
            Object.ToolTipText     =   "Option Button"
            Object.Tag             =   ""
            ImageKey        =   "option"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   1560
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   14610415
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":030A
            Key             =   "check"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":052C
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":074E
            Key             =   "label"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0970
            Key             =   "option"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main.frx":0B92
            Key             =   "text"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu z1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Add"
      Begin VB.Menu mnuEditBox 
         Caption         =   "Edit Box"
      End
      Begin VB.Menu mnuTextBox 
         Caption         =   "Text Box"
      End
      Begin VB.Menu mnuCheckBox 
         Caption         =   "Check Box"
      End
      Begin VB.Menu mnuLabel 
         Caption         =   "Label"
      End
      Begin VB.Menu mnuOptionButton 
         Caption         =   "Option Button"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCaption 
         Caption         =   "Set Form Caption"
      End
      Begin VB.Menu zi8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModeDesign 
         Caption         =   "&Design Mode"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuAlign 
         Caption         =   "Align"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Const WM_USER = &H400
Private Const TB_GETSTYLE = WM_USER + 57
Private Const TB_SETSTYLE = WM_USER + 56
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const TBSTYLE_FLAT = &H800

Private Sub NewDlg()
Dim idx As Long

frmFormDesign.Show


frmFormDesign.TextBoxCount = 0
frmFormDesign.EditCount = 0
frmFormDesign.CheckBoxCount = 0
frmFormDesign.LabelCount = 0
frmFormDesign.OptionCount = 0

frmFormDesign.Caption = "New Dialog"

For idx = 1 To frmFormDesign.Text1.Count - 1
    Unload frmFormDesign.Text1(idx)
Next

For idx = 1 To frmFormDesign.Edit1.Count - 1
    Unload frmFormDesign.Edit1(idx)
Next

For idx = 1 To frmFormDesign.Check1.Count - 1
    Unload frmFormDesign.Check1(idx)
Next

For idx = 1 To frmFormDesign.Option1.Count - 1
    Unload frmFormDesign.Option1(idx)
Next

For idx = 1 To frmFormDesign.Label1.Count - 1
    Unload frmFormDesign.Label1(idx)
Next

End Sub

Sub ToolFlat(ControlName As Control, flat As Boolean)
    Dim style As Long
    Dim hToolbar As Long
    Dim r As Long
       
'Now Make it Flat
    'First get the hWnd
    hToolbar = FindWindowEx(ControlName.hWnd, 0&, "ToolbarWindow32", vbNullString)
    'get Style
    style = SendMessageLong(hToolbar, TB_GETSTYLE, 0&, 0&)
    'Change style
    If (style And TBSTYLE_FLAT) And Not flat Then
        style = style Xor TBSTYLE_FLAT
    ElseIf flat Then
        style = style Or TBSTYLE_FLAT
    End If
    'Set the Style
    r = SendMessageLong(hToolbar, TB_SETSTYLE, 0, style)
    'Now show what we've done, this isn't neccesary if used in form_load
    ControlName.Refresh
End Sub


Private Function CtlType(CtlName As String) As String

Select Case CtlName
    
    Case "Text1": CtlType = "TextBox"
    
    Case "Edit1": CtlType = "EditBox"
    
    Case "Label1": CtlType = "Label"
    
    Case "Check1": CtlType = "CheckBox"
    
    Case "Option1": CtlType = "OptionBtn"

    Case Else

End Select

End Function
Private Function SaveAs(ByVal Text As String) As Boolean   'Form-Specific Function, DO NOT MOVE OR REMOVE
Dim ff As Integer

On Error GoTo Error_SaveAs

cdlg.FileName = ""
cdlg.Filter = "dlg|*.dlg|All Files (*.*)|*.*"
cdlg.DialogTitle = "Save As"
cdlg.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
cdlg.CancelError = False
cdlg.ShowSave

If cdlg.FileName <> "" Then
    ff = FreeFile
    Open cdlg.FileName For Output As #ff
    Print #ff, Text
    Close #ff
    SaveAs = True
Else
    SaveAs = False
End If

Exit Function
Error_SaveAs:
    MsgBox Err.Description, vbCritical, "Save As Error"
    Err.Clear
    cdlg.FileName = ""
    SaveAs = False

End Function

Private Sub Form_Load()

ToolFlat Toolbar1, True
Toolbar1.Refresh


mnuModeDesign_Click
mnuMode_Click

DoEvents

    
If Command$ <> "" Then
    frmFormDesign.LoadForm Replace(Command$, """", "")
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim r As VbMsgBoxResult

r = MsgBox("Are you sure you want to exit", vbQuestion + vbYesNo, "Exit")

If r = vbNo Then
    Cancel = True
End If

End Sub
Private Sub mnuAbout_Click()
Dim s As String
s = "Dialog Designer for Axiom Script" & vbCrLf & vbCrLf
s = s & "Based on codes by:" & vbCrLf
s = s & "Doug Marquardt" & vbCrLf
s = s & "Joseph Guadagno <jguadagno@geocities.com> -  webpage: <http://www.geocities.com/jguadagno>" & vbCrLf
s = s & "Portions Copyright (c) 1997 SoftCircuits Programming (R)" & vbCrLf
s = s & "SoftCircuits Programming - homepage: <http://www.softcircuits.com>" & vbCrLf

MsgBox s, vbInformation, "About"

End Sub
Private Sub mnuAlign_Click()

Dim X As Control
Dim s$



For Each X In frmFormDesign.Controls
    
    If TypeOf X Is Menu Then
    
    Else
       If X.Visible Then
            X.Left = 100 * Round(X.Left / 100, 0)   'convert last 2 numbers to 0
            X.Top = 100 * Round(X.Top / 100, 0)
       End If
    End If
Next

End Sub

Private Sub mnuCaption_Click()
Dim s As String

s = InputBox("Enter Form Caption:", "New Caption", frmFormDesign.Caption)

If s <> "" Then frmFormDesign.Caption = s

End Sub

Private Sub mnuCheckBox_Click()

With frmFormDesign
    .CheckBoxCount = .CheckBoxCount + 1
    
    Load .Check1(.CheckBoxCount)
    .Check1(.CheckBoxCount).Left = (.ScaleWidth - .Check1(.CheckBoxCount).Width) \ 2
    .Check1(.CheckBoxCount).Top = (.ScaleHeight - .Check1(.CheckBoxCount).Height) \ 2
    .Check1(.CheckBoxCount).Caption = "Check " & CStr(.CheckBoxCount)
    .Check1(.CheckBoxCount).Visible = True
    .Check1(.CheckBoxCount).Tag = "Check_" & CStr(.CheckBoxCount)
    .Check1(.CheckBoxCount).ZOrder
    
End With

End Sub

Private Sub mnuEditBox_Click()

With frmFormDesign
    .EditCount = .EditCount + 1
    
    Load .Edit1(.EditCount)
    
    .Edit1(.EditCount).Left = (.ScaleWidth - .Edit1(.EditCount).Width) \ 2
    .Edit1(.EditCount).Top = (.ScaleHeight - .Edit1(.EditCount).Height) \ 2
    .Edit1(.EditCount).Text = "Edit " & CStr(.EditCount)
    .Edit1(.EditCount).Visible = True
    .Edit1(.EditCount).Tag = "Edit_" & CStr(.EditCount)
    .Edit1(.EditCount).ZOrder

End With

End Sub

Private Sub mnuExit_Click()

Unload Me

End Sub
Private Sub mnuLabel_Click()

With frmFormDesign
    
    .LabelCount = .LabelCount + 1
    
    Load .Label1(.LabelCount)
    .Label1(.LabelCount).Left = (.ScaleWidth - .Label1(.LabelCount).Width) \ 2
    .Label1(.LabelCount).Top = (.ScaleHeight - .Label1(.LabelCount).Height) \ 2
    .Label1(.LabelCount).Caption = "Label " & CStr(.LabelCount)
    .Label1(.LabelCount).Visible = True
    .Label1(.LabelCount).Tag = "Label_" & CStr(.LabelCount)
    .Label1(.LabelCount).ZOrder

End With

End Sub
Private Sub mnuMode_Click()

End Sub

Private Sub mnuModeDesign_Click()

    frmFormDesign.DesignMode = Not frmFormDesign.DesignMode
    mnuModeDesign.Checked = Not mnuModeDesign.Checked

    If Not frmFormDesign.DesignMode Then
        'frmFormDesign.DragEnd
    End If

End Sub
Private Sub mnuNew_Click()
Dim idx As Long
Dim r As VbMsgBoxResult

r = MsgBox("Clear current dialog?", vbQuestion + vbYesNo, "New")

If r = vbNo Then
    Exit Sub
End If

NewDlg

End Sub
Private Sub mnuOpen_Click()
On Error GoTo Error_Open

Dim r As VbMsgBoxResult

r = MsgBox("Clear current dialog?", vbQuestion + vbYesNo, "Open")

If r = vbNo Then
    Exit Sub
End If

cdlg.FileName = ""
cdlg.Filter = "dlg|*.dlg|All Files (*.*)|*.*"
cdlg.DialogTitle = "Save As"
cdlg.Flags = cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
cdlg.CancelError = False
cdlg.ShowOpen

If cdlg.FileName <> "" Then
    'ff = FreeFile
    'Open frmHidden.cdlg.FileName For Input As #ff
    'sTemp = Input$(LOF(ff), ff)
    'Close #ff
    NewDlg
    frmFormDesign.LoadForm cdlg.FileName
End If

Exit Sub
Error_Open:
    MsgBox Err.Description, vbCritical, "Open Error"
    Err.Clear
    cdlg.FileName = ""

End Sub

Private Sub mnuOptionButton_Click()



With frmFormDesign

    .OptionCount = .OptionCount + 1
    
    Load .Option1(.OptionCount)
    .Option1(.OptionCount).Left = (.ScaleWidth - .Option1(.OptionCount).Width) \ 2
    .Option1(.OptionCount).Top = (.ScaleHeight - .Option1(.OptionCount).Height) \ 2
    .Option1(.OptionCount).Caption = "Option " & CStr(.OptionCount)
    .Option1(.OptionCount).Tag = "Option_" & CStr(.OptionCount)
    .Option1(.OptionCount).Visible = True
    .Option1(.OptionCount).ZOrder

End With

End Sub
Private Sub mnuSave_Click()
Dim X As Control
Dim s$, sForm As String
Dim sName As String, sText As String, sType As String
Dim sHeight As String, sWidth As String
Dim sxTop As String, sLeft As String

's = s & " Text= " & Chr(34) & X.Text & Chr(34) & vbCrLf
's = s & "Name= " & X.Tag & " Left=" & X.Left & " Top= " & X.Top

frmFormDesign.CtlMover.HideHandles

For Each X In frmFormDesign.Controls
    
    If TypeOf X Is Menu Then
    
    Else
       If X.Visible Then

                sName = """" & X.Tag & """"
                sLeft = CStr(X.Left)
                sxTop = CStr(X.Top)
                sHeight = CStr(X.Height)
                sWidth = CStr(X.Width)
                If TypeOf X Is TextBox Then sText = X.Text Else sText = X.Caption
                sText = """" & Replace(sText, vbCrLf, Chr$(7)) & """"
                sType = CtlType(X.Name)
                s = s & "Begin " & sType & vbCrLf & _
                    "   Name = " & sName & vbCrLf & _
                    "   Text = " & sText & vbCrLf & _
                    "   Height = " & sHeight & vbCrLf & _
                    "   Left = " & sLeft & vbCrLf & _
                    "   Width = " & sWidth & vbCrLf & _
                    "   Top = " & sxTop & vbCrLf & "End" & vbCrLf
        End If
        
    End If

Next

sForm = "Begin Form" & vbCrLf & _
    "   Name = ""AxForm""" & vbCrLf & _
    "   Text = " & """" & frmFormDesign.Caption & """" & vbCrLf & _
    "   Height = " & CStr(frmFormDesign.Height) & vbCrLf & _
    "   Left = " & CStr(frmFormDesign.Left) & vbCrLf & _
    "   Width = " & CStr(frmFormDesign.Width) & vbCrLf & _
    "   Top = " & CStr(frmFormDesign.Top) & vbCrLf & "End" & vbCrLf


s = "AxForm v1.0" & vbCrLf & sForm & s

SaveAs s

End Sub
Private Sub mnuTextBox_Click()

With frmFormDesign

        .TextBoxCount = .TextBoxCount + 1
        
        Load .Text1(.TextBoxCount)
        .Text1(.TextBoxCount).Left = (.ScaleWidth - .Text1(.TextBoxCount).Width) \ 2
        .Text1(.TextBoxCount).Top = (.ScaleHeight - .Text1(.TextBoxCount).Height) \ 2
        .Text1(.TextBoxCount).Text = "Text " & CStr(.TextBoxCount)
        .Text1(.TextBoxCount).Visible = True
        .Text1(.TextBoxCount).Tag = "Text_" & CStr(.TextBoxCount)
        .Text1(.TextBoxCount).ZOrder
End With


End Sub

Private Sub mnuTools_Click()
    
    mnuModeDesign.Checked = frmFormDesign.DesignMode
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

Select Case Button.Key
    
    Case "label"
        mnuLabel_Click
    
    Case "text"
        mnuTextBox_Click

    Case "edit"
        mnuEditBox_Click

    Case "check"
        mnuCheckBox_Click

    Case "option"
        mnuOptionButton_Click

    Case Else

End Select


End Sub
