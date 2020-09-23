VERSION 5.00
Begin VB.Form frmFormDesign 
   Caption         =   "New Dialog"
   ClientHeight    =   4140
   ClientLeft      =   2745
   ClientTop       =   3135
   ClientWidth     =   6870
   Icon            =   "FormDsgn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6870
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   315
      Index           =   0
      Left            =   3480
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Index           =   0
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   3960
      Picture         =   "FormDsgn.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   3540
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   2040
      TabIndex        =   2
      Top             =   3540
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   0
      Left            =   3480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   300
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.TextBox Edit1 
      Height          =   915
      Index           =   0
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "FormDsgn.frx":0884
      Top             =   2220
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   5880
      X2              =   6720
      Y1              =   3540
      Y2              =   4140
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   4920
      Picture         =   "FormDsgn.frx":088C
      Stretch         =   -1  'True
      Top             =   3540
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   3000
      Top             =   3540
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   3480
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmFormDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Quote = """"

Private Type CtlInfo
    CtlType As String
    CtlName As String
    Text As String
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type
    
Private CtlArray() As CtlInfo
Private idx As Long


Public TextBoxCount As Integer
Public EditCount As Integer
Public CheckBoxCount As Integer
Public LabelCount As Integer
Public OptionCount As Integer

Private Enum ActiveCtl
    Ctl_None = 0
    Ctl_TextBox
    Ctl_CheckBox
    Ctl_Option
    Ctl_Label
End Enum

Private Active_Control As ActiveCtl
Private NewText  As String
Private OldText As String
Private Kill_Ctl As Boolean

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Const NULL_BRUSH = 5
Private Const PS_SOLID = 0
Private Const R2_NOT = 6

Enum ControlState
    StateNothing = 0
    StateDragging
    StateSizing
End Enum

Private m_CurrCtl As Control
Private m_DragState As ControlState
Private m_DragHandle As Integer
Private m_DragRect As New CRect
Private m_DragPoint As POINTAPI

Public DesignMode As Boolean

Private Const VK_ESCAPE = &H1B
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public CtlMover As CControlSizer

Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Function LoadForm(ByVal FormFile As String) As String
ReDim CtlArray(0 To 99)

Dim s As String, i As Long, ff As Integer
ff = FreeFile

Open FormFile For Input As #ff
s = Input(LOF(ff), ff)
Close #ff

RX_Blocks s

ReDim Preserve CtlArray(0 To idx - 1)
idx = 0

For i = 0 To UBound(CtlArray)
    With CtlArray(i)
    'MsgBox "|" & .CtlType & "|" & .CtlName & "|" & .Text & "|" & .Width
    MakeCtl .CtlType, .CtlName, .Text, .Top, .Left, .Width, .Height
    End With
Next

'Show vbModal

'wait until form is closed

'LoadForm = Tag

End Function
Private Function RX_GenericExtractSubMatch(ByVal Text As String, ByVal Pattern As String, Optional ByVal SubMatchIndex As Integer = 0, Optional ByVal IgnoreCase As Boolean = True) As String
Dim SC As CStrCat
Dim m As Match
Dim objRegExp As RegExp

Set SC = New CStrCat
Set objRegExp = New RegExp

objRegExp.IgnoreCase = IgnoreCase
objRegExp.Global = True
objRegExp.Pattern = Pattern
objRegExp.MultiLine = True

SC.MaxLength = Len(Text)

For Each m In objRegExp.Execute(Text)
    SC.AddStr m.SubMatches(SubMatchIndex) & vbCrLf
Next

RX_GenericExtractSubMatch = Left$(SC.StrVal, SC.Length - 2)

Set objRegExp = Nothing
Set SC = Nothing

End Function

Private Function UnQuote(ByVal Text As String) As String
Dim sTemp As String

sTemp = Text

If Left$(Text, 1) = Quote And Right$(Text, 1) = Quote Then
    sTemp = Mid$(sTemp, 2, Len(sTemp) - 2)
End If

UnQuote = sTemp

End Function

Private Sub MakeCtl(ByVal CtlType As String, ByVal CtlName As String, ByVal CtlText As String, _
                   ByVal CtlTop As Single, _
                   ByVal CtlLeft As Single, _
                   ByVal Ctlwidth As Single, _
                   ByVal CtlHeight As Single)

Dim idx As Long
Dim ctlX As Object  '// form or control

'TextBoxCount = 0
'EditCount = 0
'CheckBoxCount = 0
'LabelCount = 0
'OptionCount = 0


CtlType = LCase(CtlType)
Select Case CtlType
    Case "form"
        Caption = CtlText
        Set ctlX = Me
    Case "button"
        'idx = Button.Count
        'Load Button(idx)
        'Button(idx).Caption = CtlText
        'Set ctlX = Button(idx)
    Case "editbox"
        idx = Edit1.Count
        Load Edit1(idx)
        Edit1(idx).Text = CtlText
        Set ctlX = Edit1(idx)
        EditCount = EditCount + 1
    Case "textbox"
        idx = Text1.Count
        Load Text1(idx)
        Text1(idx).Text = CtlText
        Set ctlX = Text1(idx)
        TextBoxCount = TextBoxCount + 1
    Case "label"
        idx = Label1.Count
        Load Label1(idx)
        Label1(idx).Caption = CtlText
        Set ctlX = Label1(idx)
        LabelCount = LabelCount + 1
    Case "optionbtn"
        idx = Option1.Count
        Load Option1(idx)
        Option1(idx).Caption = CtlText
        Set ctlX = Option1(idx)
        OptionCount = OptionCount + 1
     Case "checkbox"
        idx = Check1.Count
        Load Check1(idx)
        Check1(idx).Caption = CtlText
        Set ctlX = Check1(idx)
        CheckBoxCount = CheckBoxCount + 1

    Case Else
        Exit Sub
End Select

ctlX.Left = CtlLeft
ctlX.Top = CtlTop
ctlX.Width = Ctlwidth
ctlX.Height = CtlHeight
ctlX.Tag = CtlName

If TypeOf ctlX Is Form Then
    'form is displayed later
Else
    ctlX.Visible = True
    ctlX.ToolTipText = "Control Name: " & CtlName
End If

End Sub
Private Function RX_EachBlocks(ByVal Text As String) As String

CtlArray(idx).CtlType = RX_GenericExtractSubMatch(Text, "begin\s*(\w+)")

CtlArray(idx).Text = UnQuote(RX_GenericExtractSubMatch(Text, "text\s*\=\s*(.*?)\r"))
CtlArray(idx).Text = Replace(CtlArray(idx).Text, Chr$(7), vbCrLf)

CtlArray(idx).CtlName = UnQuote(RX_GenericExtractSubMatch(Text, "name\s*\=\s*(.*?)\r"))

CtlArray(idx).Width = CSng(RX_GenericExtractSubMatch(Text, "width\s*\=\s*(.*)"))
CtlArray(idx).Height = CSng(RX_GenericExtractSubMatch(Text, "height\s*\=\s*(.*)"))
CtlArray(idx).Left = CSng(RX_GenericExtractSubMatch(Text, "left\s*\=\s*(.*)"))
CtlArray(idx).Top = CSng(RX_GenericExtractSubMatch(Text, "top\s*\=\s*(.*)"))


End Function
Private Function RX_Blocks(ByVal Text As String) As String
Dim sTemp As String
Dim objRegExp As RegExp
Set objRegExp = New RegExp

objRegExp.MultiLine = True
objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "begin[^\v]*?end"

Dim m As Match

For Each m In objRegExp.Execute(Text)
   
   RX_EachBlocks m.Value
   idx = idx + 1
Next

Set objRegExp = Nothing

End Function


Private Sub DoPopMenu(ByRef ctlX As Control)
Dim sOldText As String, sNewText As String, sNewName As String
Dim lResult As Long

If TypeOf ctlX Is TextBox Then
    sOldText = ctlX.Text
Else
    sOldText = ctlX.Caption
End If


lResult = PopUp("Name", "Text", "-", "Delete")
                  '1       2     3      4
Select Case lResult
    
    Case 1  'Name
        sNewName = InputBox("Enter Name for the Control", "Name", ctlX.Tag)
        If sNewName <> "" Then
            ctlX.Tag = sNewName
        End If


    Case 2  'Text or Caption
        sNewText = InputBox("Enter new Text", "New Text", sOldText)
        If sNewText <> "" Then
                If TypeOf ctlX Is TextBox Then
                      ctlX.Text = sNewText
                Else
                      ctlX.Caption = sNewText
                End If
        End If
    
    Case 4  'Delete Control (actually just hide it!)
        ctlX.Visible = False
        CtlMover.HideHandles

End Select

End Sub

Private Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lResult As Long

If Button = vbLeftButton And DesignMode Then
    CtlMover.AttachControl Check1(Index)
End If

If Button = vbRightButton Then

    DoPopMenu Check1(Index)
        
End If

End Sub
Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewCaption  As String

If Button = vbRightButton Then
    
    NewCaption = InputBox("Enter new Caption", "New Caption", Check1(Index).Caption)
    If NewCaption <> "" Then
        Check1(Index).Caption = NewCaption
    End If

End If

End Sub

Private Sub Edit1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And DesignMode Then
    CtlMover.AttachControl Edit1(Index)
End If

If Button = vbRightButton Then

    Edit1(Index).Enabled = False
    Edit1(Index).Enabled = True

    DoPopMenu Edit1(Index)

End If

End Sub

Private Sub Form_Load()

Set CtlMover = New CControlSizer
CtlMover.GridSize = 6
CtlMover.AttachForm Me 'The form that is using the designer class
CtlMover.DrawGrid

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        PopUpMenu frmMain.mnuTools, vbLeftButton
    End If

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
    Cancel = True
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

'ctl.SaveControls
Set CtlMover = Nothing

End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
If Button = vbLeftButton And DesignMode Then
    CtlMover.AttachControl Label1(Index)
End If

If Button = vbRightButton Then

    DoPopMenu Label1(Index)

End If

End Sub
Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewCaption  As String

If Button = vbRightButton Then
    
    NewCaption = InputBox("Enter new Caption", "New Caption", Label1(Index).Caption)
    
    If NewCaption <> "" Then
    Label1(Index).Caption = NewCaption
    End If

End If


End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton And DesignMode Then
        CtlMover.AttachControl List1
    End If

End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And DesignMode Then
        CtlMover.AttachControl Image1
    End If
End Sub

Private Sub Option1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
If Button = vbLeftButton And DesignMode Then
    CtlMover.AttachControl Option1(Index)
End If

If Button = vbRightButton Then

    DoPopMenu Option1(Index)
        
End If

End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And DesignMode Then
        CtlMover.AttachControl Picture1
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton And DesignMode Then
    End If

    CtlMover.HideHandles

    Me.Refresh

End Sub
Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton And DesignMode Then
    'DragBegin Text1(Index)
    CtlMover.AttachControl Text1(Index)
End If

If Button = vbRightButton Then

    Text1(Index).Enabled = False
    Text1(Index).Enabled = True

    DoPopMenu Text1(Index)
        

End If

End Sub
