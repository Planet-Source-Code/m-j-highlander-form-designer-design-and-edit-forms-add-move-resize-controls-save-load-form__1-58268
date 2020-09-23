VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CControlSizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=======================================
'Downloaded from Visual Basic Thunder
'www.vbthunder.com
'Created on: 9/12/2002 Last modified 9/15/2002
'=======================================
' - Fixed resizing error when the size rect's right edge
'   or bottom edge was less than the left or top
' - Made GridSize and HandleSize properties, and fixed
'   the problem with controls not snapping correctly
'   to the form grid if the control placement in VB's
'   design mode used a different grid size than the
'   CControlSizer grid size.
' - Set handles' TabStop to False
' - Added an error trap for controls that have one or
'   more read-only size-related properties.
'=======================================

Public GridSize As Long
Public HandleSize As Long

'Eight pictureboxes to use as resize handles.
Dim WithEvents pTL As PictureBox 'top left
Attribute pTL.VB_VarHelpID = -1
Dim WithEvents pT As PictureBox 'top
Attribute pT.VB_VarHelpID = -1
Dim WithEvents pTR As PictureBox 'top right
Attribute pTR.VB_VarHelpID = -1
Dim WithEvents pL As PictureBox 'left
Attribute pL.VB_VarHelpID = -1
Dim WithEvents pR As PictureBox 'right
Attribute pR.VB_VarHelpID = -1
Dim WithEvents pBL As PictureBox 'bottom left
Attribute pBL.VB_VarHelpID = -1
Dim WithEvents pB As PictureBox 'bottom
Attribute pB.VB_VarHelpID = -1
Dim WithEvents pBR As PictureBox 'bottom right
Attribute pBR.VB_VarHelpID = -1
Dim WithEvents pM As PictureBox 'mover
Attribute pM.VB_VarHelpID = -1

'The form and control clients that are utilizing this class
Dim CurFrm As VB.Form
Dim CurCtl As VB.Control

'Flags for moving/sizing
Dim bSizing As Boolean
Dim lStartX As Long
Dim lStartY As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim rc As RECT
Dim sizerc As RECT
Dim bErase As Boolean

Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Sub DrawGrid()

'Draws black pixels, GridSize pixels apart on both the X and Y axis.
'This is done in the Form_Load event BEFORE the form is shown -
'if we waited until the form is visible to do this, it would
'take too much time.
'SetPixelV is used because it's faster than a PSet method.

Dim X As Long
Dim Y As Long

CurFrm.AutoRedraw = True

'Draw on the form's memory image
For Y = 0 To (Screen.Height \ Screen.TwipsPerPixelY) Step GridSize
    For X = 0 To (Screen.Width \ Screen.TwipsPerPixelX) Step GridSize
        SetPixelV CurFrm.hDC, X, Y, vbBlack
    Next
Next
'Set the form's Picture property to what we've drawn -
'a grid that's the size of the screen.
Set CurFrm.Picture = CurFrm.Image

CurFrm.AutoRedraw = False

End Sub
Public Sub AttachForm(frm As VB.Form)

On Error Resume Next
'the On Error is for the line that checks for "pTL"
' - if it errors, it resumes immediately and adds
'the controls as it should. If the control already
'exists, there will be no error and it won't attempt
'to add new pictureboxes.

Set CurFrm = frm

If frm.Controls("pTL") Is Nothing Then
    'Add the handles and size them to HandleSize by HandleSize pixels
    Set pTL = frm.Controls.Add("VB.PictureBox", "pTL")
    pTL.MousePointer = vbSizeNWSE
    SizeHandle pTL
    Set pT = frm.Controls.Add("VB.PictureBox", "pT")
    pT.MousePointer = vbSizeNS
    SizeHandle pT
    Set pTR = frm.Controls.Add("VB.PictureBox", "pTR")
    pTR.MousePointer = vbSizeNESW
    SizeHandle pTR

    Set pL = frm.Controls.Add("VB.PictureBox", "pL")
    pL.MousePointer = vbSizeWE
    SizeHandle pL
    Set pR = frm.Controls.Add("VB.PictureBox", "pR")
    pR.MousePointer = vbSizeWE
    SizeHandle pR

    Set pBL = frm.Controls.Add("VB.PictureBox", "pBL")
    pBL.MousePointer = vbSizeNESW
    SizeHandle pBL
    Set pB = frm.Controls.Add("VB.PictureBox", "pB")
    pB.MousePointer = vbSizeNS
    SizeHandle pB
    Set pBR = frm.Controls.Add("VB.PictureBox", "pBR")
    pBR.MousePointer = vbSizeNWSE
    SizeHandle pBR

    'Add the mover handle
    Set pM = frm.Controls.Add("VB.PictureBox", "pM")
    pM.MousePointer = vbSizeAll
    SizeHandle pM
    pM.Width = CurFrm.ScaleX(HandleSize * 2, vbPixels, CurFrm.ScaleMode)
    pM.Height = CurFrm.ScaleY(HandleSize * 2, vbPixels, CurFrm.ScaleMode)

End If

End Sub
Public Sub AttachControl(ctl As VB.Control)

Dim lOfs As Long
lOfs = CurFrm.ScaleX(3, vbPixels, CurFrm.ScaleMode)

'Save the current control
Set CurCtl = ctl

'Put all the handles in place around the control
'top left
pTL.Move CurCtl.Left - pTL.Width - lOfs, CurCtl.Top - pTL.Height - lOfs
'top
pT.Move CurCtl.Left + ((CurCtl.Width - pT.Width) \ 2), _
    CurCtl.Top - pT.Height - lOfs
'top right
pTR.Move CurCtl.Left + CurCtl.Width + lOfs, _
    CurCtl.Top - pTR.Height - lOfs

'left
pL.Move CurCtl.Left - pL.Width - lOfs, _
    CurCtl.Top + ((CurCtl.Height - pL.Height) \ 2)
'right
pR.Move CurCtl.Left + CurCtl.Width + lOfs, _
    CurCtl.Top + ((CurCtl.Height - pL.Height) \ 2)

'bottom left
pBL.Move CurCtl.Left - pTL.Width - lOfs, _
    CurCtl.Top + CurCtl.Height + lOfs
'bottom
pB.Move CurCtl.Left + ((CurCtl.Width - pT.Width) \ 2), _
    CurCtl.Top + CurCtl.Height + lOfs
'bottom right
pBR.Move CurCtl.Left + CurCtl.Width + lOfs, _
    CurCtl.Top + CurCtl.Height + lOfs

'mover
pM.Move CurCtl.Left + ((CurCtl.Width - pM.Width) \ 2), _
    CurCtl.Top + ((CurCtl.Height - pM.Height) \ 2)

'Show them!
pTL.Visible = True
pT.Visible = True
pTR.Visible = True

pL.Visible = True
pR.Visible = True

pBL.Visible = True
pB.Visible = True
pBR.Visible = True

pM.ZOrder 0
pM.Visible = True

End Sub

Public Sub HideHandles()

pTL.Visible = False
pT.Visible = False
pTR.Visible = False

pL.Visible = False
pR.Visible = False

pBL.Visible = False
pB.Visible = False
pBR.Visible = False

pM.Visible = False

End Sub
Private Sub SizeHandle(pic As PictureBox)

'Initializes the basic properties that are
'common to all the drag handles

'Set the Tag - this class checks the Tag
'during control loading/saving, so that it
'doesn't inadvertently save the position of
'the drag handles.
pic.TabStop = False
'pic.Tag = "VBTDragHandle"
pic.BackColor = vbHighlight
pic.BorderStyle = 0
pic.ScaleMode = vbPixels
pic.Width = CurFrm.ScaleX(HandleSize, vbPixels, CurFrm.ScaleMode)
pic.Height = CurFrm.ScaleY(HandleSize, vbPixels, CurFrm.ScaleMode)
'Bring to the front
pic.ZOrder 0

End Sub
Private Sub DrawHandle(pic As PictureBox)

'Draws a white border along the handle
pic.Line (0, 0)-(pic.ScaleWidth - 1, pic.ScaleHeight - 1), vbHighlightText, B

End Sub

Private Sub Class_Initialize()

GridSize = 8
HandleSize = 8

End Sub

Private Sub pB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

DrawSizeRect True

gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    vbPixels, CurFrm.ScaleMode)) \ gy) * gy

DrawSizeRect

End Sub

Private Sub pB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

EndSize

End Sub

Private Sub pB_Paint()

DrawHandle pB

End Sub

Private Sub pBL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pBL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

DrawSizeRect True

gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - X), _
    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    vbPixels, CurFrm.ScaleMode)) \ gy) * gy

DrawSizeRect

End Sub

Private Sub pBL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

EndSize

End Sub

Private Sub pBL_Paint()

DrawHandle pBL

End Sub

Private Sub pBR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pBR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

DrawSizeRect True

gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
    vbPixels, CurFrm.ScaleMode)) \ gy) * gy

DrawSizeRect

End Sub

Private Sub pBR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

EndSize

End Sub

Private Sub pBR_Paint()

DrawHandle pBR

End Sub

Private Sub pL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim gx As Single

If bSizing = False Then
    Exit Sub
End If

DrawSizeRect True

gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - X), _
    vbPixels, CurFrm.ScaleMode)) \ gx) * gx

DrawSizeRect

End Sub

Private Sub pL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

EndSize

End Sub

Private Sub pL_Paint()

DrawHandle pL

End Sub

Private Sub pM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pM_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

DrawSizeRect True

gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - X), _
    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - Y), _
    vbPixels, CurFrm.ScaleMode)) \ gy) * gy

sizerc.Right = sizerc.Left + CurCtl.Width
sizerc.Bottom = sizerc.Top + CurCtl.Height

'Commented out: By using the following lines instead of the
'above uncommented lines, the Bottom and Right will also
'snap to the grid during the move, which could cause some
'resizing if the grid size changes at all.
'sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
'    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
'sizerc.Bottom = ((rc.Bottom - CurFrm.ScaleY((lStartY - Y), _
'    vbPixels, CurFrm.ScaleMode)) \ gy) * gy

DrawSizeRect

End Sub

Private Sub pM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

EndSize

End Sub

Private Sub pM_Paint()

DrawHandle pM

End Sub

Private Sub pR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim gx As Single

If bSizing = False Then
    Exit Sub
End If

DrawSizeRect True

gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    vbPixels, CurFrm.ScaleMode)) \ gx) * gx

DrawSizeRect

End Sub

Private Sub pR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

EndSize

End Sub

Private Sub pR_Paint()

DrawHandle pR

End Sub

Private Sub pT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

DrawSizeRect True

gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - Y), _
    vbPixels, CurFrm.ScaleMode)) \ gy) * gy

DrawSizeRect

End Sub

Private Sub pT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

EndSize

End Sub

Private Sub pT_Paint()

DrawHandle pT

End Sub

Private Sub pTL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pTL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

DrawSizeRect True

gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
sizerc.Left = ((rc.Left - CurFrm.ScaleX((lStartX - X), _
    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - Y), _
    vbPixels, CurFrm.ScaleMode)) \ gy) * gy

DrawSizeRect

End Sub

Private Sub pTL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

EndSize

End Sub

Private Sub pTL_Paint()

DrawHandle pTL

End Sub

Private Sub pTR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BeginSize X, Y

End Sub

Private Sub pTR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim gx As Single
Dim gy As Single

If bSizing = False Then
    Exit Sub
End If

DrawSizeRect True

gx = CurFrm.ScaleX(GridSize, vbPixels, CurFrm.ScaleMode)
gy = CurFrm.ScaleY(GridSize, vbPixels, CurFrm.ScaleMode)
sizerc.Right = ((rc.Right - CurFrm.ScaleX((lStartX - X), _
    vbPixels, CurFrm.ScaleMode)) \ gx) * gx
sizerc.Top = ((rc.Top - CurFrm.ScaleY((lStartY - Y), _
    vbPixels, CurFrm.ScaleMode)) \ gy) * gy

DrawSizeRect

End Sub

Private Sub pTR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

EndSize

End Sub

Private Sub pTR_Paint()

DrawHandle pTR

End Sub


Private Sub DrawSizeRect(Optional ByVal bEraseOnly As Boolean = False)

'Draw the sizing rectangle during a size or move operation

Dim tW As Long
Dim tS As DrawModeConstants

tW = CurFrm.DrawWidth
tS = CurFrm.DrawMode
CurFrm.DrawWidth = 2
CurFrm.DrawMode = vbInvert

If bErase = True Then
    If bSizing = False Then
        CurFrm.Line (rc.Left, rc.Top)-(rc.Right, rc.Bottom), vbRed, B
    Else
        CurFrm.Line (sizerc.Left, sizerc.Top)-(sizerc.Right, sizerc.Bottom), vbRed, B
    End If
End If
If bEraseOnly = True Then
    bErase = False
    GoTo Cleanup
End If

rc.Left = CurCtl.Left
rc.Top = CurCtl.Top
rc.Right = rc.Left + CurCtl.Width
rc.Bottom = rc.Top + CurCtl.Height

If bSizing = False Then
    CurFrm.Line (rc.Left, rc.Top)-(rc.Right, rc.Bottom), vbRed, B
Else
    CurFrm.Line (sizerc.Left, sizerc.Top)-(sizerc.Right, sizerc.Bottom), vbRed, B
End If

bErase = True

Cleanup:
CurFrm.DrawWidth = tW
CurFrm.DrawMode = tS

End Sub


Private Sub SizeControl()

Dim lTemp As Long

'Check for the rect being inverted and fix it
If sizerc.Bottom < sizerc.Top Then
    lTemp = sizerc.Top
    sizerc.Top = sizerc.Bottom
    sizerc.Bottom = lTemp
End If
If sizerc.Right < sizerc.Left Then
    lTemp = sizerc.Left
    sizerc.Left = sizerc.Right
    sizerc.Right = lTemp
End If
'Make sure we aren't trying to resize the control
'to 0 width or height
If sizerc.Right - sizerc.Left < 1 Then
    sizerc.Right = sizerc.Left + 1
End If
If sizerc.Bottom - sizerc.Top < 1 Then
    sizerc.Bottom = sizerc.Top + 1
End If

On Error Resume Next
'Error trap catches "Height property is read-only"
'for certain controls.

'See the EndSize subroutine
CurCtl.Left = sizerc.Left
CurCtl.Top = sizerc.Top
CurCtl.Width = sizerc.Right - sizerc.Left
CurCtl.Height = sizerc.Bottom - sizerc.Top

End Sub

Private Sub EndSize()

'Halt the sizing operation - called from all drag handles.
'Clear the resizing rectangle
DrawSizeRect True
'Clear the sizing flag
bSizing = False
'Resize the control to the size that was requested
SizeControl
'Reposition the drag handles
AttachControl CurCtl

End Sub
Private Sub BeginSize(ByVal X As Long, ByVal Y As Long)

'Begin the sizing operation - called from all drag handles.
HideHandles
DrawSizeRect
bSizing = True
lStartX = X
lStartY = Y
sizerc = rc

End Sub
