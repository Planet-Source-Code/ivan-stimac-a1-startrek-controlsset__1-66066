VERSION 5.00
Begin VB.UserControl StarTrekCornerButtonA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "StarTrekCornerButtonA.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   270
      Top             =   540
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   540
   End
End
Attribute VB_Name = "StarTrekCornerButtonA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----- about -----------------------------------------------
'   Star Trek: ControlSet
'   author: ivan štimac
'           ivan.stimac@po.htnet.hr, flashboy01@gmail.com
'
'   price: free for any use
'------------------------------------------------------------

Private HG, WD As Long
Private BC, TC, BCH, BCD, FCD, FC, FCH As OLE_COLOR
Private CAP As String
Private FON As StdFont

Dim i, LFTLNW, TOPLNW As Integer

Private Enum eState
    Normal
    Hover
    Down
End Enum

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private bHovering, ENBL, MPRESS As Boolean
Public Event Click()

'LineWidth1
Public Property Get LineWidth1() As Integer
    LineWidth1 = LFTLNW
End Property
Public Property Let LineWidth1(ByVal nV As Integer)
    LFTLNW = nV
    If LFTLNW < 10 Or LFTLNW > 60 Then
        LFTLNW = 10
        MsgBox "Ova vrijednost može biti od 10 do 60!", vbExclamation, "Pogrešan unos"
    End If
    RedrawControl Normal
    PropertyChanged "LineWidth1"
End Property
'LineWidth2
Public Property Get LineWidth2() As Integer
    LineWidth2 = TOPLNW
End Property
Public Property Let LineWidth2(ByVal nV As Integer)
    TOPLNW = nV
    If TOPLNW < 20 Or TOPLNW > 60 Then
        TOPLNW = 20
        MsgBox "Ova vrijednost može biti od 20 do 60!", vbExclamation, "Pogrešan unos"
    End If
    RedrawControl Normal
    PropertyChanged "LineWidth2"
End Property
'caption
Public Property Get Caption() As String
    Caption = CAP
End Property
Public Property Let Caption(ByVal nV As String)
    CAP = nV
    RedrawControl Normal
    PropertyChanged "Caption"
End Property

'fore color
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = FC
End Property
Public Property Let ForeColor(ByVal nV As OLE_COLOR)
    FC = nV
    RedrawControl Normal
    PropertyChanged "ForeColor"
End Property
'fore color hover
Public Property Get ForeColorHover() As OLE_COLOR
    ForeColorHover = FCH
End Property
Public Property Let ForeColorHover(ByVal nV As OLE_COLOR)
    FCH = nV
    PropertyChanged "ForeColorHover"
End Property
'fore color down
Public Property Get ForeColorMouseDown() As OLE_COLOR
    ForeColorMouseDown = FCD
End Property
Public Property Let ForeColorMouseDown(ByVal nV As OLE_COLOR)
    FCD = nV
    PropertyChanged "ForeColorMouseDown"
End Property

'back color
Public Property Get BackColor() As OLE_COLOR
    BackColor = BC
End Property
Public Property Let BackColor(ByVal nV As OLE_COLOR)
    BC = nV
    RedrawControl Normal
    PropertyChanged "BackColor"
End Property
'back color hover
Public Property Get BackColorHover() As OLE_COLOR
    BackColorHover = BCH
End Property
Public Property Let BackColorHover(ByVal nV As OLE_COLOR)
    BCH = nV
    PropertyChanged "BackColorHover"
End Property
'back color down
Public Property Get BackColorMouseDown() As OLE_COLOR
    BackColorMouseDown = BCD
End Property
Public Property Let BackColorMouseDown(ByVal nV As OLE_COLOR)
    BCD = nV
    PropertyChanged "BackColorMouseDown"
End Property

'font
Public Property Get Font() As StdFont
    On Error Resume Next
    Set Font = FON
End Property
Public Property Set Font(ByVal nF As StdFont)
    On Error Resume Next
    Set FON = nF
    RedrawControl Normal
    PropertyChanged "Font"
End Property

'transparent color
Public Property Get TransparentColor() As OLE_COLOR
    TransparentColor = TC
End Property
Public Property Let TransparentColor(ByVal nV As OLE_COLOR)
    TC = nV
    RedrawControl Normal
    PropertyChanged "TransparentColor"
End Property



Private Sub lblCAP_Click()
    RedrawControl Down
    RaiseEvent Click
End Sub

Private Sub lblCAP_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RedrawControl Down
End Sub

Private Sub lblCAP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bHovering = False Then RedrawControl Hover
End Sub

Private Sub Timer1_Timer()
    Dim lpPos As POINTAPI
    Dim lhWnd As Long
    GetCursorPos lpPos
    lhWnd = WindowFromPoint(lpPos.x, lpPos.y)
    If lhWnd <> UserControl.hwnd And bHovering = True Then RedrawControl Normal
End Sub

Private Sub UserControl_Initialize()
    RedrawControl Normal
End Sub

Private Function RedrawControl(ByVal State As eState)
On Error Resume Next
    'Set ShapeForm = New clsTransForm
    'TC = &HC000C0
    UserControl.BackColor = TC
    UserControl.Cls
    HG = ScaleHeight
    WD = UserControl.ScaleWidth
    
    lblCAP.Caption = CAP
    Set lblCAP.Font = FON
    lblCAP.ForeColor = FC
    
    lblCAP.Left = LFTLNW - lblCAP.Width
    lblCAP.Top = HG - lblCAP.Height
    
    LFTLNW = 60
    TOPLNW = 10
    If State = Normal Then
    
        For i = LFTLNW / 2 To LFTLNW * 1.5
            Circle (i, LFTLNW / 2), LFTLNW / 2, BC, 3.14 / 2, 3.14
        Next i
        For i = LFTLNW / 2 To TOPLNW + LFTLNW / 2
            Circle (LFTLNW * 1.5, i), LFTLNW / 2, BC, 3.14 / 2, 3.14
            'MsgBox i
        Next i
        
        Line (LFTLNW / 2 - 1, 0)-(WD, TOPLNW), BC, BF
        Line (0, LFTLNW / 2 - 1)-(LFTLNW, HG), BC, BF
        'Line (WD / 2 + 5, HG - DOWNLNW)-(WD - RIGHTLNW * 1.5 - 5, HG), BC2, BF
        
        lblCAP.ForeColor = FC
        bHovering = False
        MPRESS = False
        
    ElseIf State = Hover Then
        For i = LFTLNW / 2 To LFTLNW * 1.5
            Circle (i, LFTLNW / 2), LFTLNW / 2, BCH, 3.14 / 2, 3.14
        Next i
        For i = LFTLNW / 2 To TOPLNW + LFTLNW / 2
            Circle (LFTLNW * 1.5, i), LFTLNW / 2, BCH, 3.14 / 2, 3.14
            'MsgBox i
        Next i
        
        Line (LFTLNW / 2 - 1, 0)-(WD, TOPLNW), BCH, BF
        Line (0, LFTLNW / 2 - 1)-(LFTLNW, HG), BCH, BF
        'Line (WD / 2 + 5, HG - DOWNLNW)-(WD - RIGHTLNW * 1.5 - 5, HG), BC2, BF
        
        lblCAP.ForeColor = FCH
        bHovering = True
        MPRESS = False
    Else
        For i = LFTLNW / 2 To LFTLNW * 1.5
            Circle (i, LFTLNW / 2), LFTLNW / 2, BCD, 3.14 / 2, 3.14
        Next i
        For i = LFTLNW / 2 To TOPLNW + LFTLNW / 2
            Circle (LFTLNW * 1.5, i), LFTLNW / 2, BCD, 3.14 / 2, 3.14
            'MsgBox i
        Next i
        
        Line (LFTLNW / 2 - 1, 0)-(WD, TOPLNW), BCD, BF
        Line (0, LFTLNW / 2 - 1)-(LFTLNW, HG), BCD, BF
        'Line (WD / 2 + 5, HG - DOWNLNW)-(WD - RIGHTLNW * 1.5 - 5, HG), BC2, BF
        
        lblCAP.ForeColor = FCD
        bHovering = False
        MPRESS = True
    End If
    DoEvents
    'createSkinnedForm UserControl.hWnd, UserControl.BackColor, UserControl.Height, UserControl.Width, StarTrekButton
End Function

Private Sub UserControl_InitProperties()
    BC = &HFF8080
    FC = vbBlack
    FCH = vbWhite
    FCD = vbWhite
    BCH = &HFFC0C0
    BCD = &HFF0000
    CAP = "StarTrekButtonB"
    TC = Ambient.BackColor
    Set FON = Ambient.Font
    
    RedrawControl Normal
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If UserControl.Point(x, y) <> TC Then
        RedrawControl Down
        RaiseEvent Click
    End If
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bHovering = False Then
       If UserControl.Point(x, y) <> TC Then RedrawControl Hover
    ElseIf bHovering = True And UserControl.Point(x, y) = TC Then
        RedrawControl Normal
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    TC = PropBag.ReadProperty("TransparentColor")
    BC = PropBag.ReadProperty("BackColor")
    BCH = PropBag.ReadProperty("BackColorHover")
    BCD = PropBag.ReadProperty("backColorMouseDown")
    
    FC = PropBag.ReadProperty("ForeColor")
    FCH = PropBag.ReadProperty("ForeColorHover")
    FCD = PropBag.ReadProperty("ForeColorMouseDown")
    
    CAP = PropBag.ReadProperty("Caption")
    Set FON = PropBag.ReadProperty("Font")
    
    LFTLNW = PropBag.ReadProperty("LineWidth1")
    TOPLNW = PropBag.ReadProperty("LineWidth2")
    
    RedrawControl Normal
End Sub

Private Sub UserControl_Resize()
    RedrawControl Normal
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "TransparentColor", TC
    PropBag.WriteProperty "BackColor", BC
    PropBag.WriteProperty "BackColorHover", BCH
    PropBag.WriteProperty "BackColorMouseDown", BCD
    PropBag.WriteProperty "ForeColor", FC
    PropBag.WriteProperty "ForeColorHover", FCH
    PropBag.WriteProperty "ForeColorMouseDown", FCD
    PropBag.WriteProperty "Caption", CAP
    PropBag.WriteProperty "Font", FON
    
    PropBag.WriteProperty "LineWidth1", LFTLNW
    PropBag.WriteProperty "LineWidth2", TOPLNW
End Sub


