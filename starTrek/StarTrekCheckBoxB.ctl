VERSION 5.00
Begin VB.UserControl StarTrekCheckBoxB 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   112
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "StarTrekCheckBoxB.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   270
      Top             =   510
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "StarTrekCheckBoxB"
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
Private BC, TC, BCH, BCD, FCD, FC, FCH, CC As OLE_COLOR
Private CAP As String
Private FON As StdFont

Dim i As Integer

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

Private bHovering, ENBL, MPRESS, CH As Boolean
Public Event Click()
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
'check color
Public Property Get CheckColor() As OLE_COLOR
    CheckColor = CC
End Property
Public Property Let CheckColor(ByVal nV As OLE_COLOR)
    CC = nV
    RedrawControl Normal
    PropertyChanged "CheckColor"
End Property
'value
Public Property Get Checked() As Boolean
    Checked = CH
End Property
Public Property Let Checked(ByVal nV As Boolean)
    CH = nV
    RedrawControl Normal
    PropertyChanged "Checked"
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

Private Function CHC()
    If CH = False Then
        CH = True
    Else
        CH = False
    End If
    RedrawControl Normal
End Function



Private Sub lblCAP_Click()
    RedrawControl Down
    RaiseEvent Click
    CHC
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

Private Sub UserControl_Click()
    CHC
    RaiseEvent Click
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
    
    lblCAP.Left = HG / 2 + 10
    lblCAP.Top = HG / 2 - lblCAP.Height / 2
    
    If State = Normal Then
        For i = HG / 2 To HG * 1.5
            Circle (i, HG / 2), HG / 2, BC, (3.14 / 2), ((3 / 2) * 3.14)
        Next i
        
        For i = WD - HG * 1.2 To WD - HG / 2
            Circle (i, HG / 2), HG / 2, BC, (3 / 2) * 3.14, 3.14 / 2
        Next i
        
        Line (HG / 2, 0)-(WD - HG / 2, HG), BC, BF
        Line (HG / 2, 0)-(HG / 2 + 8, HG), TC, BF
        
        lblCAP.ForeColor = FC
        bHovering = False
        MPRESS = False
        
    ElseIf State = Hover Then
        For i = HG / 2 To HG * 1.5
            Circle (i, HG / 2), HG / 2, BC, (3.14 / 2), ((3 / 2) * 3.14)
        Next i
        
        For i = WD - HG * 1.2 To WD - HG / 2
            Circle (i, HG / 2), HG / 2, BCH, (3 / 2) * 3.14, 3.14 / 2
        Next i
        
        Line (HG / 2, 0)-(WD - HG / 2, HG), BCH, BF
        Line (HG / 2, 0)-(HG / 2 + 8, HG), TC, BF
        lblCAP.ForeColor = FCH
        bHovering = True
        MPRESS = False
    Else
        For i = HG / 2 To HG * 1.5
            Circle (i, HG / 2), HG / 2, BC, (3.14 / 2), ((3 / 2) * 3.14)
        Next i
        
        For i = WD - HG * 1.2 To WD - HG / 2
            Circle (i, HG / 2), HG / 2, BCD, (3 / 2) * 3.14, 3.14 / 2
        Next i
        
        Line (HG / 2, 0)-(WD - HG / 2, HG), BCD, BF
        Line (HG / 2, 0)-(HG / 2 + 8, HG), TC, BF
        lblCAP.ForeColor = FCD
        bHovering = False
        MPRESS = True
    End If
    
    Line (WD - HG / 2 - 8, 0)-(WD - HG / 2, HG), TC, BF
    
    If CH = True Then
        Line (HG / 2 + 2, 0)-(HG / 2 + 6, HG), CC, BF
        Line (WD - HG / 2 - 6, 0)-(WD - HG / 2 - 2, HG), CC, BF
        
    End If
    
    DoEvents
    'createSkinnedForm UserControl.hWnd, UserControl.BackColor, UserControl.Height, UserControl.Width, StarTrekButton
End Function

Private Sub UserControl_InitProperties()
    BC = &H80FF&
    FC = vbBlack
    FCH = vbWhite
    FCD = vbWhite
    BCH = &H8FFF&
    BCD = &H404080
    CAP = "StarTrekCheckBox"
    TC = Ambient.BackColor
    CC = &H80C0FF
    Set FON = Ambient.Font
    
    RedrawControl Normal
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RedrawControl Down
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bHovering = False Then
       If UserControl.Point(x, y) <> TC Then RedrawControl Hover
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
    
    CH = PropBag.ReadProperty("Checked", False)
    CC = PropBag.ReadProperty("CheckColor")
    
    CAP = PropBag.ReadProperty("Caption")
    Set FON = PropBag.ReadProperty("Font")
    
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
    
    PropBag.WriteProperty "Checked", CH
    PropBag.WriteProperty "CheckColor", CC
End Sub


