VERSION 5.00
Begin VB.UserControl ToolBarButton 
   Appearance      =   0  'Flat
   BackStyle       =   0  'Transparent
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   Picture         =   "ToolBarButtob.ctx":0000
   PropertyPages   =   "ToolBarButtob.ctx":0342
   ScaleHeight     =   3870
   ScaleWidth      =   5145
   ToolboxBitmap   =   "ToolBarButtob.ctx":0377
   Begin VB.Timer tmrMouse 
      Interval        =   10
      Left            =   1830
      Top             =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   2010
      Width           =   1065
   End
   Begin VB.Shape Shape2 
      Height          =   765
      Left            =   690
      Shape           =   4  'Rounded Rectangle
      Top             =   2370
      Width           =   1155
   End
   Begin VB.Label lblButton 
      Height          =   735
      Left            =   90
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ToolBar Button"
      Height          =   195
      Left            =   1110
      TabIndex        =   0
      Top             =   870
      Width           =   1065
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   60
      Top             =   960
      Width           =   150
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   585
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "ToolBarButton"
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

Option Explicit

Private Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn _
    As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, _
    ByVal nCombineMode As Long) As Long
    
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd _
    As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Dim mChildFormRegion As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private bHovering, ENBL, MPRESS As Boolean
Private Const RGN_OR = 2
Private BC, RBC, FC, RFC, BORC, RBORC, GFBORC, GFBC, GFFC, DBC, DFC, DBORC, DOWNBC, DOWNFC, DOWNBORC As OLE_COLOR
Private PIC, RPIC, DPIC As StdPicture
Private FON As Font
Private CAP As String

Public Event Click()
Public Event MouseRollOver()
Public Event MouseRollOut()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private W, h, BW As Integer
Public Enum sbStyle
    Normal
    Graphical
End Enum
Private ST As sbStyle

Public Enum sbButtonPicture
    DefaulSize
    Sizable
End Enum
Private BUTPIC As sbButtonPicture

Public Enum sbPictureAlign
    Center
    Left
    LeftAndTop
    Right
    Top
    Down
End Enum
Private PCALG As sbPictureAlign

Public Enum sbShape
    Rectangle
    RoundedRectangle
    Elipse
End Enum
Private SH As sbShape

Public Enum sbBorderStyle
    None
    Solid
    Dash
    Dot
End Enum
Private BS As sbBorderStyle
Private GFBS As sbBorderStyle

Dim lpPos As POINTAPI
Dim lhWnd As Long



'down back color
Public Property Get ButtonDown_BackColor() As OLE_COLOR
    ButtonDown_BackColor = DOWNBC
End Property
Public Property Let ButtonDown_BackColor(ByVal nDOWNBC As OLE_COLOR)
    DOWNBC = nDOWNBC
    PropertyChanged "ButtonDown_BackColor"
End Property
'down fore color
Public Property Get ButtonDown_ForeColor() As OLE_COLOR
    ButtonDown_ForeColor = DOWNFC
End Property
Public Property Let ButtonDown_ForeColor(ByVal nDOWNFC As OLE_COLOR)
    DOWNFC = nDOWNFC
    PropertyChanged "ButtonDown_ForeColor"
End Property
'down fore color
Public Property Get ButtonDown_BorderColor() As OLE_COLOR
    ButtonDown_BorderColor = DOWNBORC
End Property
Public Property Let ButtonDown_BorderColor(ByVal nDOWNBORC As OLE_COLOR)
    DOWNBORC = nDOWNBORC
    PropertyChanged "ButtonDown_BorderColor"
End Property


'border width
Public Property Get BorderWidth() As Integer
    BorderWidth = BW
End Property
Public Property Let BorderWidth(ByVal nBW As Integer)
    BW = nBW
    If BW > 10 Then
        BW = 10
    End If
    Shape1.BorderWidth = BW
    PropertyChanged "BorderWidth"
End Property

'disabled picture
Public Property Get ButtonDisabled_Picture() As StdPicture
    Set ButtonDisabled_Picture = DPIC
End Property
Public Property Set ButtonDisabled_Picture(ByVal nDPIC As StdPicture)
    Set DPIC = nDPIC
    PropertyChanged "ButtonDisabled_Picture"
End Property
'disabled back color
Public Property Get ButtonDisabled_BackColor() As OLE_COLOR
    ButtonDisabled_BackColor = DBC
End Property
Public Property Let ButtonDisabled_BackColor(ByVal nDBC As OLE_COLOR)
    DBC = nDBC
    PropertyChanged "ButtonDisabled_BackColor"
End Property
'disabled fore color
Public Property Get ButtonDisabled_ForeColor() As OLE_COLOR
    ButtonDisabled_ForeColor = DFC
End Property
Public Property Let ButtonDisabled_ForeColor(ByVal nDFC As OLE_COLOR)
    DFC = nDFC
    PropertyChanged "ButtonDisabled_ForeColor"
End Property
'disabled border color
Public Property Get ButtonDisabled_BorderColor() As OLE_COLOR
    ButtonDisabled_BorderColor = DBORC
End Property
Public Property Let ButtonDisabled_BorderColor(ByVal nDBORC As OLE_COLOR)
    DBORC = nDBORC
    PropertyChanged "ButtonDisabled_BorderColor"
End Property
'enable
Public Property Get Enabled() As Boolean
    Enabled = ENBL
End Property
Public Property Let Enabled(ByVal nENBL As Boolean)
    On Error Resume Next
    ENBL = nENBL
    If ENBL = True Then
        UserControl.Enabled = True
        Shape1.BackColor = BC
        Shape1.BorderColor = BORC
        UserControl.BackColor = BC
        lbl.ForeColor = FC
        img.Picture = PIC
        tmrMouse.Enabled = True
    Else
        UserControl.Enabled = False
        Shape1.BackColor = DBC
        Shape1.BorderColor = DBORC
        UserControl.BackColor = DBC
        lbl.ForeColor = DFC
        img.Picture = DPIC
        tmrMouse.Enabled = False
    End If
End Property

'border style
Public Property Get BorderStyle() As sbBorderStyle
    BorderStyle = BS
End Property
Public Property Let BorderStyle(ByVal nBS As sbBorderStyle)
    BS = nBS
        If BS = Solid Then
            Shape2.BorderStyle = 1
        ElseIf BS = Dash Then
            Shape2.BorderStyle = 2
        ElseIf BS = Dot Then
            Shape2.BorderStyle = 3
        ElseIf BS = None Then
            Shape2.BorderStyle = 0
        End If
            'Shape1.BorderStyle = Shape2.BorderStyle
        PropertyChanged "BorderStyle"
End Property
'border got focus
Public Property Get ButtonGotFocus_BorderStyle() As sbBorderStyle
    ButtonGotFocus_BorderStyle = GFBS
End Property
Public Property Let ButtonGotFocus_BorderStyle(ByVal nGFBS As sbBorderStyle)
    GFBS = nGFBS
    PropertyChanged "ButtonGotFocus_BorderStyle"
End Property

'shape
Public Property Get Shape() As sbShape
    Shape = SH
End Property
Public Property Let Shape(ByVal nSH As sbShape)
    SH = nSH
    UserControl_Resize
    PropertyChanged "Shape"
End Property
'picture pos
Public Property Get PicturePos() As sbPictureAlign
    PicturePos = PCALG
End Property
Public Property Let PicturePos(ByVal nPCALG As sbPictureAlign)
    PCALG = nPCALG
    UserControl_Resize
    PropertyChanged "PicturePos"
End Property
'style
Public Property Get Style() As sbStyle
    Style = ST
End Property
Public Property Let Style(ByVal nST As sbStyle)
    ST = nST
    If Style = Graphical Then
        img.Visible = True
    Else
        img.Visible = False
    End If
    UserControl_Resize
    PropertyChanged "Style"
End Property
'button picture size
Public Property Get ButtonPictureSize() As sbButtonPicture
    ButtonPictureSize = BUTPIC
End Property
Public Property Let ButtonPictureSize(ByVal nBUTPIC As sbButtonPicture)
    BUTPIC = nBUTPIC
    PropertyChanged "ButtonPictureSize"
End Property
'picture height
Public Property Get PictureHeight() As Integer
    PictureHeight = h
End Property
Public Property Let PictureHeight(ByVal nH As Integer)
    h = nH
    If BUTPIC = DefaulSize Then
        img.Stretch = False
        img.Picture = PIC
        h = img.Height
    Else
        img.Stretch = True
        img.Height = h
    End If
End Property
'picture width
Public Property Get PictureWidth() As Integer
    PictureWidth = W
End Property
Public Property Let PictureWidth(ByVal nW As Integer)
    W = nW
    If BUTPIC = DefaulSize Then
        img.Stretch = False
        img.Picture = PIC
        W = img.Width
    Else
        img.Stretch = True
        img.Width = W
    End If
End Property
'font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal nFON As Font)
    Set UserControl.Font = nFON
    Set lbl.Font = UserControl.Font
    PropertyChanged "Font"
    UserControl_Resize
End Property
'picture
Public Property Get Picture() As StdPicture
On Error Resume Next
    Set Picture = PIC
End Property
Public Property Set Picture(ByVal nPC As StdPicture)
    Set PIC = nPC
    img.Picture = PIC
    UserControl_Resize
    PropertyChanged "Picture"
    
    W = img.Width
    h = img.Height
    PictureHeight = h
    PictureWidth = W
    
    On Error GoTo ERRNoPicture
    If RPIC.Height = 0 Then
        Set RPIC = PIC
    End If
    Exit Property
ERRNoPicture:
        Set RPIC = PIC
        DisPic
End Property
Private Sub DisPic()
On Error GoTo ERRNoPicture1
    Dim r As Double
        r = DPIC.Height
    Exit Sub
ERRNoPicture1:
    Set DPIC = PIC
End Sub
'roll over picture
Public Property Get PictureRollOver() As StdPicture
On Error Resume Next
    Set PictureRollOver = RPIC
End Property
Public Property Set PictureRollOver(ByVal nRPIC As StdPicture)
    Set RPIC = nRPIC
    PropertyChanged "PictureRollOver"
End Property
'back color
Public Property Get BackColor() As OLE_COLOR
    BackColor = BC
End Property
Public Property Let BackColor(ByVal nBC As OLE_COLOR)
    BC = nBC
    UserControl.BackColor = BC
    Shape1.BackColor = BC
    PropertyChanged "BackColor"
End Property
'back color roll over
Public Property Get BackColorRollOver() As OLE_COLOR
    BackColorRollOver = RBC
End Property
Public Property Let BackColorRollOver(ByVal nRBC As OLE_COLOR)
    RBC = nRBC
    PropertyChanged "BackColorRollOver"
End Property
'fore color
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = FC
End Property
Public Property Let ForeColor(ByVal nFC As OLE_COLOR)
    FC = nFC
    lbl.ForeColor = FC
    PropertyChanged "ForeColor"
End Property
'fore color roll over
Public Property Get ForeColorRollOver() As OLE_COLOR
    ForeColorRollOver = RFC
End Property
Public Property Let ForeColorRollOver(ByVal nRFC As OLE_COLOR)
    RFC = nRFC
    PropertyChanged "ForeColorRollOver"
End Property
'border color
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = BORC
End Property
Public Property Let BorderColor(ByVal nBORC As OLE_COLOR)
    BORC = nBORC
    Shape1.BorderColor = BORC
    PropertyChanged "BorderColor"
End Property
'border color roll over
Public Property Get BorderColorRollOver() As OLE_COLOR
    BorderColorRollOver = RBORC
End Property
Public Property Let BorderColorRollOver(ByVal nRBORC As OLE_COLOR)
    RBORC = nRBORC
    PropertyChanged "BorderColorRollOver"
End Property
'border color got focus
Public Property Get ButtonGotFocus_BorderColor() As OLE_COLOR
    ButtonGotFocus_BorderColor = GFBORC
End Property
Public Property Let ButtonGotFocus_BorderColor(ByVal nGFBORC As OLE_COLOR)
    GFBORC = nGFBORC
    PropertyChanged "ButtonGotFocus_BorderColor"
End Property
'forecolor got focus
Public Property Get ButtonGotFocus_ForeColor() As OLE_COLOR
    ButtonGotFocus_ForeColor = GFFC
End Property
Public Property Let ButtonGotFocus_ForeColor(ByVal nGFFC As OLE_COLOR)
    GFFC = nGFFC
    PropertyChanged "ButtonGotFocus_ForeColor"
End Property
'caption
Public Property Get Caption() As String
    Caption = CAP
End Property
Public Property Let Caption(ByVal nCAP As String)
    CAP = nCAP
    lbl.Caption = CAP
    PropertyChanged "Caption"
    UserControl_Resize
End Property

Private Sub MouseOut()
On Error Resume Next
    UserControl.BackColor = BC
    Shape1.BackColor = BC
    lbl.ForeColor = FC
    Shape1.BorderColor = BORC
    img.Picture = PIC
    RaiseEvent MouseRollOut
    UserControl_Resize
    bHovering = False
    tmrMouse.Enabled = False
End Sub
Private Sub RollOver()
On Error Resume Next
    If bHovering = False Then
        tmrMouse.Enabled = True
        bHovering = True
        UserControl.BackColor = RBC
        Shape1.BackColor = RBC
        lbl.ForeColor = RFC
        Shape1.BorderColor = RBORC
        img.Picture = RPIC
        RaiseEvent MouseRollOver
        UserControl_Resize
    End If
End Sub


'img click
Private Sub img_Click()
    RaiseEvent Click
End Sub
'img mouse move
Private Sub img_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RollOver
End Sub
'lbl click
Private Sub lbl_Click()
    RaiseEvent Click
End Sub
'lbl mouse move
Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RollOver
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
'lbl button click
Private Sub lblButton_Click()
    RaiseEvent Click
End Sub
'timer
Private Sub tmrMouse_Timer()
    GetCursorPos lpPos
    lhWnd = WindowFromPoint(lpPos.x, lpPos.y)
    If lhWnd <> UserControl.hwnd And bHovering = True Then MouseOut
End Sub
'userconstrol click
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
'userconstrol got focus
Private Sub UserControl_GotFocus()
    Shape1.BorderColor = GFBORC
        If GFBS = Solid Then
            Shape2.BorderStyle = 1
        ElseIf GFBS = Dash Then
            Shape2.BorderStyle = 2
        ElseIf GFBS = Dot Then
            Shape2.BorderStyle = 3
        ElseIf GFBS = None Then
            Shape2.BorderStyle = 0
        End If
    lbl.ForeColor = GFFC
End Sub
'userconstrol lost focus
Private Sub UserControl_LostFocus()
    Shape1.BorderColor = BORC
        If BS = Solid Then
            Shape2.BorderStyle = 1
        ElseIf BS = Dash Then
            Shape2.BorderStyle = 2
        ElseIf BS = Dot Then
            Shape2.BorderStyle = 3
        ElseIf BS = None Then
            Shape2.BorderStyle = 0
        End If
    lbl.ForeColor = FC
End Sub
'userconstrol key down
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
'userconstrol key press
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub
'userconstrol key up
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub
'lblbutton mousedown
Private Sub lblButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
    MPRESS = True
    lbl.ForeColor = DOWNFC
    Shape1.BackColor = DOWNBC
    Shape1.BorderColor = DOWNBORC
    UserControl.BackColor = DOWNBC
End Sub
'lblbutton mouse move
Private Sub lblButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MPRESS = False Then RollOver
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
'lblbutton mouse up
Private Sub lblButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    
    Dim lpPos As POINTAPI
    Dim lhWnd As Long
    GetCursorPos lpPos
    lhWnd = WindowFromPoint(lpPos.x, lpPos.y)
    If lhWnd <> UserControl.hwnd Then
        lbl.ForeColor = FC
        Shape1.BackColor = BC
        Shape1.BorderColor = BORC
    Else
        lbl.ForeColor = RFC
        Shape1.BackColor = RBC
        Shape1.BorderColor = RBORC
    End If
    MPRESS = False
End Sub


'user control read properties
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Set Font = UserControl.Font
    Set lbl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    BC = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    RBC = PropBag.ReadProperty("BackColorRollOver", vbWhite)
    FC = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
    RFC = PropBag.ReadProperty("ForeColorRollOver", Ambient.ForeColor)
    Set PIC = PropBag.ReadProperty("Picture", "")
    Set RPIC = PropBag.ReadProperty("PictureRollOver", "")
    Set DPIC = PropBag.ReadProperty("ButtonDisabled_Picture", "")
    CAP = PropBag.ReadProperty("Caption", "ToolBar Button")
    ST = PropBag.ReadProperty("Style", Graphical)
    W = PropBag.ReadProperty("PictureWidth", img.Width)
    h = PropBag.ReadProperty("PictureHeight", img.Height)
    BUTPIC = PropBag.ReadProperty("ButtonPictureSize", DefaulSize)
    PCALG = PropBag.ReadProperty("PicturePos", Center)
    BORC = PropBag.ReadProperty("BorderColor", vbBlack)
    RBORC = PropBag.ReadProperty("BorderColorRollOver", vbBlack)
    SH = PropBag.ReadProperty("Shape", RoundedRectangle)
    Shape1.BorderColor = BORC
    BS = PropBag.ReadProperty("BorderStyle", Solid)
    
    GFBS = PropBag.ReadProperty("ButtonGotFocus_BorderStyle", Solid)
    GFBORC = PropBag.ReadProperty("ButtonGotFocus_BorderColor", vbBlack)
    GFFC = PropBag.ReadProperty("ButtonGotFocus_ForeColor", FC)
    
    DBC = PropBag.ReadProperty("ButtonDisabled_BackColor", BC)
    ENBL = PropBag.ReadProperty("Enabled", True)
    DBORC = PropBag.ReadProperty("ButtonDisabled_BorderColor", BORC)
    DFC = PropBag.ReadProperty("ButtonDisabled_ForeColor", FC)
    BW = PropBag.ReadProperty("BorderWidth", 1)
    DOWNBORC = PropBag.ReadProperty("ButtonDown_BorderColor", BORC)
    DOWNBC = PropBag.ReadProperty("ButtonDown_BackColor", BC)
    DOWNFC = PropBag.ReadProperty("ButtonDown_ForeColor", FC)
    
    Shape1.BorderWidth = BW
    
        tmrMouse.Enabled = ENBL
        
    
    
    
        If BS = Solid Then
            Shape2.BorderStyle = 1
        ElseIf BS = Dash Then
            Shape2.BorderStyle = 2
        ElseIf BS = Dot Then
            Shape2.BorderStyle = 3
        ElseIf BS = None Then
            Shape2.BorderStyle = 0
        End If
    
    Set lbl.Font = FON
    lbl.ForeColor = FC
    UserControl.BackColor = BC
    Shape1.BackColor = BC
    lbl.Caption = CAP
    img.Picture = PIC
    
    If BUTPIC = DefaulSize Then
        img.Stretch = False
    Else
        img.Stretch = True
        img.Height = h
        img.Width = W
    End If
    
    If Style = Graphical Then
        img.Visible = True
    Else
        img.Visible = False
    End If
    
    If ENBL = True Then
            UserControl.Enabled = True
            Shape1.BackColor = BC
            Shape1.BorderColor = BORC
            UserControl.BackColor = BC
            lbl.ForeColor = FC
    Else
            UserControl.Enabled = False
            Shape1.BackColor = DBC
            Shape1.BorderColor = DBORC
            UserControl.BackColor = DBORC
            lbl.ForeColor = DFC
    End If
    
    UserControl_Resize
End Sub
'user control write properties
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    PropBag.WriteProperty "Caption", CAP, "ToolBar Button"
    PropBag.WriteProperty "BackColor", BC, Ambient.BackColor
    PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
    PropBag.WriteProperty "ForeColor", FC, Ambient.ForeColor
    PropBag.WriteProperty "Picture", PIC
    PropBag.WriteProperty "PictureRollOver", RPIC
    PropBag.WriteProperty "BackColorRollOver", RBC, vbWhite
    PropBag.WriteProperty "ForeColorRollOver", RFC, Ambient.ForeColor
    PropBag.WriteProperty "Style", ST, Graphical
    PropBag.WriteProperty "PictureHeight", h, img.Height
    PropBag.WriteProperty "PictureWidth", W, img.Width
    PropBag.WriteProperty "ButtonPictureSize", BUTPIC, DefaulSize
    PropBag.WriteProperty "PicturePos", PCALG, Center
    PropBag.WriteProperty "BorderColor", BORC, vbBlack
    PropBag.WriteProperty "BorderColorRollOver", RBORC, vbBlack
    PropBag.WriteProperty "Shape", SH, RoundedRectangle
    PropBag.WriteProperty "ButtonGotFocus_BorderColor", GFBORC
    PropBag.WriteProperty "BorderStyle", BS, Solid
    PropBag.WriteProperty "ButtonGotFocus_BorderStyle", GFBS
    PropBag.WriteProperty "ButtonGotFocus_ForeColor", GFFC, FC
    
    PropBag.WriteProperty "ButtonDisabled_BackColor", DBC, BC
    PropBag.WriteProperty "Enabled", ENBL, True
    PropBag.WriteProperty "ButtonDisabled_BorderColor", DBORC
    PropBag.WriteProperty "ButtonDisabled_ForeColor", DFC
    
    PropBag.WriteProperty "BorderWidth", BW, 1
    
    PropBag.WriteProperty "ButtonDown_BorderColor", DOWNBORC
    PropBag.WriteProperty "ButtonDown_BackColor", DOWNBC
    PropBag.WriteProperty "ButtonDown_ForeColor", DOWNFC
    PropBag.ReadProperty "ButtonDisabled_Picture", DPIC
End Sub
'user control int properties
Private Sub UserControl_InitProperties()
On Error Resume Next
    BC = &HFF8080
    RBC = &HFFC0C0
    FC = vbBlack
    RFC = vbBlack
    
    Set UserControl.Font = Ambient.Font
    Set Font = UserControl.Font
    W = img.Width
    h = img.Height
    
    BorderColor = BORC
    BorderColorRollOver = RBORC
    Set Picture = PIC
    Set PictureRollOver = RPIC
    CAP = "ToolBar Button"
    ST = Graphical
    SH = Rectangle
    
    GFBORC = vbBlack
    BS = None
    GFBS = None
    GFFC = FC
    
    ENBL = True
    DBC = &H80000016
    DFC = &H80000011
    DBORC = &H80000010
    
    DOWNFC = FC
    DOWNBC = &HFF0000
    DOWNBORC = BORC
    
    BW = 1
    Shape1.BorderWidth = BW
    
    Shape1.BackColor = BC
    UserControl_Resize
End Sub
'user control initalize
Private Sub UserControl_Initialize()
    lblButton.Top = 0
    lblButton.Height = 0
    lblButton.BackStyle = 0
    
    Shape2.Top = 0
    Shape2.Left = 0
    
    UserControl_Resize
    
    On Error GoTo ERRNoPicture
        Dim var As Double
        var = Picture.Height
    Exit Sub
    
ERRNoPicture:
        img.Height = 60
        img.Width = 60
        UserControl_Resize
End Sub
'user control resize
Private Sub UserControl_Resize()
    RefreshButton
End Sub

Private Sub RefreshButton()
    'xp = Screen.TwipsPerPixelX
    'yp = Screen.TwipsPerPixelY
    
    If SH = Rectangle Then
        Shape2.Shape = 0
        'mChildFormRegion = CreateRoundRectRgn(0, 0, UserControl.Width / xp, UserControl.Height / yp, 0, 0)
        'SetWindowRgn UserControl.hwnd, mChildFormRegion, False
    ElseIf SH = RoundedRectangle Then
        Shape2.Shape = 4
       ' mChildFormRegion = CreateRoundRectRgn(0, 0, UserControl.Width / xp, UserControl.Height / yp, 5, 5)
        'SetWindowRgn UserControl.hwnd, mChildFormRegion, False
    ElseIf SH = Elipse Then
        Shape2.Shape = 2
    End If
    
    Shape1.Shape = Shape2.Shape
    
    Shape1.Height = UserControl.Height - 20
    Shape1.Width = UserControl.Width - 20
    
    Shape2.Height = Shape1.Height
    Shape2.Width = Shape1.Width
    
    lblButton.Height = UserControl.Height
    lblButton.Width = UserControl.Width
    
    If ST = Normal Then
        lbl.Top = UserControl.Height / 2 - lbl.Height / 2
        lbl.Left = UserControl.Width / 2 - lbl.Width / 2
        If lbl.Left < 50 Then lbl.Left = 50
    Else
        If PCALG = Center Then
            img.Top = UserControl.Height / 2 - img.Height / 2
            img.Left = UserControl.Width / 2 - img.Width / 2
                If CAP <> "" Then img.Top = UserControl.Height / 2 - img.Height / 2 - lbl.Height
                    lbl.Left = UserControl.Width / 2 - lbl.Width / 2
                    lbl.Top = img.Top + img.Height + 50
        ElseIf PCALG = Down Then
            img.Top = UserControl.Height - img.Height - 50
            img.Left = UserControl.Width / 2 - img.Width / 2
                If CAP <> "" And img.Top - lbl.Height < 50 Then img.Top = UserControl.Height - img.Height - 50 + lbl.Height
                    lbl.Left = UserControl.Width / 2 - lbl.Width / 2
                    lbl.Top = img.Top - lbl.Height - 50
        ElseIf PCALG = Top Then
            img.Top = 50
            img.Left = UserControl.Width / 2 - img.Width / 2
                If CAP <> "" And img.Top + lbl.Height + 50 > UserControl.Height - 50 Then img.Top = img.Top - lbl.Height
                    lbl.Left = UserControl.Width / 2 - lbl.Width / 2
                    lbl.Top = img.Top + img.Height + 50
        ElseIf PCALG = Left Then
            img.Top = UserControl.Height / 2 - img.Height / 2
            img.Left = 50
                lbl.Top = UserControl.Height / 2 - lbl.Height / 2
                lbl.Left = img.Left + img.Width + 60
        ElseIf PCALG = LeftAndTop Then
            img.Top = 30
            img.Left = 30
                lbl.Top = UserControl.Height / 2 - lbl.Height / 2
                lbl.Left = img.Left + img.Width + 60
        ElseIf PCALG = Right Then
            img.Top = UserControl.Height / 2 - img.Height / 2
            img.Left = UserControl.Width - img.Width - 50
                lbl.Top = UserControl.Height / 2 - lbl.Height / 2
                lbl.Left = lbl.Width / 2 - lbl.Width / 2
        End If
    End If
End Sub




