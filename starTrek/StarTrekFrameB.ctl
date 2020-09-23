VERSION 5.00
Begin VB.UserControl StarTrekFrameMC_A 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7860
   ControlContainer=   -1  'True
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   524
   ToolboxBitmap   =   "StarTrekFrameB.ctx":0000
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H0080C0FF&
      Height          =   195
      Left            =   630
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "StarTrekFrameMC_A"
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

Private i, k, LFTLNW, TOPLNW, DOWNLNW, RIGHTLNW As Integer
Private WD, HG As Long
Private BC, TC, FC, BC2 As OLE_COLOR
Private CAP As String
Private FON As StdFont

'caption
Public Property Get Caption() As String
    Caption = CAP
End Property
Public Property Let Caption(ByVal nV As String)
    CAP = nV
    RedrawControl
    PropertyChanged "Caption"
End Property
'LineWidth_Left
Public Property Get LineWidth_Left() As Integer
    LineWidth_Left = LFTLNW
End Property
Public Property Let LineWidth_Left(ByVal nV As Integer)
    LFTLNW = nV
    If LFTLNW < 10 Or LFTLNW > 60 Then
        LFTLNW = 10
        MsgBox "Ova vrijednost može biti od 10 do 60!", vbExclamation, "Pogrešan unos"
    End If
    RedrawControl
    PropertyChanged "LineWidth_Left"
End Property
'LineWidth_Top
Public Property Get LineWidth_Top() As Integer
    LineWidth_Top = TOPLNW
End Property
Public Property Let LineWidth_Top(ByVal nV As Integer)
    TOPLNW = nV
    If TOPLNW < 20 Or TOPLNW > 60 Then
        TOPLNW = 20
        MsgBox "Ova vrijednost može biti od 20 do 60!", vbExclamation, "Pogrešan unos"
    End If
    RedrawControl
    PropertyChanged "LineWidth_Top"
End Property
'LineWidth_Botton
Public Property Get LineWidth_Botton() As Integer
    LineWidth_Botton = DOWNLNW
End Property
Public Property Let LineWidth_Botton(ByVal nV As Integer)
    DOWNLNW = nV
    If DOWNLNW < 10 Or DOWNLNW > 60 Then
        DOWNLNW = 10
        MsgBox "Ova vrijednost može biti od 10 do 60!", vbExclamation, "Pogrešan unos"
    End If
    RedrawControl
    PropertyChanged "LineWidth_Botton"
End Property
'LineWidth_Right
Public Property Get LineWidth_Right() As Integer
    LineWidth_Right = RIGHTLNW
End Property
Public Property Let LineWidth_Right(ByVal nV As Integer)
    RIGHTLNW = nV
    If RIGHTLNW < 10 Or RIGHTLNW > 60 Then
        RIGHTLNW = 10
        MsgBox "Ova vrijednost može biti od 10 do 60!", vbExclamation, "Pogrešan unos"
    End If
    RedrawControl
    PropertyChanged "LineWidth_Right"
End Property
'border color 1
Public Property Get BorderColor1() As OLE_COLOR
    BorderColor1 = BC
End Property
Public Property Let BorderColor1(ByVal nV As OLE_COLOR)
    BC = nV
    RedrawControl
    PropertyChanged "BorderColor1"
End Property
'border color 2
Public Property Get BorderColor2() As OLE_COLOR
    BorderColor2 = BC2
End Property
Public Property Let BorderColor2(ByVal nV As OLE_COLOR)
    BC2 = nV
    RedrawControl
    PropertyChanged "BorderColor2"
End Property
'fore color
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = FC
End Property
Public Property Let ForeColor(ByVal nV As OLE_COLOR)
    FC = nV
    RedrawControl
    PropertyChanged "ForeColor"
End Property
'transparent color
Public Property Get TransparentColor() As OLE_COLOR
    TransparentColor = TC
End Property
Public Property Let TransparentColor(ByVal nV As OLE_COLOR)
    TC = nV
    RedrawControl
    PropertyChanged "TransparentColor"
End Property
'font
Public Property Get Font() As StdFont
    On Error Resume Next
    Set Font = FON
End Property
Public Property Set Font(ByVal nF As StdFont)
    Set FON = nF
    RedrawControl
    PropertyChanged "Font"
End Property

Private Function RedrawControl()
    On Error Resume Next
    WD = Width / Screen.TwipsPerPixelX
    HG = Height / Screen.TwipsPerPixelY
    'MsgBox HG
    'poèetne vrijednosti
    'BC = &H80FF&
    'BC2 = &H80C0FF
    
   ' LFTLNW = 60
    'TOPLNW = 20
    'DOWNLNW = 20
    'RIGHTLNW = 60
    '***************
    '*******lblcap***********
    lblCAP.Left = LFTLNW * 1.5 + 5
    lblCAP.Caption = CAP
    Set lblCAP.Font = FON
    lblCAP.ForeColor = FC
    UserControl.BackColor = TC
    
    
    UserControl.Cls
    DoEvents
    'If lblCAP.Height < TOPLNW Then
        lblCAP.Top = TOPLNW / 2 - lblCAP.Height / 2
    'Else
       ' lblCAP.Top = lblCAP.Height / 2 - TOPLNW / 2
        'MsgBox lblCAP.Top
   ' End If
    
    'If lblCAP.Top < 0 Then
       ' lblCAP.Top = 0
        'K = lblCAP.Height / 2
   ' Else
       ' K = 0
   ' End If
    '****************************
    
    'left
    'lijevi kut gore
    For i = LFTLNW / 2 To LFTLNW * 1.5
        Circle (i, LFTLNW / 2), LFTLNW / 2, BC, 3.14 / 2, 3.14
    Next i
    For i = LFTLNW / 2 To TOPLNW + LFTLNW / 2
        Circle (LFTLNW * 1.5, i), LFTLNW / 2, BC, 3.14 / 2, 3.14
        'MsgBox i
    Next i

    'linije lijevo
    Line (0, LFTLNW / 2 - 1)-(LFTLNW, LFTLNW + TOPLNW), BC, BF
    Line (0, LFTLNW + TOPLNW + 5)-(LFTLNW, HG - (LFTLNW / 2 + 5 + DOWNLNW)), BC2, BF
    Line (0, HG - (LFTLNW / 2 + DOWNLNW))-(LFTLNW, HG - LFTLNW / 2), BC, BF
    
    'botton
    'dolje kut lijevo
    For i = LFTLNW / 2 To LFTLNW * 1.5
        Circle (i, HG - LFTLNW / 2), LFTLNW / 2, BC, 3.14, 1.5 * 3.14
    Next i
    For i = LFTLNW / 2 To DOWNLNW + LFTLNW / 2
        Circle (LFTLNW * 1.5, HG - i), LFTLNW / 2, BC, 3.14, 1.5 * 3.14
    Next i
    
    Line (lblCAP.Left + lblCAP.Width + 5, 0)-(WD - RIGHTLNW * 1.5 - 5, TOPLNW), BC2, BF
    Line (LFTLNW, HG)-(WD / 2, HG - DOWNLNW), BC, BF
    Line (WD / 2 + 5, HG - DOWNLNW)-(WD - RIGHTLNW * 1.5 - 5, HG), BC2, BF
    
    'right
    'desni kut dolje
    For i = RIGHTLNW / 2 To RIGHTLNW * 1.5
        Circle (WD - i, HG - RIGHTLNW / 2), RIGHTLNW / 2, BC, 1.5 * 3.14, 2 * 3.14
    Next i
    
    For i = RIGHTLNW / 2 To DOWNLNW + RIGHTLNW / 2
        Circle (WD - RIGHTLNW * 1.5, HG - i), RIGHTLNW / 2, BC, 1.5 * 3.14, 2 * 3.14
    Next i
    
    'desni kut gore
    For i = RIGHTLNW / 2 To RIGHTLNW * 1.5
        Circle (WD - i, RIGHTLNW / 2), RIGHTLNW / 2, BC, 0, 3.14 / 2
    Next i
    
    For i = RIGHTLNW / 2 To TOPLNW + RIGHTLNW / 2
        Circle (WD - RIGHTLNW * 1.5, i), RIGHTLNW / 2, BC, 0, (3.14 / 2)
    Next i
    
    Line (WD - RIGHTLNW, HG - RIGHTLNW / 2)-(WD, HG - DOWNLNW - RIGHTLNW), BC, BF
    Line (WD, HG - DOWNLNW - RIGHTLNW - 5)-(WD - RIGHTLNW, TOPLNW + RIGHTLNW * 1.5), BC2, BF
    Line (WD - RIGHTLNW, TOPLNW + RIGHTLNW * 1.5 - 5)-(WD, RIGHTLNW / 2), BC, BF
End Function

Private Sub UserControl_InitProperties()
    'On Error Resume Next
    BC = &H80FF&
    BC2 = &H80C0FF
    
    LFTLNW = 60
    TOPLNW = 20
    DOWNLNW = 20
    RIGHTLNW = 60
    
    Set FON = Ambient.Font
    TC = Ambient.BackColor
    FC = &H80C0FF
    
    CAP = "StarTrekFrame"
    
    RedrawControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    CAP = PropBag.ReadProperty("Caption")
    BC = PropBag.ReadProperty("BorderColor1")
    BC2 = PropBag.ReadProperty("BorderColor2")
    FC = PropBag.ReadProperty("ForeColor")
    Set FON = PropBag.ReadProperty("Font")
    TC = PropBag.ReadProperty("TransparentColor")
    
    LFTLNW = PropBag.ReadProperty("LineWidth_Left")
    TOPLNW = PropBag.ReadProperty("LineWidth_Top")
    RIGHTLNW = PropBag.ReadProperty("LineWidth_Right")
    DOWNLNW = PropBag.ReadProperty("LineWidth_Botton")
    
    RedrawControl
End Sub

Private Sub UserControl_Resize()
    RedrawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", CAP
    PropBag.WriteProperty "BorderColor1", BC
    PropBag.WriteProperty "BorderColor2", BC2
    PropBag.WriteProperty "ForeColor", FC
    PropBag.WriteProperty "Font", FON
    PropBag.WriteProperty "TransparentColor", TC
    
    PropBag.WriteProperty "LineWidth_Left", LFTLNW
    PropBag.WriteProperty "LineWidth_Top", TOPLNW
    PropBag.WriteProperty "LineWidth_Right", RIGHTLNW
    PropBag.WriteProperty "LineWidth_Botton", DOWNLNW
End Sub
