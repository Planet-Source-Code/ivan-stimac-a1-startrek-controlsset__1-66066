VERSION 5.00
Begin VB.UserControl StarTrekProgressBarA 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "StarTrekProgressBar.ctx":0000
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "StarTrekProgressBarA"
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

Private HG, WD, maxVAL, MINVAL, VAL1 As Long
Private BC, TC, FC, BC1, BC2 As OLE_COLOR
Private CAP As String
Private FON As StdFont

Dim i As Integer

Private Enum eState
    Normal
    Hover
    Down
End Enum

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
'back color
Public Property Get BackColor() As OLE_COLOR
    BackColor = BC
End Property
Public Property Let BackColor(ByVal nV As OLE_COLOR)
    BC = nV
    RedrawControl Normal
    PropertyChanged "BackColor"
End Property
'back color scroll
Public Property Get BackColorScroll() As OLE_COLOR
    BackColorScroll = BC1
End Property
Public Property Let BackColorScroll(ByVal nV As OLE_COLOR)
    BC1 = nV
    RedrawControl Normal
    PropertyChanged "BackColorScroll"
End Property
'back color scroll2
Public Property Get BackColorScroll2() As OLE_COLOR
    BackColorScroll2 = BC2
End Property
Public Property Let BackColorScroll2(ByVal nV As OLE_COLOR)
    BC2 = nV
    RedrawControl Normal
    PropertyChanged "BackColorScroll2"
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
'max val
Public Property Get MaxValue() As Long
    MaxValue = maxVAL
End Property
Public Property Let MaxValue(ByVal nV As Long)
    maxVAL = nV
    RedrawControl Normal
    PropertyChanged "MaxValue"
End Property
'val
Public Property Get Value() As Long
    Value = VAL1
End Property
Public Property Let Value(ByVal nV As Long)
    VAL1 = nV
    RedrawControl Normal
    PropertyChanged "Value"
End Property
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
    
        For i = HG / 2 To HG * 1.5
            Circle (i, HG / 2), HG / 2, BC, (3.14 / 2), ((3 / 2) * 3.14)
        Next i
        
        For i = WD - HG * 1.2 To WD - HG / 2
            Circle (i, HG / 2), HG / 2, BC, (3 / 2) * 3.14, 3.14 / 2
        Next i
        
        Line (HG / 2, 0)-(WD - HG / 2 - 2, HG), BC2, BF
        Line (HG / 2, 0)-(HG / 2 + 5, HG), TC, BF
        
        If VAL1 > 0 Then
            Line (HG / 2 + 5, 0)-((WD - HG / 2 - 2) * (VAL1 / maxVAL), HG), BC1, BF
        End If
        
        lblCAP.ForeColor = FC
        bHovering = False
        MPRESS = False
        
       DoEvents
    'createSkinnedForm UserControl.hWnd, UserControl.BackColor, UserControl.Height, UserControl.Width, StarTrekButton
End Function

Private Sub UserControl_InitProperties()
    BC = &HFF8080
    BC1 = &HFF8080
    BC2 = &HFFC0C0
    FC = vbBlack
    CAP = "StarTrekProgressBar"
    TC = Ambient.BackColor
    
    VAL1 = 0
    maxVAL = 100
    Set FON = Ambient.Font
    
    RedrawControl Normal
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    TC = PropBag.ReadProperty("TransparentColor")
    BC = PropBag.ReadProperty("BackColor")
       
    FC = PropBag.ReadProperty("ForeColor")
    
    CAP = PropBag.ReadProperty("Caption")
    
    BC1 = PropBag.ReadProperty("BackColorScroll")
    BC2 = PropBag.ReadProperty("BackColorScroll2")
    
    maxVAL = PropBag.ReadProperty("MaxValue")
    VAL1 = PropBag.ReadProperty("Value")
    Set FON = PropBag.ReadProperty("Font")
    
    RedrawControl Normal
End Sub

Private Sub UserControl_Resize()
    RedrawControl Normal
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "TransparentColor", TC
    PropBag.WriteProperty "BackColor", BC
    PropBag.WriteProperty "ForeColor", FC
    PropBag.WriteProperty "Caption", CAP
    PropBag.WriteProperty "Font", FON
    PropBag.WriteProperty "BackColorScroll2", BC2
    PropBag.WriteProperty "BackColorScroll", BC1
    
    PropBag.WriteProperty "Value", VAL1
    PropBag.WriteProperty "MaxValue", maxVAL
End Sub

