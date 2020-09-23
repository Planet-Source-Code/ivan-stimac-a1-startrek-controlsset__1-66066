VERSION 5.00
Begin VB.UserControl StarTrekLabelA 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "StarTrekLabel.ctx":0000
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
Attribute VB_Name = "StarTrekLabelA"
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

Private bHovering, ENBL, MPRESS As Boolean
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


'back color
Public Property Get BackColor() As OLE_COLOR
    BackColor = BC
End Property
Public Property Let BackColor(ByVal nV As OLE_COLOR)
    BC = nV
    RedrawControl Normal
    PropertyChanged "BackColor"
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
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
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
    
    lblCAP.Left = WD / 2 - lblCAP.Width / 2
    lblCAP.Top = HG / 2 - lblCAP.Height / 2
    
        For i = HG / 2 To HG * 1.5
            Circle (i, HG / 2), HG / 2, BC, (3.14 / 2), ((3 / 2) * 3.14)
        Next i
        
        For i = WD - HG * 1.2 To WD - HG / 2
            Circle (i, HG / 2), HG / 2, BC, (3 / 2) * 3.14, 3.14 / 2
        Next i
        
        Line (HG / 2, 0)-(WD - HG / 2, HG), BC, BF
        'Line (HG / 2, 0)-(HG / 2 + 5, HG), TC, BF
        
        lblCAP.ForeColor = FC
        bHovering = False
        MPRESS = False
    DoEvents
    'createSkinnedForm UserControl.hWnd, UserControl.BackColor, UserControl.Height, UserControl.Width, StarTrekButton
End Function

Private Sub UserControl_InitProperties()
    BC = &HFF8080
    FC = vbBlack
    CAP = "StarTrekLabelA"
    TC = Ambient.BackColor
    Set FON = Ambient.Font
    
    RedrawControl Normal
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    TC = PropBag.ReadProperty("TransparentColor")
    BC = PropBag.ReadProperty("BackColor")
    
    FC = PropBag.ReadProperty("ForeColor")
    
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
    PropBag.WriteProperty "ForeColor", FC
    PropBag.WriteProperty "Caption", CAP
    PropBag.WriteProperty "Font", FON
End Sub

