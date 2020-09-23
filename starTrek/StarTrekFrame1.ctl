VERSION 5.00
Begin VB.UserControl StarTrekFrame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5100
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   340
   ToolboxBitmap   =   "StarTrekFrame1.ctx":0000
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   105
      Left            =   0
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   105
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   135
      Left            =   0
      Top             =   180
      Width           =   390
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   120
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   300
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   450
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
End
Attribute VB_Name = "StarTrekFrame"
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

Private i, a, b As Long
Private k, xkv, x, c, kx, X1, Y2 As Double
Private FC, BC, BCC As OLE_COLOR
Private CAP As String
Private FON As StdFont


Const PI = 3.1415
Dim RAD As Single
'caption
Public Property Get Caption() As String
    Caption = CAP
End Property
Public Property Let Caption(ByVal nC As String)
    CAP = nC
    RedrawControl
    PropertyChanged "Caption"
End Property
'back color
Public Property Get BackColor() As OLE_COLOR
    BackColor = BCC
End Property
Public Property Let BackColor(ByVal nV As OLE_COLOR)
    BCC = nV
    RedrawControl
    PropertyChanged "BackColor"
End Property
'BorderColor
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = BC
End Property
Public Property Let BorderColor(ByVal nV As OLE_COLOR)
    BC = nV
    RedrawControl
    PropertyChanged "BorderColor"
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
'font
Public Property Get Font() As StdFont
    Set Font = FON
End Property
Public Property Set Font(ByVal nF As StdFont)
    On Error Resume Next
    Set FON = nF
    RedrawControl
    PropertyChanged "Font"
End Property


Private Sub UserControl_Initialize()
    RedrawControl
End Sub


Public Function RedrawControl()
On Error Resume Next
    UserControl.Cls
    Shape2.BackColor = BC
    Shape1.BackColor = BC
    Shape3.BackColor = BC
    Shape4.BackColor = BC
    
    lblCAP.Caption = CAP
    UserControl.BackColor = BCC
    Set lblCAP.Font = FON
    lblCAP.ForeColor = FC
    UserControl.Refresh
    
    For i = 0 To 20
        Line (lblCAP.Left + lblCAP.Width + 5, i)-(Width, i), BC
    Next i
    
    For i = 0 To 4
        Line (i, 10)-(i, (Height - Shape2.Height) / Screen.TwipsPerPixelY - 2), BC
    Next i
    
    
    Shape2.Top = (Height - Shape2.Height) / Screen.TwipsPerPixelY - 7
    lblCAP.Top = 10 - lblCAP.Height / 2
    
    
    For i = Shape2.Top To Shape2.Top + 10
        Line (Shape2.Height / 2, i)-(UserControl.Width, i), BC
    Next i
   ' MsgBox Shape2.Top
    
   ' For i = 10 To lblCAP.Left - 5
     '   Line (i, 0)-(i, 20)
   ' Next i
End Function

Private Function DToR(ByVal x As Long)
    RAD = PI / 180
    DToR = x * RAD
End Function

Private Sub UserControl_InitProperties()
    FC = &H80FF&
    BC = &H80FF&
    BCC = vbBlack
    CAP = "Caption"
    Set FON = Ambient.Font
    'Set FON.Bold = True
    'Set FON.Name = "Verdana"
   ' Set FON.Size = 9
    RedrawControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    FC = PropBag.ReadProperty("ForeColor")
    BC = PropBag.ReadProperty("BorderColor")
    BCC = PropBag.ReadProperty("BackColor")
    Set FON = PropBag.ReadProperty("Font")
    CAP = PropBag.ReadProperty("Caption")

End Sub

Private Sub UserControl_Resize()
    RedrawControl
End Sub

Private Sub UserControl_Show()
    RedrawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ForeColor", FC
    PropBag.WriteProperty "BorderColor", BC
    PropBag.WriteProperty "BackColor", BCC
    PropBag.WriteProperty "Font", FON
    PropBag.WriteProperty "Caption", CAP

End Sub
