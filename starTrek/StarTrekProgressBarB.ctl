VERSION 5.00
Begin VB.UserControl StarTrekProgressBarB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   ToolboxBitmap   =   "StarTrekProgressBarB.ctx":0000
End
Attribute VB_Name = "StarTrekProgressBarB"
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

Private BC, TC, SC, BDRC As OLE_COLOR
Private SCW, i, DW, VAL1, maxVAL As Integer
Private PRCNT As Single
Private WD, HG, k, LNT As Long

'draw width
Public Property Get DrawWidth() As Integer
    DrawWidth = DW
End Property
Public Property Let DrawWidth(ByVal nV As Integer)
    If nV < 1 Or nV > 20 Then
        MsgBox "Nevaljan unos! Ova varijabla može biti u iznosu od 1 do 20!", vbExclamation
        Exit Property
    End If
    DW = nV
    RedrawControl
    PropertyChanged "DrawWidth"
End Property
'scrol width
Public Property Get ScrollFldWidth() As Integer
    ScrollFldWidth = SCW
End Property
Public Property Let ScrollFldWidth(ByVal nV As Integer)
    If nV < 1 Then
        MsgBox "Nevaljan unos! Najmanja vrijednost ove varijeble je 1!", vbExclamation
        Exit Property
    End If
    SCW = nV
    RedrawControl
    PropertyChanged "ScrollFldWidth"
End Property
'val
Public Property Get Value() As Integer
    Value = VAL1
End Property
Public Property Let Value(ByVal nV As Integer)
    VAL1 = nV
    RedrawControl
    PropertyChanged "Value"
End Property
'max val
Public Property Get MaxValue() As Integer
    MaxValue = maxVAL
End Property
Public Property Let MaxValue(ByVal nV As Integer)
    If nV < VAL1 Then
        MsgBox "Nevaljan unos! Ova varijabla mora biti veæa od varijable Value!", vbExclamation
        Exit Property
    End If
    maxVAL = nV
    RedrawControl
    PropertyChanged "MaxValue"
End Property
'transparent color
Public Property Get BackColor() As OLE_COLOR
    BackColor = TC
End Property
Public Property Let BackColor(ByVal nV As OLE_COLOR)
    TC = nV
    RedrawControl
    PropertyChanged "BackColor"
End Property
'border color
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = BDRC
End Property
Public Property Let BorderColor(ByVal nV As OLE_COLOR)
    BDRC = nV
    RedrawControl
    PropertyChanged "BorderColor"
End Property
'scrol color
Public Property Get ScrollColor() As OLE_COLOR
    ScrollColor = SC
End Property
Public Property Let ScrollColor(ByVal nV As OLE_COLOR)
    SC = nV
    RedrawControl
    PropertyChanged "ScrollColor"
End Property

Private Function RedrawControl()
    Cls
    If DW <= 0 Then DW = 1
    If maxVAL = 0 Then maxVAL = 100
    If SCW = 0 Then SCW = 10
    
    UserControl.DrawWidth = DW
    UserControl.BackColor = TC
    k = 0
    LNT = 0
    WD = Width / Screen.TwipsPerPixelX - UserControl.DrawWidth
    HG = Height / Screen.TwipsPerPixelX - UserControl.DrawWidth
    
    PRCNT = VAL1 / maxVAL
    Line (0, 0)-((WD * PRCNT), HG), SC, BF
    
    Line (0, 0)-(WD, HG), BDRC, B
    Do While LNT <= WD
        k = k + 1
        LNT = SCW * k
        Line (LNT, 0)-(LNT, HG), BDRC
    Loop
End Function


Private Sub UserControl_InitProperties()
    maxVAL = 100
    VAL1 = 0
    TC = Ambient.BackColor
    BDRC = &H80FF&
    SCW = 10
    SC = &HC0E0FF
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BDRC = PropBag.ReadProperty("BorderColor")
    SC = PropBag.ReadProperty("ScrollColor")
    TC = PropBag.ReadProperty("BackColor")
    SCW = PropBag.ReadProperty("ScrollFldWidth")
    
    maxVAL = PropBag.ReadProperty("MaxValue")
    VAL1 = PropBag.ReadProperty("Value")
    DW = PropBag.ReadProperty("DrawWidth")
    
    DoEvents
    
    RedrawControl
End Sub

Private Sub UserControl_Resize()
    RedrawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BorderColor", BDRC
    PropBag.WriteProperty "ScrollColor", SC
    PropBag.WriteProperty "BackColor", TC
    PropBag.WriteProperty "ScrollFldWidth", SCW
    PropBag.WriteProperty "MaxValue", maxVAL
    PropBag.WriteProperty "Value", VAL1
    PropBag.WriteProperty "DrawWidth", DW
End Sub
