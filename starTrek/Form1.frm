VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   963
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.StarTrekProgressBarB progBar2 
      Height          =   315
      Left            =   9240
      TabIndex        =   23
      Top             =   1380
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      BorderColor     =   33023
      ScrollColor     =   12640511
      BackColor       =   0
      ScrollFldWidth  =   10
      MaxValue        =   100
      Value           =   0
      DrawWidth       =   1
   End
   Begin Project1.StarTrekProgressBarA progBar1 
      Height          =   315
      Left            =   2400
      TabIndex        =   22
      Top             =   1320
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
      TransparentColor=   0
      BackColor       =   16744576
      ForeColor       =   0
      Caption         =   "StarTrekProgressBar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorScroll2=   16761024
      BackColorScroll =   16744576
      Value           =   0
      MaxValue        =   100
   End
   Begin Project1.StarTrekLabelC StarTrekLabelC1 
      Height          =   405
      Left            =   8880
      TabIndex        =   21
      Top             =   3420
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   714
      TransparentColor=   0
      BackColor       =   16744576
      ForeColor       =   0
      Caption         =   "StarTrekLabelC"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.StarTrekFrameMC_A StarTrekFrameMC_A1 
      Height          =   3975
      Left            =   8100
      TabIndex        =   12
      Top             =   1800
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7011
      Caption         =   "StarTrekFrame"
      BorderColor1    =   33023
      BorderColor2    =   8438015
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentColor=   0
      LineWidth_Left  =   10
      LineWidth_Top   =   20
      LineWidth_Right =   20
      LineWidth_Botton=   15
      Begin Project1.StarTrekLabelB StarTrekLabelB1 
         Height          =   375
         Left            =   540
         TabIndex        =   20
         Top             =   1140
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   16744576
         ForeColor       =   0
         Caption         =   "StarTrekLabelB"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.StarTrekLabelA StarTrekLabelA1 
         Height          =   375
         Left            =   780
         TabIndex        =   19
         Top             =   660
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   16744576
         ForeColor       =   0
         Caption         =   "StarTrekLabelA"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Project1.StarTrekFrame StarTrekFrame1 
      Height          =   3915
      Left            =   1260
      TabIndex        =   11
      Top             =   1800
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   6906
      ForeColor       =   33023
      BorderColor     =   33023
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "StarTrekFrame"
      Begin Project1.StarTrekOptionButton StarTrekOptionButton1 
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   3300
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   16744576
         BackColorHover  =   16761024
         BackColorMouseDown=   16711680
         ForeColor       =   0
         ForeColorHover  =   16777215
         ForeColorMouseDown=   16777215
         Caption         =   "StarTrekOptionButton"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckColor      =   8438015
         Value           =   0   'False
      End
      Begin Project1.StarTrekCheckBoxB StarTrekCheckBoxB1 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   33023
         BackColorHover  =   36863
         BackColorMouseDown=   4210816
         ForeColor       =   0
         ForeColorHover  =   16777215
         ForeColorMouseDown=   16777215
         Caption         =   "StarTrekCheckBox"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checked         =   0   'False
         CheckColor      =   8438015
      End
      Begin Project1.StarTrekCheckBoxA StarTrekCheckBoxA1 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2220
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   33023
         BackColorHover  =   36863
         BackColorMouseDown=   4210816
         ForeColor       =   0
         ForeColorHover  =   16777215
         ForeColorMouseDown=   16777215
         Caption         =   "StarTrekCheckBox"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checked         =   0   'False
         CheckColor      =   8438015
      End
      Begin Project1.StarTrekButtonC StarTrekButtonC1 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1620
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   16744576
         BackColorHover  =   16761024
         BackColorMouseDown=   16711680
         ForeColor       =   0
         ForeColorHover  =   16777215
         ForeColorMouseDown=   16777215
         Caption         =   "StarTrekButton"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.StarTrekButtonB StarTrekButtonB1 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1140
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   16744576
         BackColorHover  =   16761024
         BackColorMouseDown=   16711680
         ForeColor       =   0
         ForeColorHover  =   16777215
         ForeColorMouseDown=   16777215
         Caption         =   "StarTrekButtonB"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.StarTrekButtonA StarTrekButtonA1 
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   16744576
         BackColorHover  =   16761024
         BackColorMouseDown=   16711680
         ForeColor       =   0
         ForeColorHover  =   16777215
         ForeColorMouseDown=   16777215
         Caption         =   "StarTrekButton"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.StarTrekOptionButton StarTrekOptionButton2 
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   3300
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   16744576
         BackColorHover  =   16761024
         BackColorMouseDown=   16711680
         ForeColor       =   0
         ForeColorHover  =   16777215
         ForeColorMouseDown=   16777215
         Caption         =   "StarTrekOptionButton"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckColor      =   8438015
         Value           =   0   'False
      End
      Begin Project1.StarTrekCheckBoxB StarTrekCheckBoxB2 
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   2760
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   33023
         BackColorHover  =   36863
         BackColorMouseDown=   4210816
         ForeColor       =   0
         ForeColorHover  =   16777215
         ForeColorMouseDown=   16777215
         Caption         =   "StarTrekCheckBox"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checked         =   0   'False
         CheckColor      =   8438015
      End
      Begin Project1.StarTrekCheckBoxA StarTrekCheckBoxA2 
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Top             =   2220
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   661
         TransparentColor=   0
         BackColor       =   33023
         BackColorHover  =   36863
         BackColorMouseDown=   4210816
         ForeColor       =   0
         ForeColorHover  =   16777215
         ForeColorMouseDown=   16777215
         Caption         =   "StarTrekCheckBox"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checked         =   0   'False
         CheckColor      =   8438015
      End
   End
   Begin Project1.StarTrekCornerButtonA StarTrekCornerButtonA1 
      Height          =   795
      Left            =   60
      TabIndex        =   10
      Top             =   900
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1402
      TransparentColor=   0
      BackColor       =   16744576
      BackColorHover  =   16761024
      BackColorMouseDown=   16711680
      ForeColor       =   0
      ForeColorHover  =   16777215
      ForeColorMouseDown=   16777215
      Caption         =   "About"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LineWidth1      =   60
      LineWidth2      =   10
   End
   Begin Project1.StarTrekCornerButtonB StarTrekCornerButtonB1 
      Height          =   855
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1508
      TransparentColor=   0
      BackColor       =   16744576
      BackColorHover  =   16761024
      BackColorMouseDown=   16711680
      ForeColor       =   0
      ForeColorHover  =   16777215
      ForeColorMouseDown=   16777215
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LineWidth1      =   60
      LineWidth2      =   10
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7770
      Top             =   90
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080C0FF&
      Height          =   3570
      Left            =   60
      TabIndex        =   8
      Top             =   3105
      Width           =   915
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FFFF&
      Height          =   360
      Left            =   60
      TabIndex        =   7
      Top             =   2715
      Width           =   915
   End
   Begin VB.Label Label7 
      BackColor       =   &H008083E0&
      Height          =   930
      Left            =   60
      TabIndex        =   6
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Height          =   150
      Left            =   10440
      TabIndex        =   5
      Top             =   900
      Width           =   4110
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Height          =   150
      Left            =   6990
      TabIndex        =   4
      Top             =   900
      Width           =   3420
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404080&
      Height          =   150
      Left            =   4320
      TabIndex        =   3
      Top             =   900
      Width           =   2640
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404080&
      Height          =   150
      Left            =   4320
      TabIndex        =   2
      Top             =   705
      Width           =   10230
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Height          =   150
      Left            =   4080
      TabIndex        =   1
      Top             =   900
      Width           =   180
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Height          =   150
      Left            =   4080
      TabIndex        =   0
      Top             =   705
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----- about -----------------------------------------------
'   Star Trek: ControlSet
'   author: ivan štimac
'           ivan.stimac@po.htnet.hr, flashboy01@gmail.com
'
'   price: free for any use
'------------------------------------------------------------
Dim RNDNum As Integer

Private Sub StarTrekCornerButtonA1_Click()
    Dim msg As String
    
    msg = "Star Trek: ControlSet v. 1.0.0." & vbCrLf & _
          "By: Ivan Stimac"
    MsgBox msg, vbInformation, "Star Trek: ControlSet"
End Sub

Private Sub StarTrekCornerButtonB1_Click()
    End
End Sub

Private Sub Timer1_Timer()
    Randomize
    RNDNum = Int(Rnd * 100)
    progBar1.Value = RNDNum
    
    Randomize
    RNDNum = Int(Rnd * 100)
    progBar2.Value = RNDNum
End Sub
