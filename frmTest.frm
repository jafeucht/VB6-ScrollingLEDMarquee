VERSION 5.00
Object = "*\AMarquee.vbp"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin MarqueeCtl.Marquee Marquee4 
      Height          =   4515
      Left            =   7440
      TabIndex        =   3
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   7964
      BackColor       =   -2147483633
      Orientation     =   1
      LEDSize         =   2
      Interval        =   1
   End
   Begin MarqueeCtl.Marquee Marquee3 
      Height          =   4515
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   7964
      Orientation     =   1
      LEDSize         =   1
      LEDSpacing      =   2
      OffSet          =   1
      Text            =   "Vertical"
   End
   Begin MarqueeCtl.Marquee Marquee2 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1720
      BackColor       =   -2147483633
      Text            =   "-- Horizontal --"
   End
   Begin MarqueeCtl.Marquee Marquee1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   873
      RowCount        =   100
      LEDSize         =   1
      OffSet          =   5
      Text            =   "This is a test!!!"
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer
    ' Load all 256 characters.
    For i = 0 To 255
        Marquee4.Text = Marquee4.Text & Chr(i)
    Next i
End Sub
