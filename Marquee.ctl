VERSION 5.00
Begin VB.UserControl Marquee 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   FillStyle       =   0  'Solid
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   Begin VB.Timer tmTimer 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Marquee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim Characters(0 To 255) As String

Const CHAR_HEIGHT  As Integer = 8
Const CHAR_WIDTH As Integer = 7

Enum EnumOrientation
    Horizontal
    Vertical
End Enum

Dim pOrientation As Integer
Dim pRowCount As Integer
Dim pLEDSize As Integer
Dim pLEDOnColor As Long
Dim pLEDOffColor As Long
Dim pLEDSpacing As Integer
Dim pOffSet As Integer
Dim pText As String
Dim pEnabled As Boolean

Dim WaitStatus As Integer
Dim CharIn As Integer
Dim CurChar As Integer
Dim CurTextPos As Integer
Dim NewRow As String * CHAR_HEIGHT

Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Property Let BackColor(NewColor As OLE_COLOR)
    UserControl.BackColor = NewColor
    PropertyChanged "BackColor"
End Property

Property Get Orientation() As EnumOrientation
    Orientation = pOrientation
End Property

Property Get RowCount() As Integer
    RowCount = pRowCount
End Property

Property Let RowCount(NewValue As Integer)
Dim i As Integer
    pRowCount = NewValue
    PropertyChanged "RowCount"
    For i = 1 To CHAR_HEIGHT
        Characters(i) = String$(pRowCount, "0")
    Next i
    UserControl_Paint
End Property

Property Let Orientation(NewValue As EnumOrientation)
    pOrientation = NewValue
    UserControl_Paint
    PropertyChanged "Orientation"
End Property

Property Get LEDSize() As Integer
    LEDSize = pLEDSize
End Property

Property Let LEDSize(NewValue As Integer)
    pLEDSize = NewValue
    UserControl_Paint
    PropertyChanged "LEDSize"
End Property

Property Get LEDOnColor() As OLE_COLOR
    LEDOnColor = pLEDOnColor
End Property

Property Let LEDOnColor(NewColor As OLE_COLOR)
    pLEDOnColor = NewColor
    PropertyChanged "LEDOnColor"
End Property

Property Get LEDOffColor() As OLE_COLOR
    LEDOffColor = pLEDOffColor
End Property

Property Let LEDOffColor(NewColor As OLE_COLOR)
    pLEDOffColor = NewColor
    PropertyChanged "LEDOffColor"
End Property

Property Get LEDSpacing() As Integer
    LEDSpacing = pLEDSpacing
End Property

Property Let LEDSpacing(NewValue As Integer)
    pLEDSpacing = NewValue
    UserControl_Paint
    PropertyChanged "LEDSpacing"
End Property

Property Get Interval() As Integer
    Interval = tmTimer.Interval
End Property

Property Let Interval(NewValue As Integer)
    tmTimer.Interval = NewValue
    PropertyChanged "Interval"
End Property

Property Get OffSet() As Integer
    OffSet = pOffSet
End Property

Property Let OffSet(NewValue As Integer)
    pOffSet = NewValue
    PropertyChanged "OffSet"
End Property

Property Get Text() As String
    Text = pText
End Property

Property Let Text(NewValue As String)
    CurChar = 0
    pText = NewValue
    PropertyChanged "Text"
End Property

Property Get Enabled() As OLE_CANCELBOOL
    Enabled = pEnabled
End Property

Property Let Enabled(NewValue As OLE_CANCELBOOL)
    pEnabled = NewValue
    If Ambient.UserMode = True Then tmTimer.Enabled = NewValue
    PropertyChanged "Enabled"
End Property

' Algorithm that operates the whole dang thing
Sub Animate()
Static Wait As Boolean
Dim i As Integer
    If Wait Then
        WaitStatus = WaitStatus + 1
        If WaitStatus >= pOffSet Then
            CharIn = CHAR_WIDTH - 1
            WaitStatus = 0
            CurTextPos = 0
            Wait = False
        End If
    End If
    If Not Wait And Len(Text) > 0 Then
        CharIn = CharIn + 1
        If CharIn >= CHAR_WIDTH Then
            CharIn = 1
            CurTextPos = CurTextPos + 1
            If CurTextPos > Len(Text) Then
                Wait = True
            Else
                CurChar = Asc(Mid(Text, CurTextPos, 1))
            End If
        End If
    End If
    For i = 1 To CHAR_HEIGHT
        Mid(NewRow, i, 1) = GetChar(CurChar, i, CharIn)
    Next i
    PrintData
End Sub

Sub PrintData()
Dim i As Integer, y As Integer
Dim DrawCol As Long, BorderSpace As Integer
    BorderSpace = pLEDSize + pLEDSpacing
    Picture = Image
    Select Case pOrientation
        Case Horizontal
            UserControl.PaintPicture UserControl.Picture, 1, 1, ScaleWidth, ScaleHeight, (pLEDSize + pLEDSpacing) * 2 + 1, 1, ScaleWidth, ScaleHeight, vbSrcCopy
        Case Vertical
            UserControl.PaintPicture UserControl.Picture, 1, 1, ScaleWidth, ScaleHeight, 1, -(pLEDSize + pLEDSpacing) * 2 + 1, ScaleWidth, ScaleHeight, vbSrcCopy
    End Select
    For y = 0 To CHAR_HEIGHT - 1
        PaintLED pRowCount - 1, y, Mid$(NewRow, y + 1, 1) = 1
    Next y
    DoEvents
End Sub

Sub PaintLED(x As Integer, y As Integer, LEDOn As Boolean)
Dim DrawCol As Long, XLen As Integer, YLen As Integer
Dim BorderSpace As Integer
    BorderSpace = pLEDSize + pLEDSpacing
    DrawCol = IIf(LEDOn, pLEDOnColor, pLEDOffColor)
    UserControl.FillColor = DrawCol
    XLen = BorderSpace + x * BorderSpace * 2
    YLen = BorderSpace + y * BorderSpace * 2
    Select Case pOrientation
        Case Horizontal
            UserControl.Circle (XLen, YLen), pLEDSize, DrawCol
        Case Vertical
            UserControl.Circle (YLen, ScaleHeight - XLen), pLEDSize, DrawCol
    End Select
End Sub

' Blanks out all LEDs on the form
Sub LEDReset()
Dim x As Integer, y As Integer
    For y = 0 To CHAR_HEIGHT - 1
        For x = 0 To pRowCount - 1
            PaintLED x, y, False
        Next x
    Next y
End Sub

' Retrieves a part of a character
Function GetChar(Index As Integer, Row As Integer, Col As Integer) As Integer
    If Col > CHAR_WIDTH Or Row > CHAR_HEIGHT Then Exit Function
    GetChar = Mid(Characters(Index), Col + CHAR_WIDTH * (Row - 1), 1)
End Function

Private Sub tmTimer_Timer()
    Animate
End Sub

Private Sub UserControl_Initialize()
Dim i As Integer
    For i = 0 To 255
        Characters(i) = LoadResString(i + 1)
    Next i
    tmTimer.Interval = 100
    pLEDOffColor = &HFF0000   ' Blue
    pLEDOnColor = &HFF& ' Red
    pLEDSpacing = 1
    pLEDSize = 3
    pOffSet = 100
    pRowCount = 50
    pEnabled = True
End Sub

Private Sub UserControl_Paint()
Dim BorderSpace As Integer, XLen As Long, YLen As Long
Static InPaint As Boolean
    If InPaint Then Exit Sub
    InPaint = True
    BorderSpace = pLEDSize + pLEDSpacing
    XLen = pRowCount * BorderSpace * 2 + 1
    YLen = CHAR_HEIGHT * BorderSpace * 2 + 1
    Select Case pOrientation
        Case Horizontal
            UserControl.Width = ScaleX(XLen, 3, 1)
            UserControl.Height = ScaleY(YLen, 3, 1)
        Case Vertical
            UserControl.Width = ScaleX(YLen, 3, 1)
            UserControl.Height = ScaleY(XLen, 3, 1)
    End Select
    InPaint = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Integer
    UserControl.BackColor = PropBag.ReadProperty("BackColor", 0)
    pRowCount = PropBag.ReadProperty("RowCount", 50)
    pOrientation = PropBag.ReadProperty("Orientation", Horizontal)
    pLEDSize = PropBag.ReadProperty("LEDSize", 3)
    pLEDOnColor = PropBag.ReadProperty("LEDOnColor", &HFF&)
    pLEDOffColor = PropBag.ReadProperty("LEDOffColor", &HFF0000)
    pLEDSpacing = PropBag.ReadProperty("LEDSpacing", 1)
    tmTimer.Interval = PropBag.ReadProperty("Interval", 100)
    pOffSet = PropBag.ReadProperty("OffSet", 100)
    pText = PropBag.ReadProperty("Text", "")
    pEnabled = PropBag.ReadProperty("Enabled", True)
    If Ambient.UserMode Then tmTimer.Enabled = pEnabled
    UserControl_Paint
    LEDReset
End Sub

Private Sub UserControl_Resize()
    UserControl_Paint
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", UserControl.BackColor, 0
    PropBag.WriteProperty "RowCount", pRowCount, 50
    PropBag.WriteProperty "Orientation", pOrientation, Horizontal
    PropBag.WriteProperty "LEDSize", pLEDSize, 3
    PropBag.WriteProperty "LEDOnColor", pLEDOnColor, &HFF&
    PropBag.WriteProperty "LEDOffColor", pLEDOffColor, &HFF0000
    PropBag.WriteProperty "LEDSpacing", pLEDSpacing, 1
    PropBag.WriteProperty "Interval", tmTimer.Interval, 100
    PropBag.WriteProperty "OffSet", pOffSet, 100
    PropBag.WriteProperty "Text", pText, ""
    PropBag.WriteProperty "Enabled", pEnabled, True
End Sub
