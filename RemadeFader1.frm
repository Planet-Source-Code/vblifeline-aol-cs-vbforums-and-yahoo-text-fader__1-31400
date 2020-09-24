VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C07800&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   3735
   ClientTop       =   2730
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RtBox 
      Height          =   1215
      Left            =   150
      TabIndex        =   30
      Top             =   5475
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2143
      _Version        =   393217
      BackColor       =   12613632
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"RemadeFader1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox TempLate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5175
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   29
      Top             =   75
      Width           =   375
   End
   Begin VB.CheckBox Check4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Italic"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1575
      TabIndex        =   24
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this option to make the faded text Italic"
      Top             =   1350
      Width           =   690
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Bold"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   825
      TabIndex        =   26
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this option to make the faded text bold"
      Top             =   1350
      Width           =   690
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "UnderLine"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2325
      TabIndex        =   25
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this option to make the faded text Underlined"
      Top             =   1350
      Width           =   1065
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Wavy"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   75
      TabIndex        =   27
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click This Option To Make the Text Come Out Faded"
      Top             =   1350
      Width           =   765
   End
   Begin VB.Timer FadeTxt 
      Interval        =   100
      Left            =   2250
      Tag             =   "made by leo/jason"
      Top             =   3150
   End
   Begin VB.TextBox tCode 
      Appearance      =   0  'Flat
      Height          =   2655
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Tag             =   "t"
      Text            =   "RemadeFader1.frx":011A
      Top             =   1800
      Width           =   5055
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   4725
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   16
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   4365
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   15
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   4005
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   14
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   3645
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   13
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   3285
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   12
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   2925
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   11
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   2565
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   10
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   2205
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   9
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   1845
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   8
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   1485
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   7
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   1125
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   6
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   765
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   5
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.TextBox TxtToFade 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   45
      TabIndex        =   4
      Tag             =   "t"
      Text            =   "Text To Fade"
      Top             =   525
      Width           =   5055
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   405
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   3
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.PictureBox Color 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   45
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   0
      Tag             =   "cb"
      Top             =   930
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   1365
      Left            =   75
      Top             =   5400
      Width           =   4965
   End
   Begin VB.Label ButtonA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Fade Job"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   75
      TabIndex        =   28
      Top             =   4950
      Width           =   2565
   End
   Begin VB.Label ButtonA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   4725
      TabIndex        =   23
      Top             =   75
      Width           =   390
   End
   Begin VB.Label ButtonA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   3435
      TabIndex        =   22
      Top             =   1350
      Width           =   1665
   End
   Begin VB.Label ButtonA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   4350
      TabIndex        =   21
      Top             =   75
      Width           =   390
   End
   Begin VB.Label ButtonA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copy Code To Clip Board"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   75
      TabIndex        =   20
      Top             =   4500
      Width           =   2865
   End
   Begin VB.Label Move 
      BackStyle       =   0  'Transparent
      Height          =   510
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LeoText Fader V2.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   45
      TabIndex        =   18
      Top             =   75
      Width           =   4380
   End
   Begin VB.Label ButtonA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Color Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2700
      TabIndex        =   2
      Top             =   4950
      Width           =   2340
   End
   Begin VB.Label ButtonA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fade Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   4500
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private MouseDownForm
Private MouseDownFormX
Private MouseDownFormY
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private FadeTT(43)

Private Sub Command1_Click()
If Check2.Value = 1 Then
RtBox.SelBold = True
Else
RtBox.SelBold = False
End If
End Sub

Private Sub FadeTxt_Timer()
DoEvents
FadeTT(0) = FadeTT(0) + 1
If FadeTT(0) = 44 Then
FadeTT(0) = 1
End If
Title.ForeColor = FadeTT(FadeTT(0))
Form1.caption = FadeTT(FadeTT(0))
App.Title = FadeTT(FadeTT(0))
End Sub
Private Sub Form_Load()
ReloadInterface
FadeTT(0) = 0
FadeTT(1) = "&HFF0000"
FadeTT(2) = "&HFF1212"
FadeTT(3) = "&HFF2424"
FadeTT(4) = "&HFF3636"
FadeTT(5) = "&HFF4848"
FadeTT(6) = "&HFF5A5A"
FadeTT(7) = "&HFF6C6C"
FadeTT(8) = "&HFF7E7E"
FadeTT(9) = "&HFF9090"
FadeTT(10) = "&HFFA2A2"
FadeTT(11) = "&HFFB4B4"
FadeTT(12) = "&HFFC6C6"
FadeTT(13) = "&HFFD8D8"
FadeTT(14) = "&HFFEAEA"
FadeTT(15) = "&HFFFCFC"
FadeTT(16) = "&HEDFFFF"
FadeTT(17) = "&HDBEDFF"
FadeTT(18) = "&HC9DBFF"
FadeTT(19) = "&HB7C9FF"
FadeTT(20) = "&HA5B7FF"
FadeTT(21) = "&H93A5FF"
FadeTT(22) = "&H8193FF"
FadeTT(23) = "&H6F81FF"
FadeTT(24) = "&H5D6FFF"
FadeTT(25) = "&H4B5DFF"
FadeTT(26) = "&H394BFF"
FadeTT(27) = "&H2739FF"
FadeTT(28) = "&H1527FF"
FadeTT(29) = "&H0315FF"
FadeTT(30) = "&H0003FF"
FadeTT(31) = "&H1200ED"
FadeTT(32) = "&H2400DB"
FadeTT(33) = "&H3600C9"
FadeTT(34) = "&H4800B7"
FadeTT(35) = "&H5A00A5"
FadeTT(36) = "&H6C0093"
FadeTT(37) = "&H7E0081"
FadeTT(38) = "&H90006F"
FadeTT(39) = "&HA2005D"
FadeTT(40) = "&HB4004B"
FadeTT(41) = "&HC60039"
FadeTT(42) = "&HD80027"
FadeTT(43) = "&HEA0015"
End Sub
Private Sub move_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDownForm = 1
MouseDownFormX = X
MouseDownFormY = Y
End Sub
Private Sub move_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDownForm = 0
End Sub
Private Sub move_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseDownForm <> 1 Then
Exit Sub
End If
Dim Z As POINTAPI
Call GetCursorPos(Z)
Form1.Top = (Z.Y * 15) - MouseDownFormY
Form1.Left = (Z.X * 15) - MouseDownFormX
End Sub
Private Sub ButtonA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonA(Index).BackColor = MdBgColor
ButtonA(Index).ForeColor = MdFgColor
End Sub
Private Sub ButtonA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonA(Index).BackColor = BGcolor
ButtonA(Index).ForeColor = FgColor
End Sub
Private Sub ButtonA_Click(Index As Integer)
On Error GoTo Error
If Index = 0 Then
' Start of setting colors into an array
Dim fyi(10000)
tcount = 0
For X = 0 To Color.Count - 1
DoEvents
If Color(X).Tag = "1" Then
fyi(tcount) = Color(X).BackColor
tcount = tcount + 1
End If
Next
If tcount < 2 Then
MyMSG "Need More Colors", "You Need At Lest To Colors"
Exit Sub
End If
' End of making array
If Form3.Check3.Value = 1 Then
Fade_Text TxtToFade.text, fyi, False
Else
Fade_Text TxtToFade.text, fyi, True
End If
tCode.text = ""
If Check1.Value = 1 Then
ntempb = 0
RtBox.SelText = ""
RtBox.text = ""
For X = 1 To Len(TxtToFade.text)
DoEvents
ntempb = ntempb + 1
If ntempb = 5 Then
ntempb = 1
End If
If ntempb = 1 Then
ntempc = "<sup>"
End If
If ntempb = 2 Then
ntempc = "</sup>"
End If
If ntempb = 3 Then
ntempc = "<sub>"
End If
If ntempb = 4 Then
ntempc = "</sub>"
End If
' code for preview
RtBox.Font.Name = Form3.Text15.text
RtBox.SelFontSize = Form3.Text5.text * 2
If Check2.Value = 1 Then
RtBox.SelBold = True
Else
RtBox.SelBold = False
End If
If Check3.Value = 1 Then
RtBox.SelUnderline = True
Else
RtBox.SelUnderline = False
End If
If Check4.Value = 1 Then
RtBox.SelItalic = True
Else
RtBox.SelItalic = False
End If
If cHTML = True Then
RtBox.SelColor = "&H" & (Right(tColor(X), 2) & Right(Left(tColor(X), 4), 2) & Left(tColor(X), 2))
Else
RtBox.SelColor = "&H" & tColor(X)
End If
RtBox.SelText = RtBox.SelText & Right(Left(TxtToFade.text, X), 1)
' end of code for preview
RtBox.SelFontName = Form3.Text15.text
RtBox.SelFontSize = Form3.Text5.text * 2
tCode.text = tCode.text & Form3.Text1.text & tColor(X) & Form3.Text2.text & ntempc & Right(Left(TxtToFade.text, X), 1) & Form3.Text3.text
Next
Else
RtBox.SelText = ""
RtBox.text = ""
For X = 1 To Len(TxtToFade.text)
DoEvents
' code for preview
RtBox.Font.Name = Form3.Text15.text
RtBox.SelFontSize = Form3.Text5.text * 2
If Check2.Value = 1 Then
RtBox.SelBold = True
Else
RtBox.SelBold = False
End If
If Check3.Value = 1 Then
RtBox.SelUnderline = True
Else
RtBox.SelUnderline = False
End If
If Check4.Value = 1 Then
RtBox.SelItalic = True
Else
RtBox.SelItalic = False
End If
If cHTML = True Then
RtBox.SelColor = "&H" & (Right(tColor(X), 2) & Right(Left(tColor(X), 4), 2) & Left(tColor(X), 2))
Else
RtBox.SelColor = "&H" & tColor(X)
End If
RtBox.SelText = RtBox.SelText & Right(Left(TxtToFade.text, X), 1)
' end of code for preview
tCode.text = tCode.text & Form3.Text1.text & tColor(X) & Form3.Text2.text & Right(Left(TxtToFade.text, X), 1) & Form3.Text3.text
Next
End If
If Check3.Value = 1 Then
tCode.text = "<u>" & tCode.text & "</u>"
End If
If Check4.Value = 1 Then
tCode.text = "<i>" & tCode.text & "</i>"
End If
If Check2.Value = 1 Then
tCode.text = "<b>" & tCode.text & "</b>"
End If
If Form3.Check12.Value = 1 Then
tCode.text = Form3.Text12.text & Form3.Text15.text & Form3.Text13.text & tCode.text & Form3.Text14.text
End If
tCode.text = Form3.Text11.text & Form3.Text5.text & Form3.Text10.text & tCode.text & Form3.Text9.text
If Form3.Check1.Value = 1 Then
tCode.text = Form3.Text6.text & Form3.Text4.text & Form3.Text7.text & tCode.text & Form3.Text8.text
End If
End If

If Index = 1 Then
Form4.Show
End If

If Index = 2 Then
tCode.SetFocus
tCode.SelStart = 0
tCode.SelLength = Len(tCode.text)
SendKeys "^c"
End If

If Index = 3 Then
Form1.WindowState = 1
End If

If Index = 4 Then
Form3.Show
End If

If Index = 5 Then
End
End If

If Index = 6 Then
For X = 0 To Color.Count - 1
Color(X).Tag = "0"
Color(X).BackColor = TempLate.BackColor
TxtToFade.text = "Text To Fade"
Next
End If

Exit Sub
Error:
ErrorHandler
End Sub
Private Sub Color_Click(Index As Integer)
nTemp = CustomColor
If nTemp <> "N/A" Then
Color(Index).BackColor = nTemp
Color(Index).Tag = "1"
End If
End Sub
Public Sub Delay(HowLong As Date)
TempTime = DateAdd("s", HowLong, Now)
While TempTime > Now
DoEvents 'Allows windows to handle other stuff
Wend
End Sub

