VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C07800&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   LinkTopic       =   "Form4"
   ScaleHeight     =   6945
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Blue And Orage"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   8
      Left            =   3375
      TabIndex        =   31
      ToolTipText     =   "Color that remind me of golf"
      Top             =   4800
      Width           =   1590
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "PasTells"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   7
      Left            =   3375
      TabIndex        =   30
      ToolTipText     =   "Color that remind me of golf"
      Top             =   4500
      Width           =   1590
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Crazy Colors"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   6
      Left            =   3375
      TabIndex        =   29
      ToolTipText     =   "Color that remind me of golf"
      Top             =   4200
      Width           =   1590
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Green"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   1725
      TabIndex        =   28
      ToolTipText     =   "Color that remind me of golf"
      Top             =   4800
      Width           =   1590
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   0
      Left            =   2250
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   13
      Tag             =   "fb"
      ToolTipText     =   "This is the color that will be the program text color"
      Top             =   450
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   1050
      TabIndex        =   12
      Tag             =   "fb"
      Text            =   "Preview Text Box"
      ToolTipText     =   "This is a text box preview"
      Top             =   5925
      Width           =   4590
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   150
      ScaleHeight     =   735
      ScaleWidth      =   810
      TabIndex        =   11
      Tag             =   "fb"
      ToolTipText     =   "Preview of a color box"
      Top             =   5925
      Width           =   840
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   2250
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   10
      Tag             =   "fb"
      ToolTipText     =   "This is the color that will be the program Back Groung Color"
      Top             =   2250
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   2
      Left            =   2250
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   9
      Tag             =   "fb"
      ToolTipText     =   "this color will be used as the backround color of text boxes"
      Top             =   1050
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   3
      Left            =   5025
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   8
      Tag             =   "fb"
      ToolTipText     =   "This color will be the color of the text boxes text color... Cunfusing Huhh"
      Top             =   450
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   4
      Left            =   2250
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   7
      Tag             =   "fb"
      ToolTipText     =   "This is the color that will be the BackRound color of the button when the butten is push Down"
      Top             =   1650
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   5
      Left            =   5025
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   6
      Tag             =   "fb"
      ToolTipText     =   "This is the color that will be the buttons text color when the butten is depressed"
      Top             =   1050
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   7
      Left            =   5025
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   5
      Tag             =   "fb"
      Top             =   1650
      Width           =   540
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Origanal"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   225
      TabIndex        =   4
      ToolTipText     =   "This is what the program normal has"
      Top             =   4200
      Width           =   1440
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Grey"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   225
      TabIndex        =   3
      ToolTipText     =   "This is one of my favoirt"
      Top             =   4500
      Width           =   1440
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Easter Colors"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   225
      TabIndex        =   2
      ToolTipText     =   "Nothing like Easter"
      Top             =   4800
      Width           =   1440
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "American Colors"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   1725
      TabIndex        =   1
      ToolTipText     =   "American Color. the good old reb white and blue"
      Top             =   4200
      Width           =   1590
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Golf Colors"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   1725
      TabIndex        =   0
      ToolTipText     =   "Color that remind me of golf"
      Top             =   4500
      Width           =   1590
   End
   Begin VB.Label Move 
      BackStyle       =   0  'Transparent
      Height          =   435
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   5835
   End
   Begin VB.Label tt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Program Colors"
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
      Height          =   315
      Left            =   75
      TabIndex        =   26
      Tag             =   "made by leo/jason"
      Top             =   75
      Width           =   5640
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Program Text"
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
      Height          =   390
      Left            =   225
      TabIndex        =   25
      Tag             =   "made by leo/jason"
      Top             =   525
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Program BG"
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
      Height          =   390
      Left            =   225
      TabIndex        =   24
      Tag             =   "made by leo/jason"
      Top             =   2325
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Text Box BG"
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
      Height          =   390
      Left            =   225
      TabIndex        =   23
      Tag             =   "made by leo/jason"
      Top             =   1125
      Width           =   1515
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Text Box Text"
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
      Height          =   390
      Left            =   3000
      TabIndex        =   22
      Tag             =   "made by leo/jason"
      Top             =   525
      Width           =   1515
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Button Down BG"
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
      Height          =   390
      Left            =   225
      TabIndex        =   21
      Tag             =   "made by leo/jason"
      Top             =   1725
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Button Down Text"
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
      Height          =   390
      Left            =   3000
      TabIndex        =   20
      Tag             =   "made by leo/jason"
      Top             =   1125
      Width           =   2040
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Color Box BG"
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
      Height          =   390
      Left            =   3000
      TabIndex        =   19
      Tag             =   "made by leo/jason"
      Top             =   1725
      Width           =   1965
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
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
      Height          =   315
      Left            =   75
      TabIndex        =   18
      Tag             =   "fb"
      ToolTipText     =   "This is the box that you can see color changes in"
      Top             =   5475
      Width           =   5640
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Preview Button"
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
      Height          =   315
      Left            =   1050
      TabIndex        =   17
      Tag             =   "fb"
      ToolTipText     =   "This is a preview of a button and it's colors"
      Top             =   6375
      Width           =   4590
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
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
      Height          =   315
      Left            =   3075
      TabIndex        =   16
      Tag             =   "made by leo/jason"
      ToolTipText     =   "A forget what i did i like the setting i have now"
      Top             =   3375
      Width           =   1515
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Height          =   315
      Left            =   1275
      TabIndex        =   15
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Accept Changes"
      Top             =   3375
      Width           =   1515
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "PreSet Settings"
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
      Height          =   315
      Left            =   75
      TabIndex        =   14
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Setting that have been predifined"
      Top             =   3825
      Width           =   5640
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   3075
      Top             =   3375
      Width           =   1515
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   1275
      Top             =   3375
      Width           =   1515
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   3390
      Left            =   75
      Top             =   375
      Width           =   5640
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   1290
      Left            =   75
      Top             =   4125
      Width           =   5640
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   75
      Top             =   3825
      Width           =   5640
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   75
      Tag             =   "fb"
      Top             =   5475
      Width           =   5640
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   1050
      Tag             =   "fb"
      Top             =   6375
      Width           =   4590
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   1065
      Left            =   75
      Tag             =   "fb"
      Top             =   5775
      Width           =   5640
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   75
      Top             =   75
      Width           =   5640
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tMdBgColor
Private tMdFgColor
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private MouseDownForm
Private MouseDownFormX
Private MouseDownFormY
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Sub move_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDownForm = 1
MouseDownFormX = X
MouseDownFormY = Y
End Sub
Private Sub move_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDownForm = 0
End Sub
Private Sub move_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DoEvents
If MouseDownForm <> 1 Then
Exit Sub
End If
Dim Z As POINTAPI
Call GetCursorPos(Z)
Form4.Top = (Z.Y * 15) - MouseDownFormY
Form4.Left = (Z.X * 15) - MouseDownFormX
End Sub
Private Sub Check1_Click(Index As Integer)
If Check1(Index).Value = 1 Then
For X = 0 To Check1.UBound
If X = Index Then
Else
Check1(X).Value = 0
End If
Next
If Index = 0 Then
ntempbb = Invis
changecolor &HC07800, vbWhite, vbBlack, vbWhite, vbBlack, vbWhite, vbBlack, vbWhite
UpdateColors
End If
If Index = 1 Then
ntempbb = Invis
changecolor 12632256, 4210752, 4210752, 8421504, 14737632, 14737632, 4210752, 8421504
UpdateColors
End If
If Index = 2 Then
ntempbb = Invis
changecolor 12632319, 4210752, 0, 16777215, 8421504, 16744576, 16777215, 16761087
UpdateColors
End If
If Index = 3 Then
ntempbb = Invis
changecolor 255, 16777215, 4210752, 16711680, 14737632, 16711680, 16777215, 16777215
UpdateColors
End If
If Index = 4 Then
ntempbb = Invis
changecolor 192, 16384, 4210752, 8438015, 4210752, 32768, 16777215, 64
UpdateColors
End If
If Index = 5 Then
ntempbb = Invis
changecolor 32768, 8454016, 16711680, 49152, 16384, 32768, 65280, 16384
UpdateColors
End If
If Index = 6 Then
ntempbb = Invis
changecolor 65280, 255, 16711680, 65535, 16711680, 33023, 16776960, 16711935
UpdateColors
End If
If Index = 7 Then
ntempbb = Invis
changecolor 8438015, 16711680, 16711680, 12632064, 16777152, 33023, 16711680, 8421631
UpdateColors
End If
If Index = 8 Then
ntempbb = Invis
changecolor 33023, 16384, 16711680, 8438015, 4210752, 16761024, 0, 12640511
UpdateColors
End If
UpdatePreview
End If
End Sub

Private Sub Form_Load()
UpdateColors
UpdatePreview
End Sub
Private Sub Label10_Click()
Form4.Hide
End Sub
Private Sub Label11_Click()
changecolor Picture1(1).BackColor, Picture1(0).BackColor, vbBlue, Picture1(4).BackColor, Picture1(5).BackColor, Picture1(2).BackColor, Picture1(3).BackColor, Picture1(7).BackColor
Me.Hide
End Sub
Private Sub Label12_Click()
MyMSG "Preview", "This Is A Preview"
End Sub
Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = tMdFgColor
Shape8.BackColor = tMdBgColor
End Sub
Private Sub Label12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UpdatePreview
End Sub
Private Sub Picture1_Click(Index As Integer)
nTemp = CustomColor
If nTemp <> "N/A" Then
Picture1(Index).BackColor = nTemp
Picture1(Index).Tag = "1"
End If
UpdatePreview
End Sub
Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.BackColor = MdBgColor
Label11.ForeColor = MdFgColor
End Sub
Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.BackColor = BGcolor
Label11.ForeColor = FgColor
End Sub
Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape6.BackColor = MdBgColor
Label10.ForeColor = MdFgColor
End Sub
Private Sub Label10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape6.BackColor = BGcolor
Label10.ForeColor = FgColor
End Sub
Private Sub UpdatePreview()
Shape4.FillColor = Picture1(1).BackColor
Shape3.FillColor = Picture1(1).BackColor
Shape8.BackColor = Picture1(1).BackColor
Label9.ForeColor = Picture1(0).BackColor
Label12.ForeColor = Picture1(0).BackColor
Picture9.BackColor = Picture1(7).BackColor
tMdBgColor = Picture1(4).BackColor
tMdFgColor = Picture1(5).BackColor
Text1.BackColor = Picture1(2).BackColor
Text1.ForeColor = Picture1(3).BackColor
End Sub
Private Sub UpdateColors()
Picture1(0).BackColor = FgColor
Picture1(1).BackColor = BGcolor
Picture1(2).BackColor = TxBgColor
Picture1(3).BackColor = TxFgColor
Picture1(4).BackColor = MdBgColor
Picture1(5).BackColor = MdFgColor
Picture1(7).BackColor = CbBgColor
End Sub

