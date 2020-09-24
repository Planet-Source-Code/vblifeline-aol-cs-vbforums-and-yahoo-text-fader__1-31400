VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C07800&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   7515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   LinkTopic       =   "Form3"
   ScaleHeight     =   7515
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check13 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Yahoo Chat"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2475
      TabIndex        =   45
      Top             =   2850
      Width           =   1380
   End
   Begin VB.CheckBox Check12 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Put In Font Code"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2700
      TabIndex        =   44
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1425
      TabIndex        =   42
      Tag             =   "t"
      Text            =   "MS Sans Serif"
      Top             =   1275
      Width           =   3855
   End
   Begin VB.CheckBox Check11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Vbforums.Com"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1050
      TabIndex        =   34
      Top             =   6225
      Width           =   1380
   End
   Begin VB.CheckBox Check10 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "HTML"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      TabIndex        =   38
      Top             =   6225
      Value           =   1  'Checked
      Width           =   930
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2925
      TabIndex        =   37
      Tag             =   "t"
      Text            =   """>"
      Top             =   6525
      Width           =   735
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   225
      TabIndex        =   36
      Tag             =   "t"
      Text            =   "<font style=""FONT-FAMILY:"
      Top             =   6525
      Width           =   2040
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4275
      TabIndex        =   35
      Tag             =   "t"
      Text            =   "</font>"
      Top             =   6525
      Width           =   1035
   End
   Begin VB.CheckBox Check8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Vbforums.Com"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1050
      TabIndex        =   26
      Top             =   5100
      Width           =   1380
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4275
      TabIndex        =   32
      Tag             =   "t"
      Text            =   "</a>"
      Top             =   5400
      Width           =   1035
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   225
      TabIndex        =   29
      Tag             =   "t"
      Text            =   "<Font Size="
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2625
      TabIndex        =   28
      Tag             =   "t"
      Text            =   ">"
      Top             =   5400
      Width           =   735
   End
   Begin VB.CheckBox Check9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "HTML"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      TabIndex        =   27
      Top             =   5100
      Value           =   1  'Checked
      Width           =   930
   End
   Begin VB.CheckBox Check7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Vbforums.Com"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1050
      TabIndex        =   25
      Top             =   3975
      Width           =   1380
   End
   Begin VB.CheckBox Check6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "HTML"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      TabIndex        =   24
      Top             =   3975
      Value           =   1  'Checked
      Width           =   930
   End
   Begin VB.CheckBox Check5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Vbforums.Com"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1050
      TabIndex        =   22
      Top             =   2850
      Width           =   1380
   End
   Begin VB.CheckBox Check4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "HTML"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      TabIndex        =   21
      Top             =   2850
      Value           =   1  'Checked
      Width           =   930
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4275
      TabIndex        =   17
      Tag             =   "t"
      Text            =   "</a>"
      Top             =   4275
      Width           =   1035
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2625
      TabIndex        =   16
      Tag             =   "t"
      Text            =   """>"
      Top             =   4275
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   225
      TabIndex        =   15
      Tag             =   "t"
      Text            =   "<a href="""
      Top             =   4275
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      TabIndex        =   12
      Tag             =   "t"
      Text            =   "4"
      Top             =   2100
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Vb Code = [B][G][R]"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2175
      TabIndex        =   11
      Top             =   1800
      Width           =   2040
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Html Code = [R][G][B]"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   225
      TabIndex        =   10
      Top             =   1800
      Value           =   1  'Checked
      Width           =   2040
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      Caption         =   "Make Faded Text A WebLink"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      TabIndex        =   9
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1425
      TabIndex        =   7
      Tag             =   "t"
      Text            =   "http://www.leotown.com"
      Top             =   900
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4275
      TabIndex        =   4
      Tag             =   "t"
      Top             =   3150
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2625
      TabIndex        =   3
      Tag             =   "t"
      Text            =   """>"
      Top             =   3150
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   225
      TabIndex        =   2
      Tag             =   "t"
      Text            =   "<font color="""
      Top             =   3150
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   43
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C07800&
      Caption         =   "Settings For Font"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   40
      Top             =   5925
      Width           =   1905
   End
   Begin VB.Shape Shape4 
      Height          =   990
      Left            =   150
      Top             =   6000
      Width           =   5280
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   41
      Top             =   6525
      Width           =   690
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2250
      TabIndex        =   39
      Top             =   6525
      Width           =   690
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1425
      TabIndex        =   33
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C07800&
      Caption         =   "Settings For font size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   30
      Top             =   4800
      Width           =   1905
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3375
      TabIndex        =   31
      Top             =   5400
      Width           =   690
   End
   Begin VB.Shape Shape6 
      Height          =   990
      Left            =   150
      Top             =   4875
      Width           =   5280
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C07800&
      Caption         =   "Settings for outputing faded Text URL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   23
      Top             =   3675
      Width           =   3480
   End
   Begin VB.Shape Shape2 
      Height          =   990
      Left            =   150
      Top             =   3750
      Width           =   5280
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C07800&
      Caption         =   "Settings for outputing faded Text:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   20
      Top             =   2550
      Width           =   2955
   End
   Begin VB.Shape Shape5 
      Height          =   765
      Left            =   150
      Top             =   1725
      Width           =   5280
   End
   Begin VB.Shape Shape3 
      Height          =   1140
      Left            =   150
      Top             =   525
      Width           =   5280
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3375
      TabIndex        =   19
      Top             =   4275
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Web Link"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1350
      TabIndex        =   18
      Top             =   4275
      Width           =   1365
   End
   Begin VB.Label ButtonA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   14
      Top             =   7050
      Width           =   2445
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C07800&
      Caption         =   "Size Of Text:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   13
      Top             =   2100
      Width           =   1230
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Web Link:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   8
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label Move 
      BackStyle       =   0  'Transparent
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5610
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Letter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3375
      TabIndex        =   6
      Top             =   3150
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ColorCode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1350
      TabIndex        =   5
      Top             =   3150
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      Height          =   990
      Left            =   150
      Top             =   2625
      Width           =   5280
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fading Settings"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form3"
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
Private Sub ButtonA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonA(Index).BackColor = MdBgColor
ButtonA(Index).ForeColor = MdFgColor
End Sub
Private Sub ButtonA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ButtonA(Index).BackColor = BGcolor
ButtonA(Index).ForeColor = FgColor
End Sub
Private Sub ButtonA_Click(Index As Integer)
If Index = 0 Then Me.Hide
End Sub

Private Sub Check10_Click()
If Check10.Value = 1 Then
Text12.text = "<font style=" & Chr(34) & "FONT-FAMILY:"
Text13.text = Chr(34) & ">"
Text14.text = "</font>"
Check11.Value = 0
Else
Check11.Value = 1
End If
End Sub
Private Sub Check11_Click()
If Check11.Value = 1 Then
Text12.text = "[FONT="
Text13.text = "]"
Text14.text = "[/FONT]"
Check10.Value = 0
Else
Check10.Value = 1
End If
End Sub
Private Sub Check13_Click()
If Check13.Value = 1 Then
Check4.Value = 0
Check5.Value = 0
Text1.text = "<#"
Text2.text = ">"
Text3.text = ""
Else
End If
End Sub
Private Sub Check2_Click()
If Check2.Value = 0 Then
Check2.Value = 0
Check3.Value = 1
Exit Sub
End If
Check3.Value = 0
End Sub
Private Sub Check3_Click()
If Check3.Value = 0 Then
Check3.Value = 0
Check2.Value = 1
Exit Sub
End If
Check2.Value = 0
End Sub
Private Sub Check4_Click()
If Check4.Value = 1 Then
Check5.Value = 0
Check13.Value = 0
Text1.text = "<font color=" & Chr(34)
Text2.text = Chr(34) & ">"
Text3.text = ""
Else
Check5.Value = 1
End If
End Sub
Private Sub Check5_Click()
If Check5.Value = 1 Then
Check4.Value = 0
Check13.Value = 0
Text1.text = "[COLOR="
Text2.text = "]"
Text3.text = "[/COLOR]"
Else
Check13.Value = 1
End If
End Sub
Private Sub Check6_Click()
If Check6.Value = 1 Then
Check7.Value = 0
Text6.text = "<a href=" & Chr(34)
Text7.text = Chr(34) & ">"
Text8.text = "</a>"
Else
Check7.Value = 1
End If
End Sub
Private Sub Check7_Click()
If Check7.Value = 1 Then
Check6.Value = 0
Text6.text = "[URL="
Text7.text = "]"
Text8.text = "[/URL]"
Else
Check6.Value = 1
End If
End Sub
Private Sub Check8_Click()
If Check8.Value = 1 Then
Check9.Value = 0
Text11.text = "[SIZE="
Text10.text = "]"
Text9.text = "[/SIZE]"
Else
Check9.Value = 1
End If
End Sub
Private Sub Check9_Click()
If Check9.Value = 1 Then
Check8.Value = 0
Text11.text = "<Font Size="
Text10.text = ">"
Text9.text = "</a>"
Else
Check8.Value = 1
End If
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
Form3.Top = (Z.Y * 15) - MouseDownFormY
Form3.Left = (Z.X * 15) - MouseDownFormX
End Sub

