VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C07800&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4965
   LinkTopic       =   "Form6"
   ScaleHeight     =   1530
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label move 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4965
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Temp"
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
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   1500
      Width           =   4965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "There Has Been An Error In the Program It is recomend you close it but you can ignore it"
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
      Height          =   690
      Left            =   0
      TabIndex        =   5
      Tag             =   "made by leo/jason"
      ToolTipText     =   "There Has Been An Error In the Program It is recomend you close it but you can ignore it"
      Top             =   375
      Width           =   4965
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close "
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
      TabIndex        =   4
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click here to close down the program"
      Top             =   1125
      Width           =   1515
   End
   Begin VB.Label tt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "E"
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
      Left            =   0
      TabIndex        =   3
      Tag             =   "made by leo/jason"
      ToolTipText     =   "An Error Has Been Created"
      Top             =   0
      Width           =   4965
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ignore"
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
      Left            =   1725
      TabIndex        =   2
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click this button to ignore the error and go on with what you where doing"
      Top             =   1125
      Width           =   1515
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Details>"
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
      Left            =   3375
      TabIndex        =   1
      Tag             =   "made by leo/jason"
      ToolTipText     =   "Click here to bring up the details of the error"
      Top             =   1125
      Width           =   1515
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   4965
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   1515
      Left            =   0
      Top             =   1500
      Width           =   4965
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   1725
      Top             =   1125
      Width           =   1515
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   3375
      Top             =   1125
      Width           =   1515
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C07800&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   75
      Top             =   1125
      Width           =   1515
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C07800&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   0
      Tag             =   "made by leo/jason"
      Top             =   300
      Width           =   4965
   End
End
Attribute VB_Name = "Form6"
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
Form6.Top = (Z.Y * 15) - MouseDownFormY
Form6.Left = (Z.X * 15) - MouseDownFormX
End Sub
Private Sub Label2_Click()
Form6.Hide
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.BackColor = MdBgColor
Label2.ForeColor = MdFgColor
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.BackColor = BGcolor
Label2.ForeColor = FgColor
End Sub
Private Sub Label3_Click()
If Label3.caption = "Details>" Then
Form6.Height = 3030
Label3.caption = "<Details"
Else
Form6.Height = 1530
Label3.caption = "Details>"
End If
End Sub
Private Sub Label4_Click()
End
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BackColor = MdBgColor
Label4.ForeColor = MdFgColor
End Sub
Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape3.BackColor = BGcolor
Label4.ForeColor = FgColor
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.BackColor = MdBgColor
Label3.ForeColor = MdFgColor
End Sub
Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.BackColor = BGcolor
Label3.ForeColor = FgColor
End Sub

