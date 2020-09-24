VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C07800&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6870
   LinkTopic       =   "Form2"
   ScaleHeight     =   4065
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Blue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      MaxLength       =   3
      TabIndex        =   56
      Tag             =   "t"
      Top             =   3660
      Width           =   735
   End
   Begin VB.TextBox Green 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      MaxLength       =   3
      TabIndex        =   55
      Tag             =   "t"
      Top             =   3390
      Width           =   735
   End
   Begin VB.TextBox red 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6000
      MaxLength       =   3
      TabIndex        =   54
      Tag             =   "t"
      Top             =   3120
      Width           =   735
   End
   Begin VB.PictureBox cColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4080
      ScaleHeight     =   825
      ScaleWidth      =   945
      TabIndex        =   53
      Tag             =   "fb"
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   47
      Left            =   3600
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   48
      Tag             =   "fb"
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   46
      Left            =   3120
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   47
      Tag             =   "fb"
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   45
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   46
      Tag             =   "fb"
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   44
      Left            =   2160
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   45
      Tag             =   "fb"
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00004040&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   43
      Left            =   1680
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   44
      Tag             =   "fb"
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   42
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   43
      Tag             =   "fb"
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   42
      Tag             =   "fb"
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   240
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   41
      Tag             =   "fb"
      Top             =   3120
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   3600
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   40
      Tag             =   "fb"
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   3120
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   39
      Tag             =   "fb"
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   38
      Tag             =   "fb"
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   2160
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   37
      Tag             =   "fb"
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   1680
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   36
      Tag             =   "fb"
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   35
      Tag             =   "fb"
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   34
      Tag             =   "fb"
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   240
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   33
      Tag             =   "fb"
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   3600
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   32
      Tag             =   "fb"
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   3120
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   31
      Tag             =   "fb"
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   30
      Tag             =   "fb"
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   2160
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   29
      Tag             =   "fb"
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   1680
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   28
      Tag             =   "fb"
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   27
      Tag             =   "fb"
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   26
      Tag             =   "fb"
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   240
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   25
      Tag             =   "fb"
      Top             =   2160
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   3600
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   24
      Tag             =   "fb"
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   3120
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   23
      Tag             =   "fb"
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   22
      Tag             =   "fb"
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   2160
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   21
      Tag             =   "fb"
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   1680
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   20
      Tag             =   "fb"
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   19
      Tag             =   "fb"
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   18
      Tag             =   "fb"
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   240
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      Tag             =   "fb"
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   3600
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   16
      Tag             =   "fb"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   3120
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   15
      Tag             =   "fb"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   14
      Tag             =   "fb"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   2160
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   13
      Tag             =   "fb"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1680
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   12
      Tag             =   "fb"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   11
      Tag             =   "fb"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   10
      Tag             =   "fb"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   240
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   9
      Tag             =   "fb"
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3600
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   8
      Tag             =   "fb"
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3120
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   7
      Tag             =   "fb"
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   2640
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Tag             =   "fb"
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   2160
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Tag             =   "fb"
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1680
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Tag             =   "fb"
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1200
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Tag             =   "fb"
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   720
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Tag             =   "fb"
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Pick 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Tag             =   "fb"
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   4080
      Picture         =   "RemadeFader2.frx":0000
      ScaleHeight     =   2745
      ScaleWidth      =   2610
      TabIndex        =   0
      Tag             =   "fb"
      ToolTipText     =   "Click somewhere on the picture to difine your cutom color"
      Top             =   120
      Width           =   2640
   End
   Begin VB.Label Move 
      BackStyle       =   0  'Transparent
      Height          =   585
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   4065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
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
      Left            =   5400
      TabIndex        =   59
      Top             =   3675
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5400
      TabIndex        =   58
      Top             =   3390
      Width           =   735
   End
   Begin VB.Label a 
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
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
      Left            =   5415
      TabIndex        =   57
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label ButtonA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Custom>"
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
      Left            =   2760
      TabIndex        =   52
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label ButtonA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   51
      Top             =   3600
      Width           =   1215
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
      Left            =   120
      TabIndex        =   50
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C07800&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Colors"
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
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Width           =   3855
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   47
      Left            =   3480
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   46
      Left            =   3000
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   45
      Left            =   2520
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   44
      Left            =   2040
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   43
      Left            =   1560
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   42
      Left            =   1080
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   41
      Left            =   600
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   40
      Left            =   120
      Top             =   3000
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   39
      Left            =   3480
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   38
      Left            =   3000
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   37
      Left            =   2520
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   36
      Left            =   2040
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   35
      Left            =   1560
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   34
      Left            =   1080
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   33
      Left            =   600
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   32
      Left            =   120
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   31
      Left            =   3480
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   30
      Left            =   3000
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   29
      Left            =   2520
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   28
      Left            =   2040
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   27
      Left            =   1560
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   26
      Left            =   1080
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   25
      Left            =   600
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   24
      Left            =   120
      Top             =   2040
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   23
      Left            =   3480
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   22
      Left            =   3000
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   21
      Left            =   2520
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   20
      Left            =   2040
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   19
      Left            =   1560
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   18
      Left            =   1080
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   17
      Left            =   600
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   16
      Left            =   120
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   15
      Left            =   3480
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   14
      Left            =   3000
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   13
      Left            =   2520
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   12
      Left            =   2040
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   11
      Left            =   1560
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   10
      Left            =   1080
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   9
      Left            =   600
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   8
      Left            =   120
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   7
      Left            =   3480
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   6
      Left            =   3000
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   5
      Left            =   2520
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   4
      Left            =   2040
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   3
      Left            =   1560
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   2
      Left            =   1080
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   1
      Left            =   600
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape OutLine 
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
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
If MouseDownForm <> 1 Then
Exit Sub
End If
Dim Z As POINTAPI
Call GetCursorPos(Z)
Form2.Top = (Z.Y * 15) - MouseDownFormY
Form2.Left = (Z.X * 15) - MouseDownFormX
End Sub
Private Sub Form_Load()
For nTemp = 0 To 47
OutLine(nTemp).Visible = False
Next
Form2.Width = 4080
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
If Index = 2 And ButtonA(2).caption = "Custom>" Then
ButtonA(2).caption = "<Custom"
Form2.Width = 6825
Exit Sub
End If
If Index = 2 And ButtonA(2).caption = "<Custom" Then
ButtonA(2).caption = "Custom>"
Form2.Width = 4080
Exit Sub
End If
If Index = 1 Then
cColorA = "N/A"
Form2.Hide
Exit Sub
End If
If Index = 0 Then
cColorA = cColor.BackColor
Form2.Hide
Exit Sub
End If
End Sub
Private Sub Pick_Click(Index As Integer)
For nTemp = 0 To 47
OutLine(nTemp).Visible = False
Next
OutLine(Index).Visible = True
cColor.BackColor = Pick(Index).BackColor
nTemp = Hex(Pick(Index).BackColor)
If Len(nTemp) = 2 Then
nTemp = "0000" & nTemp
End If
If Len(nTemp) = 4 Then
nTemp = "00" & nTemp
End If
red.text = Val("&H" & (Right(nTemp, 2)))
Green.text = Val("&H" & (Right(Left(nTemp, 4), 2)))
Blue.text = Val("&H" & (Left(nTemp, 2)))
End Sub
Private Sub Pick_DblClick(Index As Integer)
cColorA = cColor.BackColor
Form2.Hide
End Sub
Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cColor.BackColor = Picture2.Point(X, Y)
nTemp = Hex(cColor.BackColor)
red.text = Val("&H" & (Right(nTemp, 2)))
Green.text = Val("&H" & (Right(Left(nTemp, 4), 2)))
Blue.text = Val("&H" & (Left(nTemp, 2)))
End Sub
Private Sub red_Change()
On Error GoTo Error
If red.text <> "" And Green.text <> "" And Blue.text <> "" Then
cColor.BackColor = "&H" & IIf(Len(Hex(Blue.text)) = 2, Hex(Blue.text), Hex(Blue.text) & "0") & IIf(Len(Hex(Green.text)) = 2, Hex(Green.text), Hex(Green.text) & "0") & IIf(Len(Hex(red.text)) = 2, Hex(red.text), Hex(red.text) & "0")
End If
Exit Sub
Error:
ErrorHandler
End Sub
Private Sub Blue_Change()
On Error GoTo Error
If red.text <> "" And Green.text <> "" And Blue.text <> "" Then
cColor.BackColor = "&H" & IIf(Len(Hex(Blue.text)) = 2, Hex(Blue.text), Hex(Blue.text) & "0") & IIf(Len(Hex(Green.text)) = 2, Hex(Green.text), Hex(Green.text) & "0") & IIf(Len(Hex(red.text)) = 2, Hex(red.text), Hex(red.text) & "0")
End If
Exit Sub
Error:
ErrorHandler
End Sub
Private Sub Green_Change()
On Error GoTo Error
If red.text <> "" And Green.text <> "" And Blue.text <> "" Then
cColor.BackColor = "&H" & IIf(Len(Hex(Blue.text)) = 2, Hex(Blue.text), Hex(Blue.text) & "0") & IIf(Len(Hex(Green.text)) = 2, Hex(Green.text), Hex(Green.text) & "0") & IIf(Len(Hex(red.text)) = 2, Hex(red.text), Hex(red.text) & "0")
End If
Exit Sub
Error:
ErrorHandler
End Sub
Private Function HexCheck(ntext)
X = Hex(ntext)
If Len(X) = 1 Then
X = "0" & X
End If
HexCheck = X
End Function
