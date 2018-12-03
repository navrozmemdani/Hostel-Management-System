VERSION 5.00
Begin VB.Form objective 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin HMSLFMJD.jcbutton btnback 
      Height          =   735
      Left            =   18840
      TabIndex        =   10
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16744576
      Caption         =   ""
      Picture         =   "vision-LF_MJD.frx":0000
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Timer Timer8 
      Interval        =   900
      Left            =   4320
      Top             =   2640
   End
   Begin VB.Timer Timer7 
      Interval        =   900
      Left            =   3720
      Top             =   2640
   End
   Begin VB.Timer Timer6 
      Interval        =   900
      Left            =   7440
      Top             =   1200
   End
   Begin VB.Timer Timer5 
      Interval        =   900
      Left            =   6720
      Top             =   1200
   End
   Begin VB.Timer Timer4 
      Interval        =   900
      Left            =   7440
      Top             =   240
   End
   Begin VB.Timer Timer3 
      Interval        =   900
      Left            =   6600
      Top             =   240
   End
   Begin VB.Timer Timer2 
      Interval        =   900
      Left            =   3000
      Top             =   4920
   End
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   2640
      Top             =   5640
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GOALS AND OBJECTIVES"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   9
      Top             =   2520
      Width           =   7935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MISSION"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   8
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  VISION"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   8760
      TabIndex        =   7
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4365
      Left            =   840
      Picture         =   "vision-LF_MJD.frx":10DA
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   18810
   End
   Begin VB.Image Image2 
      Height          =   4335
      Left            =   4800
      Picture         =   "vision-LF_MJD.frx":37ECC2
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   11055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "                            Dual Objective: ""Giving elderly a Grand Child's Love and Loving Grand Parents to a Lovely Child"""
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   20415
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   20055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"vision-LF_MJD.frx":43ACF4
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   20415
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   20055
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GOALS AND OBJECTIVES"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   6720
      TabIndex        =   4
      Top             =   2520
      Width           =   7935
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   20055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"vision-LF_MJD.frx":43AE0A
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   19935
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   20055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " MISSION"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   8640
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   20055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "                   TO TRANSFORM A LESS  PRIVILEGED PERSON INTO A SUCCESSFUL HUMAN BEING"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   15855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   20055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  VISION"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   20055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   4575
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   19935
   End
End
Attribute VB_Name = "objective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnback_Click()
objectivve.Hide
MDIForm1.Show
End Sub

Private Sub Form_Load()
MDIForm1.Show
vision.Hide
'Unload Me
End Sub

Private Sub Timer1_Timer()
Image2.Visible = True
Image1.Visible = False
End Sub

Private Sub Timer2_Timer()
Image2.Visible = False
Image1.Visible = True
End Sub

Private Sub Timer3_Timer()
Label1.Visible = False
Label8.Visible = True
End Sub

Private Sub Timer4_Timer()
Label1.Visible = True
Label8.Visible = False
End Sub

Private Sub Timer5_Timer()
Label3.Visible = False
Label9.Visible = True
End Sub

Private Sub Timer6_Timer()
Label3.Visible = True
Label9.Visible = False
End Sub

Private Sub Timer7_Timer()
Label5.Visible = True
Label10.Visible = False
End Sub

Private Sub Timer8_Timer()
Label5.Visible = False
Label10.Visible = True
End Sub



