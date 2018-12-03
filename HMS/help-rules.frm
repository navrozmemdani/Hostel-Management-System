VERSION 5.00
Begin VB.Form help 
   BackColor       =   &H00C0C000&
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
   Begin HMSLFMJD.jcbutton btnexit 
      Height          =   735
      Left            =   19200
      TabIndex        =   10
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1296
      ButtonStyle     =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "help-rules.frx":0000
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnsupport 
      Height          =   735
      Left            =   18360
      TabIndex        =   9
      Top             =   240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      ButtonStyle     =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Picture         =   "help-rules.frx":10DA
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   " 5) The student should get atleast  60% to continue studing in the hostel ,otherwise the admission wil be terminated"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   8
      Top             =   7200
      Width           =   5775
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   " 4) The student should follow the daily routine of the hostel setup by the management "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   6120
      Width           =   5775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "  3) The student and there parents should be present during the admission meeting time "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   5775
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   " 2) The student must have at least 60 % and more in the previous year exam to be eligible for the admission in the hostel"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   5
      Top             =   3720
      Width           =   5775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"help-rules.frx":21B4
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"help-rules.frx":2247
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   7200
      TabIndex        =   3
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"help-rules.frx":22FE
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   10080
      TabIndex        =   2
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   1815
      Left            =   10320
      Picture         =   "help-rules.frx":23BE
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   7200
      Picture         =   "help-rules.frx":118202
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   6975
      Left            =   13800
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   6255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      Height          =   7095
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   6255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   6975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   6255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      Caption         =   "                                                          RULES AND REGULATION"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   19215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      Caption         =   "                                                         HELP AND SUPPORT"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   19215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   20055
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnexit_Click()
help.Hide
MDIForm1.Show
End Sub

Private Sub Form_Load()
Shape3.Visible = False
Image1.Visible = False
Image2.Visible = False
Label1.Visible = False
Label4.Visible = False
Label3.Visible = False

End Sub



Private Sub btnsupport_Click()
Shape3.Visible = True
Image1.Visible = True
Image2.Visible = True
Label1.Visible = True
Label4.Visible = True
Label3.Visible = True
Label2.Visible = False
Shape4.Visible = False
Shape2.Visible = flase
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
End Sub
