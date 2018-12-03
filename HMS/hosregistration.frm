VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form hosregisteration 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   10470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10470
   ScaleWidth      =   20250
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin HMSLFMJD.jcbutton btnsignup 
      Height          =   855
      Left            =   12840
      TabIndex        =   12
      Top             =   6960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      ButtonStyle     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483647
      Caption         =   "SIGNUP"
      Picture         =   "hosregistration.frx":0000
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnexit 
      Height          =   855
      Left            =   9240
      TabIndex        =   11
      Top             =   6960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      ButtonStyle     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483647
      Caption         =   "EXIT"
      Picture         =   "hosregistration.frx":10DA
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   8520
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   6000
      Width           =   6735
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   8520
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   4680
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   8400
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3240
      Width           =   6855
   End
   Begin VB.Timer Timer4 
      Left            =   4560
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   4080
      Top             =   1200
   End
   Begin VB.Timer Timer3 
      Interval        =   600
      Left            =   5280
      Top             =   1440
   End
   Begin VB.Timer Timer5 
      Interval        =   200
      Left            =   5760
      Top             =   1200
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3960
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"hosregistration.frx":21B4
      OLEDBString     =   $"hosregistration.frx":2247
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "studentregistration"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Caption         =   "      EMAIL:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   6120
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      Caption         =   "  PASSWORD :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   4800
      Width           =   3495
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   3735
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      Caption         =   " USER NAME :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8280
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   7215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   7215
   End
   Begin VB.Label AL2 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "               STUDENT SIGN UP"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   3600
      TabIndex        =   7
      Top             =   1080
      Width           =   12615
   End
   Begin VB.Label AL3 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "               STUDENT SIGN UP"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   3600
      TabIndex        =   6
      Top             =   1080
      Width           =   12615
   End
   Begin VB.Label AL4 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "               STUDENT SIGN UP "
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   3600
      TabIndex        =   5
      Top             =   1080
      Width           =   12615
   End
   Begin VB.Label AL5 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "               STUDENT SIGN UP"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   12615
   End
   Begin VB.Label AL1 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "               STUDENT SIGN UP"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   855
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   12615
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00C0C0C0&
      FillStyle       =   3  'Vertical Line
      Height          =   6495
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   13095
   End
   Begin VB.Image Image1 
      Height          =   6375
      Left            =   3480
      Picture         =   "hosregistration.frx":22DA
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   12810
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   12855
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      DrawMode        =   7  'Invert
      Height          =   2775
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   15135
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      DrawMode        =   7  'Invert
      Height          =   2775
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   15135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      DrawMode        =   7  'Invert
      Height          =   2775
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   15135
   End
End
Attribute VB_Name = "hosregisteration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnsignup_Click()
Adodc1.Recordset.Fields("username") = Text1.Text
Adodc1.Recordset.Fields("password") = Text2.Text
Adodc1.Recordset.Fields("email") = Text3.Text
Adodc1.Recordset.Update
MsgBox ("You have Registered Successfully")
MDIForm1.Show
hosregisteration.Hide
End Sub

Private Sub btnexit_Click()
MDIForm1.Show
hosregisteration.Hide
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
MDIForm1.Show
hosregisteration.Hide
End Sub

Private Sub Timer1_Timer()
AL1.Visible = True
AL2.Visible = False
AL3.Visible = False
AL4.Visible = False
AL5.Visible = False
End Sub

Private Sub Timer2_Timer()
AL1.Visible = False
AL2.Visible = True
AL3.Visible = False
AL4.Visible = False
AL5.Visible = False
End Sub

Private Sub Timer3_Timer()
AL1.Visible = False
AL2.Visible = False
AL3.Visible = True
AL4.Visible = False
AL5.Visible = False
End Sub

Private Sub Timer4_Timer()
AL1.Visible = False
AL2.Visible = False
AL3.Visible = False
AL4.Visible = True
AL5.Visible = False
End Sub

Private Sub Timer5_Timer()
AL1.Visible = False
AL2.Visible = False
AL3.Visible = False
AL4.Visible = False
AL5.Visible = True
End Sub
