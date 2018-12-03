VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form regadmin 
   BackColor       =   &H00004080&
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin HMSLFMJD.jcbutton btnsignup 
      Height          =   735
      Left            =   11880
      TabIndex        =   16
      Top             =   7440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16512
      Caption         =   "SIGNUP"
      Picture         =   "adminsignup.frx":0000
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnexit 
      Height          =   735
      Left            =   9360
      TabIndex        =   15
      Top             =   7440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16512
      Caption         =   "EXIT"
      Picture         =   "adminsignup.frx":10DA
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
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
      Height          =   495
      Left            =   8880
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   6600
      Width           =   6015
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
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
      Height          =   495
      Left            =   8880
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   5760
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
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
      Height          =   495
      Left            =   8880
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4920
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
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
      Left            =   9000
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3960
      Width           =   6015
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
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
      Height          =   495
      Left            =   9000
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3240
      Width           =   6015
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   4800
      Top             =   1800
   End
   Begin VB.Timer Timer3 
      Interval        =   600
      Left            =   4320
      Top             =   1920
   End
   Begin VB.Timer Timer4 
      Interval        =   400
      Left            =   4920
      Top             =   1920
   End
   Begin VB.Timer Timer5 
      Interval        =   200
      Left            =   4080
      Top             =   1920
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   4320
      Top             =   7800
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Connect         =   $"adminsignup.frx":21B4
      OLEDBString     =   $"adminsignup.frx":2244
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Users"
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
   Begin VB.Shape Shape19 
      BackColor       =   &H00FFC0FF&
      BackStyle       =   1  'Opaque
      DrawMode        =   15  'Merge Pen Not
      Height          =   1815
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   8520
      Width           =   2055
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      DrawMode        =   15  'Merge Pen Not
      Height          =   855
      Left            =   18240
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape17 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      DrawMode        =   15  'Merge Pen Not
      Height          =   1095
      Left            =   6000
      Shape           =   3  'Circle
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H000040C0&
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
      Left            =   4680
      TabIndex        =   2
      Top             =   6600
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      Caption         =   "    PHONE NO:"
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
      Left            =   4680
      TabIndex        =   3
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H000040C0&
      Caption         =   "  PASSWORD :"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   4920
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
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
      Left            =   4680
      TabIndex        =   0
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   3975
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   3975
   End
   Begin VB.Label Label5 
      BackColor       =   &H000040C0&
      Caption         =   "  FULL NAME  :"
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
      Left            =   4680
      TabIndex        =   4
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   6495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   6495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8640
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   6495
   End
   Begin VB.Label AL2 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "            ADMINISTRATOR SIGN UP"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   1680
      Width           =   12735
   End
   Begin VB.Label AL3 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "            ADMINISTRATOR SIGN UP"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   1680
      Width           =   12735
   End
   Begin VB.Label AL4 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "            ADMINISTRATOR SIGN UP"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   1680
      Width           =   12735
   End
   Begin VB.Label AL5 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "            ADMINISTRATOR SIGN UP"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   1680
      Width           =   12735
   End
   Begin VB.Label AL1 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "            ADMINISTRATOR SIGN UP"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   1680
      Width           =   12735
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00C0C0C0&
      FillStyle       =   3  'Vertical Line
      Height          =   6375
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   13455
   End
   Begin VB.Image Image1 
      Height          =   6135
      Left            =   3480
      Picture         =   "adminsignup.frx":22D4
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   12810
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   3360
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   12975
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      DrawMode        =   15  'Merge Pen Not
      Height          =   3495
      Left            =   13440
      Shape           =   3  'Circle
      Top             =   0
      Width           =   2295
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      DrawMode        =   15  'Merge Pen Not
      Height          =   3855
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   3855
   End
   Begin VB.Shape Shape14 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      DrawMode        =   15  'Merge Pen Not
      Height          =   1695
      Left            =   600
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      DrawMode        =   15  'Merge Pen Not
      Height          =   2415
      Left            =   15480
      Shape           =   3  'Circle
      Top             =   4560
      Width           =   2895
   End
End
Attribute VB_Name = "regadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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

Private Sub btnsignup_Click()
Adodc1.Recordset.Fields("username") = Text1.Text
Adodc1.Recordset.Fields("password") = Text2.Text
Adodc1.Recordset.Fields("email") = Text3.Text
Adodc1.Recordset.Fields("phonenumber") = Text4.Text
Adodc1.Recordset.Fields("nameofperson") = Text5.Text
Adodc1.Recordset.Update
MsgBox ("Now  Enter Your New Username And Password")
adminlogin.Show
regadmin.Hide
End Sub

Private Sub btnexit_Click()
adminlogin.Show
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
adminlogin.Show
regadmin.Hide
End Sub

