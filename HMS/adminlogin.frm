VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form adminlogin 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   10380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10380
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin HMSLFMJD.jcbutton btnlogin 
      Height          =   735
      Left            =   14040
      TabIndex        =   13
      Top             =   6960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421504
      Caption         =   "LOGIN"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnnewreg 
      Height          =   735
      Left            =   10200
      TabIndex        =   12
      Top             =   6960
      Width           =   3255
      _ExtentX        =   5741
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
      BackColor       =   8421504
      Caption         =   "NEW ADMIN REGISTRATION"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnexit 
      Height          =   735
      Left            =   8160
      TabIndex        =   11
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
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
      BackColor       =   8421504
      Caption         =   "EXIT"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnresetpass 
      Height          =   495
      Left            =   10920
      TabIndex        =   10
      Top             =   6240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      ButtonStyle     =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421504
      Caption         =   "CHANGE PASSWORD"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00808080&
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      Top             =   6240
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   20.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1005
      IMEMode         =   3  'DISABLE
      Left            =   8880
      MultiLine       =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   5040
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   960
      Left            =   8880
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3480
      Width           =   6135
   End
   Begin VB.Timer Timer5 
      Interval        =   200
      Left            =   5040
      Top             =   2040
   End
   Begin VB.Timer Timer4 
      Interval        =   400
      Left            =   5400
      Top             =   1560
   End
   Begin VB.Timer Timer3 
      Interval        =   600
      Left            =   4800
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   4200
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3600
      Top             =   1560
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3840
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   $"adminlogin.frx":0000
      OLEDBString     =   $"adminlogin.frx":0090
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from  Users"
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
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00C0C0C0&
      FillStyle       =   3  'Vertical Line
      Height          =   6135
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   13815
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000010&
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
      Height          =   615
      Left            =   4800
      TabIndex        =   2
      Top             =   5280
      Width           =   3375
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   4  'Mask Not Pen
      Height          =   855
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000010&
      Caption         =   " USER NAME :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      DrawMode        =   4  'Mask Not Pen
      Height          =   855
      Left            =   4680
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   6375
   End
   Begin VB.Image Image2 
      Height          =   540
      Left            =   9000
      Picture         =   "adminlogin.frx":0120
      Top             =   3600
      Width           =   540
   End
   Begin VB.Label AL1 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                ADMINISTRATOR LOGIN"
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
      TabIndex        =   0
      Top             =   1560
      Width           =   12855
   End
   Begin VB.Label AL5 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                ADMINISTRATOR LOGIN"
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
      TabIndex        =   6
      Top             =   1560
      Width           =   12855
   End
   Begin VB.Label AL4 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Top             =   1560
      Width           =   12855
   End
   Begin VB.Label AL3 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                ADMINISTRATOR LOGIN"
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
      TabIndex        =   4
      Top             =   1560
      Width           =   12855
   End
   Begin VB.Label AL2 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                ADMINISTRATOR LOGIN"
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
      TabIndex        =   3
      Top             =   1560
      Width           =   12855
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   3480
      Picture         =   "adminlogin.frx":31D6
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   13290
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   13335
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00C0C0C0&
      Height          =   4695
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   19815
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00C0C0C0&
      Height          =   4695
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   19815
   End
End
Attribute VB_Name = "adminlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnlogin_Click()
Adodc1.RecordSource = " select  *  from Users where Username='" + Text1.Text + "'  and Password='" + Text2.Text + "' "
Adodc1.Refresh
If Adodc1.Recordset.EOF Then

   MsgBox "Login Failed.. Please login with correct credentials", vbCritical
   Unload Me
   adminlogin.Show
   Else
   MsgBox "Well Done..Login Successful", vbInformation
   MDIForm1.Show
   adminlogin.Hide
End If
End Sub

Private Sub btnnewreg_Click()
regadmin.Show
adminlogin.Hide
End Sub

Private Sub btnexit_Click()
If MsgBox("Are you sure to close this Application?", vbQuestion + vbYesNo, "System") = vbYes Then
    End
End If
End Sub

Private Sub Check1_Click()
If Check1.value = 1 Then
    Text2.PasswordChar = ""
    Text2.Font = "Rockwell"
Else
    Text2.Font = "Rockwell"
    Text2.PasswordChar = "*"
End If
End Sub

Private Sub btnresetpass_Click()
forgetpass.Show
adminlogin.Hide
End Sub

Private Sub Check2_Click()

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
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

