VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form forgetpass 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   9555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin HMSLFMJD.jcbutton btnchangepass 
      Height          =   735
      Left            =   11400
      TabIndex        =   19
      Top             =   7560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   128
      Caption         =   "CHANGE PASSWORD"
      Picture         =   "forgetpassword.frx":0000
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnnexit 
      Height          =   735
      Left            =   8640
      TabIndex        =   18
      Top             =   7560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   128
      Caption         =   "EXIT"
      Picture         =   "forgetpassword.frx":10DA
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnverify 
      Height          =   735
      Left            =   14520
      TabIndex        =   17
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   128
      Caption         =   "VERIFY"
      Picture         =   "forgetpassword.frx":21B4
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btncheck 
      Height          =   855
      Left            =   14400
      TabIndex        =   16
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   128
      Caption         =   "CHECK"
      Picture         =   "forgetpassword.frx":328E
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
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
      TabIndex        =   15
      Top             =   6600
      Width           =   5655
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
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
      TabIndex        =   14
      Top             =   5760
      Width           =   5415
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
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
      TabIndex        =   13
      Top             =   3960
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
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
      TabIndex        =   12
      Top             =   2640
      Width           =   5535
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3720
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   3360
      Top             =   1440
   End
   Begin VB.Timer Timer3 
      Interval        =   600
      Left            =   3960
      Top             =   1440
   End
   Begin VB.Timer Timer4 
      Interval        =   400
      Left            =   4680
      Top             =   1560
   End
   Begin VB.Timer Timer5 
      Interval        =   200
      Left            =   4200
      Top             =   1200
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3840
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Connect         =   $"forgetpassword.frx":4368
      OLEDBString     =   $"forgetpassword.frx":43F8
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
   Begin VB.Label Label6 
      BackColor       =   &H000000C0&
      Caption         =   " Confirm  Password :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Width           =   3615
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000C0&
      Caption         =   " Enter New Password :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      Caption         =   " Enter the Mobile No :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000C0&
      Caption         =   " Enter the User Name:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8400
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   5895
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8400
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   5775
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8400
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   8400
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   5775
   End
   Begin VB.Label AL2 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                CHANGE PASSWORD"
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
      TabIndex        =   11
      Top             =   1080
      Width           =   12735
   End
   Begin VB.Label AL3 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                CHANGE PASSWORD"
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
      TabIndex        =   10
      Top             =   1080
      Width           =   12735
   End
   Begin VB.Label AL4 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                CHANGE PASSWORD "
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
      TabIndex        =   9
      Top             =   1080
      Width           =   12735
   End
   Begin VB.Label AL5 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                CHANGE PASSWORD"
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
      TabIndex        =   8
      Top             =   1080
      Width           =   12735
   End
   Begin VB.Label AL1 
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                CHANGE PASSWORD"
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
      TabIndex        =   7
      Top             =   1080
      Width           =   12735
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H00C0C0C0&
      FillStyle       =   3  'Vertical Line
      Height          =   6855
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   13455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   5280
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   4800
      Width           =   5775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      Top             =   3480
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   6735
      Left            =   3240
      Picture         =   "forgetpassword.frx":4488
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   13290
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   13335
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00404040&
      Height          =   7095
      Left            =   -1800
      Shape           =   3  'Circle
      Top             =   360
      Width           =   11535
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00404040&
      Height          =   7095
      Left            =   10920
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   11535
   End
End
Attribute VB_Name = "forgetpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnchangepass_Click()
If Text7.Text = Text6.Text Then
Adodc1.Recordset.Fields("Password") = Text6.Text
Adodc1.Recordset.Update
MsgBox "Password Changed Successfully", vbInformation, "Password Change: Success"
forgetpass.Hide
adminlogin.Show
Else
MsgBox "Password Does not matched,Please Enter Correct Details", vbExclamation, "Change Password: Failed"
Text7.Text = ""
Text6.Text = ""
End If
End Sub

Private Sub btncheck_Click()
Adodc1.RecordSource = " select  *  from Users where Username='" + Text1.Text + "' "
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Label5.Caption = "User ID Not Found ..Sorry Can't Reset The Password!!!"
Label5.ForeColor = &HFF&
Else
Label5.Caption = "User ID Found in the Database"
Label5.ForeColor = &H8000&
End If
End Sub


Private Sub btnverify_Click()
Adodc1.RecordSource = " select  *  from Users where phonenumber='" + Text4.Text + "' "
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
Label2.Caption = "Account not Verified, Can't Reset The Password!!!"
Label1.Caption = "Sorry..Mobile Number Not Matched!!!"
Label1.ForeColor = &HFF&
Label2.ForeColor = &HFF&
Else
Label1.ForeColor = &H8000&
Label2.ForeColor = &H8000&
Label1.Caption = "Congratulations!!"
Label2.Caption = "Account is verified Now,set your new Password"

Text2.Visible = True
Text3.Visible = True
btnchangepass.Visible = True
Label7.Visible = True
Label6.Visible = True
Shape3.Visible = True
Shape4.Visible = True
Shape9.Visible = True
Shape10.Visible = True
End If
End Sub

Private Sub btnnexit_Click()
forgetpass.Hide
adminlogin.Show
End Sub

Private Sub Form_Load()
Text2.Visible = False
Text3.Visible = False
btnchangepass.Visible = False
Label7.Visible = False
Label6.Visible = False
Shape3.Visible = False
Shape4.Visible = False
Shape9.Visible = False
Shape10.Visible = False
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


