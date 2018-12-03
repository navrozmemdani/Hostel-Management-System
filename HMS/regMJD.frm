VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form regmjd 
   BackColor       =   &H000000C0&
   Caption         =   "Form1"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "Rockwell"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "regMJD.frx":0000
   ScaleHeight     =   9645
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin HMSLFMJD.jcbutton btnlast 
      Height          =   615
      Left            =   18720
      TabIndex        =   50
      Top             =   8040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      BackColor       =   192
      Caption         =   "LAST"
      Picture         =   "regMJD.frx":10CA
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnnext 
      Height          =   615
      Left            =   17160
      TabIndex        =   49
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
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
      BackColor       =   192
      Caption         =   "NEXT"
      Picture         =   "regMJD.frx":21A4
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnprevious 
      Height          =   615
      Left            =   18600
      TabIndex        =   48
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
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
      BackColor       =   192
      Caption         =   "PREVIOUS"
      Picture         =   "regMJD.frx":327E
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnfirst 
      Height          =   615
      Left            =   17160
      TabIndex        =   47
      Top             =   7320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
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
      BackColor       =   192
      Caption         =   "FIRST"
      Picture         =   "regMJD.frx":4358
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnexit 
      Height          =   495
      Left            =   17760
      TabIndex        =   46
      Top             =   6720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      ButtonStyle     =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   "EXIT"
      Picture         =   "regMJD.frx":5432
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnaddnew 
      Height          =   615
      Left            =   17760
      TabIndex        =   45
      Top             =   6000
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      ButtonStyle     =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   "ADD NEW"
      Picture         =   "regMJD.frx":650C
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnsave 
      Height          =   495
      Left            =   17760
      TabIndex        =   44
      Top             =   5400
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      ButtonStyle     =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   "SAVE"
      Picture         =   "regMJD.frx":75E6
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btndelete 
      Height          =   495
      Left            =   17760
      TabIndex        =   43
      Top             =   4800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      ButtonStyle     =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   "DELETE"
      Picture         =   "regMJD.frx":86C0
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnupdate 
      Height          =   495
      Left            =   17760
      TabIndex        =   42
      Top             =   4200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      ButtonStyle     =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   "UPDATE"
      Picture         =   "regMJD.frx":979A
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnupload 
      Height          =   495
      Left            =   17400
      TabIndex        =   41
      Top             =   3600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      ButtonStyle     =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   "UPLOAD PHOTO"
      Picture         =   "regMJD.frx":A874
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnsearch 
      Height          =   615
      Left            =   7320
      TabIndex        =   40
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      ButtonStyle     =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Caption         =   "SEARCH"
      Picture         =   "regMJD.frx":B94E
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.PictureBox Feesreceipt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15360
      ScaleHeight     =   555
      ScaleWidth      =   795
      TabIndex        =   39
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   37
      Top             =   7680
      Width           =   4695
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12120
      MultiLine       =   -1  'True
      TabIndex        =   36
      Top             =   6960
      Width           =   4575
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12120
      TabIndex        =   35
      Text            =   "Select Blood Group"
      Top             =   6240
      Width           =   4695
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12120
      TabIndex        =   31
      Text            =   "Select Council"
      Top             =   4800
      Width           =   4695
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12120
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   4080
      Width           =   4695
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   3360
      Width           =   4695
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   12120
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12120
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   1920
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   7680
      Width           =   5415
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   6960
      Width           =   5415
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   6240
      Width           =   5415
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5520
      Width           =   5055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "FEMALE"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6360
      TabIndex        =   12
      Top             =   4800
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      Caption         =   "MALE"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3720
      TabIndex        =   11
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   4080
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3360
      Width           =   5415
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   5415
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1950
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   17160
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   8520
      TabIndex        =   18
      Top             =   5520
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Format          =   110886913
      CurrentDate     =   43319
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   7440
      Picture         =   "regMJD.frx":CA28
      Stretch         =   -1  'True
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label18 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Health Details"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   38
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Shape Shape22 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   9360
      Top             =   5640
      Width           =   3975
   End
   Begin VB.Label Label17 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "3. Any Major Health Issue :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   34
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "2.Height ( in cm) :"
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
      Left            =   9600
      TabIndex        =   33
      Top             =   7080
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1.Blood Group :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   32
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Shape Shape21 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   9360
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   7575
   End
   Begin VB.Shape Shape20 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   9360
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   7575
   End
   Begin VB.Shape Shape19 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   9360
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   7575
   End
   Begin VB.Label Label15 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 14.Name of Council :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   30
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Shape Shape18 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   9360
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   7575
   End
   Begin VB.Label Label14 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "13.Name of Jamatkhana :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   25
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 12.Permanent Address :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   24
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 11.Email :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   23
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 10.Aadhaar No :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   22
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Shape Shape17 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   9360
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   7575
   End
   Begin VB.Shape Shape16 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   9360
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   7575
   End
   Begin VB.Shape Shape15 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   9360
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   7575
   End
   Begin VB.Shape Shape14 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   9360
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   7575
   End
   Begin VB.Label Label7 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 9.Mobile No :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   7800
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 8.Caste (with Sub-Caste) :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "7.Age :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 6.Date of Birth :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Shape Shape13 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   8655
   End
   Begin VB.Shape Shape12 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   8655
   End
   Begin VB.Shape Shape11 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   7680
      Width           =   8655
   End
   Begin VB.Label Label5 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 5.Gender :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 4.Father’s/Husband’s Name:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   4080
      Width           =   8655
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   8655
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   8655
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 3.Name of The Guardian :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 2.Name of  The Person :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   8655
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   8655
   End
   Begin VB.Label Label21 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " 1.Registration Number :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   6855
   End
   Begin VB.Label Label20 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Details"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   6735
      Left            =   240
      Top             =   1680
      Width           =   16815
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "                                                                                REGISTRATION"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   19815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   -120
      Top             =   0
      Width           =   20535
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   17520
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      Height          =   2535
      Left            =   17400
      Top             =   960
      Width           =   2775
   End
End
Attribute VB_Name = "regmjd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim sr As New ADODB.Recordset
Dim str As String
Dim confirm As Integer
Dim Y As Integer

Private Sub btnupload_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
'Picture1.Picture = LoadPicture(str)
Image1.Picture = LoadPicture(str)
End Sub
Sub clear()
Text14.Text = ""
Text16.Text = ""
Text1.Text = ""
Text2.Text = ""
Option3.value = False
Option1.value = False
Text3.Text = ""
DTPicker1.value = "8/5/2018"
Text6.Text = ""
Text5.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Combo1.Text = "Select Council"
Combo2.Text = "Select Blood Group"
Text12.Text = ""
Text13.Text = ""
'Picture1.Picture = LoadPicture("")
Image1.Picture = LoadPicture("")
End Sub
Sub reload()
sr.Close
sr.Open "select * from MJDregister", con, adOpenDynamic, adLockPessimistic
End Sub
Sub display()
Text14.Text = sr!RegistrationNumber
Text16.Text = sr!Nameoftheperson
Text1.Text = sr!Nameoftheguardian
Text2.Text = sr!father / husbandname
If sr!Gender = "Male" Then
Option3.value = True
Else
Option1.value = True
End If
Text3.Text = sr!Age
DTPicker1.value = sr!DOB
Text6.Text = sr!Caste
Text5.Text = sr!Mobileno
Text7.Text = sr!Aadhaarno
Text8.Text = sr!email
Text9.Text = sr!PermanentAddress
Text10.Text = sr!NameofJamatkhana
Combo1.Text = sr!NameofCouncil
Combo2.Text = sr!BloodGroup
Text12.Text = sr!Height
Text13.Text = sr!AnyMajorHealthIssue
'Picture1.Picture = LoadPicture(sr!photo)
Image1.Picture = LoadPicture(sr!photo)
End Sub
Sub refreshdata()
sr.Close
sr.Open "select * from MJDregister", con, adOpenStatic, adLockPessimistic
If Not sr.EOF Then
sr.MoveNext
display
Else
MsgBox "No Record Found"
End If
End Sub

Private Sub btnexit_Click()
MDIForm1.Show
regmjd.Hide
End Sub

Private Sub btnsearch_Click()
sr.Close
sr.Open "select * from MJDregister where registrationnumber='" + Text14.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not sr.EOF Then
display
reload
Else
MsgBox "Record Profile not Found...!", vbInformation
End If
End Sub

Private Sub btnfirst_Click()
sr.MoveFirst
display
End Sub

Private Sub btnprevious_Click()
sr.MovePrevious
If sr.BOF Then
sr.MoveLast
display
Else
display
End If
End Sub

Private Sub btnnext_Click()
sr.MoveNext
If Not sr.EOF Then
display
Else
sr.MoveFirst
display
End If
End Sub

Private Sub btnlast_Click()
sr.MoveLast
display
End Sub

Private Sub btnupdate_Click()
If Message = True Then
Exit Sub
Else
sr.Fields("registrationnumber").value = Text14.Text
sr.Fields("Nameoftheperson").value = Text16.Text
sr.Fields("Nameoftheguardian").value = Text1.Text
sr.Fields("father/husbandname").value = Text2.Text
If Option3.value = True Then
sr.Fields("Gender").value = Option3.Caption
Else
sr.Fields("Gender").value = Option1.Caption
End If
sr.Fields("Age").value = Text3.Text
sr.Fields("DOB").value = DTPicker1.value
sr.Fields("Caste").value = Text6.Text
sr.Fields("Mobileno").value = Text5.Text
sr.Fields("Aadhaarno").value = Text7.Text
sr.Fields("Email").value = Text8.Text
sr.Fields("PermanentAddress").value = Text9.Text
sr.Fields("NameofJamatkhana").value = Text10.Text
sr.Fields("NameofCouncil").value = Combo1.Text
sr.Fields("BloodGroup").value = Combo2.Text
sr.Fields("Height").value = Text12.Text
sr.Fields("AnyMajorHealthIssue").value = Text13.Text
sr.Fields("Photo").value = str
MsgBox "Data is Updated successfully...!", vbInformation
sr.Update
End Sub

Private Sub btndelete_Click()
confirm = MsgBox("Do you want to Delete the student profile", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
sr.Delete adAffectCurrent
MsgBox "Record has been Deleted successfully", vbInformation, "message"
sr.Update
refreshdata
Else
MsgBox "Profile Not Deleted..!", vbInformation, "message"
End Sub

Private Sub btnsave_Click()
sr.Fields("registrationnumber").value = Text14.Text
sr.Fields("Nameoftheperson").value = Text16.Text
sr.Fields("Nameoftheguardian").value = Text1.Text
sr.Fields("father/husbandname").value = Text2.Text
If Option3.value = True Then
sr.Fields("Gender").value = Option3.Caption
Else
sr.Fields("Gender").value = Option1.Caption
End If
sr.Fields("Age").value = Text3.Text
sr.Fields("DOB").value = DTPicker1.value
sr.Fields("Caste").value = Text6.Text
sr.Fields("Mobileno").value = Text5.Text
sr.Fields("Aadhaarno").value = Text7.Text
sr.Fields("Email").value = Text8.Text
sr.Fields("PermanentAddress").value = Text9.Text
sr.Fields("NameofJamatkhana").value = Text10.Text
sr.Fields("NameofCouncil").value = Combo1.Text
sr.Fields("BloodGroup").value = Combo2.Text
sr.Fields("Height").value = Text12.Text
sr.Fields("AnyMajorHealthIssue").value = Text13.Text
sr.Fields("Photo").value = str
MsgBox "Your data is saved successfully... ", vbInformation
sr.Update
End Sub

Private Sub btnaddnew_Click()
sr.AddNew
clear
End Sub

Private Sub DTPicker1_Change()
Text4.Text = DTPicker1
End Sub

Private Sub Feesreceipt_Click()
DataReportMJDdetails.Show
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Desktop\ngo management system\ngo\database\login.mdb;Persist Security Info=False"
sr.Open " select * from MJDregister", con, adOpenDynamic, adLockPessimistic

Combo1.AddItem "Eastern council"
Combo1.AddItem "Western council"
Combo1.AddItem "Northern council"
Combo1.AddItem "Southern council"
Combo2.AddItem "A+"
Combo2.AddItem "A-"
Combo2.AddItem "B+"
Combo2.AddItem "B-"
Combo2.AddItem "AB+"
Combo2.AddItem "AB-"
Combo2.AddItem "O+"
Combo2.AddItem "O-"
'display
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Text3_Click()
Y = (DateValue(Date) - DateValue(Text4.Text)) / 365
Text3.Text = Y
End Sub

Private Sub Text5_LostFocus()
If Not IsNumeric(Text5.Text) Or Len(Trim(Text5.Text)) < 10 Then
    MsgBox "Invalid Mobile Number"
    Text5.Text = ""
    'Text5.SetFocus
    End If
End Sub



Private Sub Text7_LostFocus()
If Not IsNumeric(Text7.Text) Or Len(Trim(Text7.Text)) < 12 Then
    MsgBox "Aadhaar Number should be of 12 character"
    Text7.Text = ""
    'Text7.SetFocus
End If
End Sub



Private Sub Text8_LostFocus()
Dim myAt As Integer
Dim myDot As Integer
Dim myDotDot As Integer
isEmail = True
myAt = InStr(1, Text8.Text, "@", vbTextCompare)
myDot = InStr(myAt + 2, Text8.Text, ".", vbTextCompare)
myDotDot = InStr(myAt + 2, Text8.Text, "..", vbTextCompare)
If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(Text8.Text, 1) = "." Then
MsgBox ("Entered Email is Invalid!")
Text8.Text = ""
'Text8.SetFocus
End If
End Sub


Function Message() As Boolean
   
   If Text14.Text = "" Then
       MsgBox "Please Enter Registration Number"
       Text14.SetFocus
       Message = True
    ElseIf Text16.Text = "" Then
       MsgBox "Please Enter Name of the person"
       Text16.SetFocus
       Message = True
    ElseIf Text1.Text = "" Then
       MsgBox "Please Enter Name of the guardian"
       Text1.SetFocus
       Message = True
   ElseIf Text2.Text = "" Then
       MsgBox "Please Enter father/husband Name"
       Text2.SetFocus
       Message = True
   ElseIf Text6.Text = "" Then
       MsgBox "Please Enter caste"
       Text6.SetFocus
       Message = True
   ElseIf Text5.Text = "" Then
       MsgBox "Please Enter Mobile Number"
       Text5.SetFocus
       Message = True
   ElseIf Text7.Text = "" Then
       MsgBox "Please Enter Aadhaar card Number"
       Text7.SetFocus
       Message = True
  ElseIf Text8.Text = "" Then
       MsgBox "Please Enter the Email ID"
       Text8.SetFocus
       Message = True
   ElseIf Text9.Text = "" Then
       MsgBox "Please Enter Address"
       Text9.SetFocus
       Message = True
   ElseIf Text10.Text = "" Then
       MsgBox "Please Enter the Name of jamatkhana"
       Text10.SetFocus
       Message = True
   ElseIf Combo1.Text = "" Then
       MsgBox "Please Enter Name of council"
       Combo1.SetFocus
       Message = True
    ElseIf Text13.Text = "" Then
       MsgBox "Please Enter Any major Health issue"
       Text13.SetFocus
       Message = True
    ElseIf Text12.Text = "" Then
       MsgBox "Please Enter Height"
       Text12.SetFocus
       Message = True
    ElseIf Combo2.Text = "" Then
       MsgBox "Please Enter Blood Group"
       Combo2.SetFocus
       Message = True
       End If
End Function
