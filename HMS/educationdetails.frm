VERSION 5.00
Begin VB.Form details 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin HMSLFMJD.jcbutton lastbtn 
      Height          =   735
      Left            =   17880
      TabIndex        =   55
      Top             =   8280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      ButtonStyle     =   7
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
      Caption         =   "LAST"
      Picture         =   "educationdetails.frx":0000
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton nextbtn 
      Height          =   735
      Left            =   17880
      TabIndex        =   54
      Top             =   7320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      ButtonStyle     =   7
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
      Caption         =   "NEXT"
      Picture         =   "educationdetails.frx":10DA
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton previousbtn 
      Height          =   735
      Left            =   17880
      TabIndex        =   53
      Top             =   6480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      ButtonStyle     =   7
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
      Caption         =   "PREVIOUS"
      Picture         =   "educationdetails.frx":21B4
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton firstbtn 
      Height          =   735
      Left            =   17880
      TabIndex        =   52
      Top             =   5640
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      ButtonStyle     =   7
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
      Caption         =   "FIRST"
      Picture         =   "educationdetails.frx":328E
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton delrecbtn 
      Height          =   735
      Left            =   17880
      TabIndex        =   51
      Top             =   4680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      ButtonStyle     =   4
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
      Caption         =   "DELETE"
      Picture         =   "educationdetails.frx":4368
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton updrecbtn 
      Height          =   735
      Left            =   17880
      TabIndex        =   50
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      ButtonStyle     =   4
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
      Caption         =   "UPDATE"
      Picture         =   "educationdetails.frx":5442
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton addrecbtn 
      Height          =   705
      Left            =   17880
      TabIndex        =   48
      Top             =   2160
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1244
      ButtonStyle     =   4
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
      Caption         =   "ADD NEW"
      Picture         =   "educationdetails.frx":651C
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton detailsprint 
      Height          =   735
      Left            =   18480
      TabIndex        =   47
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
      ButtonStyle     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      Caption         =   ""
      Picture         =   "educationdetails.frx":75F6
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   8880
      Width           =   4815
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      MultiLine       =   -1  'True
      TabIndex        =   45
      Top             =   8400
      Width           =   4815
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   7920
      Width           =   4815
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      MultiLine       =   -1  'True
      TabIndex        =   43
      Top             =   7440
      Width           =   4815
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      MultiLine       =   -1  'True
      TabIndex        =   42
      Top             =   6960
      Width           =   4815
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      MultiLine       =   -1  'True
      TabIndex        =   41
      Top             =   6480
      Width           =   4815
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   40
      Top             =   8880
      Width           =   2295
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   39
      Top             =   8400
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   7920
      Width           =   2295
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   7440
      Width           =   2295
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   36
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   35
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   8880
      Width           =   3255
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   8400
      Width           =   3255
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   7920
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   7440
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   6960
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   6480
      Width           =   3255
   End
   Begin VB.TextBox Text23 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   4920
      Width           =   13455
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   4440
      Width           =   13455
   End
   Begin VB.TextBox Text22 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   3960
      Width           =   13455
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   3480
      Width           =   13455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   3000
      Width           =   13455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   2520
      Width           =   13455
   End
   Begin VB.TextBox Text26 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2040
      Width           =   13455
   End
   Begin VB.TextBox Text25 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1560
      Width           =   13455
   End
   Begin HMSLFMJD.jcbutton saverecbtn 
      Height          =   705
      Left            =   17880
      TabIndex        =   49
      Top             =   3000
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1244
      ButtonStyle     =   4
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
      Caption         =   "SAVE"
      Picture         =   "educationdetails.frx":86D0
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   5040
      Picture         =   "educationdetails.frx":97AA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "                              4th YEAR"
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
      Left            =   360
      TabIndex        =   28
      Top             =   8880
      Width           =   3855
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "                              3rd YEAR"
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
      Left            =   360
      TabIndex        =   27
      Top             =   8400
      Width           =   3855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "                               2nd YEAR"
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
      Left            =   360
      TabIndex        =   26
      Top             =   7920
      Width           =   3855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "                                 1st YEAR"
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
      Left            =   360
      TabIndex        =   25
      Top             =   7440
      Width           =   3855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "                                      HSC"
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
      Left            =   360
      TabIndex        =   24
      Top             =   6960
      Width           =   3855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "                                       SSC"
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
      Left            =   360
      TabIndex        =   23
      Top             =   6480
      Width           =   3855
   End
   Begin VB.Shape Shape19 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   8880
      Width           =   17175
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   7440
      Width           =   17175
   End
   Begin VB.Shape Shape17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   7920
      Width           =   17175
   End
   Begin VB.Shape Shape16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   8400
      Width           =   17175
   End
   Begin VB.Shape Shape15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   17175
   End
   Begin VB.Shape Shape14 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   17175
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "               School / College / Name"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   22
      Top             =   5880
      Width           =   4815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "              Year"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   21
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "                Percentage"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "                    Particulars"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   19
      Top             =   5880
      Width           =   3855
   End
   Begin VB.Shape Shape13 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   17175
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   3375
      Left            =   240
      Top             =   5880
      Width           =   17415
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Past Year Track Record :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   5520
      Width           =   3735
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   240
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " 6. Future Goals  :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "  7. Current Institution :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "  8. Co-Curricular Activities :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3960
      Width           =   17415
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   17415
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   17415
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "4. Academic Medium :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " 5. Admission Applied for  :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3000
      Width           =   17415
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   17415
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "3. Last Appeared Exam :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   17415
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " 2. Full Name of The Student :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   17415
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   " 1.Registration Number :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   17415
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "        Education Details"
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
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "                                                        ACADEMIC / PROFESSIONAL DETAILS"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
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
      Width           =   19935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   -600
      Top             =   0
      Width           =   21090
   End
End
Attribute VB_Name = "details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub addrecbtn_Click()
rs.AddNew
clear
End Sub
Sub display()
Text25.Text = rs!RegistrationNumber
Text26.Text = rs!FullNameofTheStudent
Text1.Text = rs!LastAppearedExam
Text2.Text = rs!AcademicMedium
Text3.Text = rs!AdmissionAppliedFor
Text22.Text = rs!FutureGoals
Text24.Text = rs!CurrentInstitution
Text23.Text = rs!CoCurricularActivities
Text4.Text = rs!SSCPercentage
Text10.Text = rs!SSCYear
Text16.Text = rs!SSCSchool
Text5.Text = rs!HSCPercentage
Text11.Text = rs!HSCYear
Text17.Text = rs!HSCCollegeName
Text6.Text = rs!FirstYearPercentage
Text12.Text = rs!FirstYearYear
Text18.Text = rs!FirstYearCollegeName
Text7.Text = rs!SecondYearPercentage
Text13.Text = rs!SecondYearYear
Text19.Text = rs!SecondYearCollegeName
Text8.Text = rs!ThirdYearPercentage
Text14.Text = rs!ThirdYearYear
Text20.Text = rs!ThirdYearCollegeName
Text9.Text = rs!ForthYearPercentage
Text15.Text = rs!ForthYearYear
Text21.Text = rs!ForthYearCollegeName
End Sub



Private Sub delrecbtn_Click()
confirm = MsgBox("Do you want to Delete the student educational details", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox "Record has been Deleted successfully", vbInformation, "message"
rs.Update
refreshdata
Else
MsgBox "Profile Not Deleted..!", vbInformation, "message"
End If
End Sub

Private Sub firstbtn_Click()
rs.MoveFirst
display
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Desktop\ngo management system\ngo\database\login.mdb;Persist Security Info=False"
rs.Open " select * from educationdetails", con, adOpenDynamic, adLockPessimistic
'display
End Sub

Sub clear()
Text25.Text = ""
Text26.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text22.Text = ""
Text24.Text = ""
Text23.Text = ""
Text4.Text = ""
Text10.Text = ""
Text16.Text = ""
Text5.Text = ""
Text11.Text = ""
Text17.Text = ""
Text6.Text = ""
Text12.Text = ""
Text18.Text = ""
Text7.Text = ""
Text13.Text = ""
Text19.Text = ""
Text8.Text = ""
Text14.Text = ""
Text20.Text = ""
Text9.Text = ""
Text15.Text = ""
Text21.Text = ""
End Sub
Sub reload()
rs.Close
rs.Open "select * from educationdetails", con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label20_Click()

End Sub

Private Sub lastbtn_Click()
rs.MoveLast
display
End Sub

Private Sub nextbtn_Click()
rs.MoveNext
If Not rs.EOF Then
display
Else
rs.MoveFirst
display
End If
End Sub

Private Sub previousbtn_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Else
display
End If
End Sub

Private Sub detailsprint_Click()
DataReporteducationaldetails.Show
End Sub

Private Sub saverecbtn_Click()
rs.Fields("RegistrationNumber").value = Text25.Text
rs.Fields("FullNameofTheStudent").value = Text26.Text
rs.Fields("LastAppearedExam").value = Text1.Text
rs.Fields("AcademicMedium").value = Text2.Text
rs.Fields("AdmissionAppliedFor").value = Text3.Text
rs.Fields("FutureGoals").value = Text22.Text
rs.Fields("CurrentInstitution").value = Text24.Text
rs.Fields("CoCurricularActivities").value = Text23.Text
rs.Fields("SSCPercentage").value = Text4.Text
rs.Fields("SSCYear").value = Text10.Text
rs.Fields("SSCSchool").value = Text16.Text
rs.Fields("HSCPercentage").value = Text5.Text
rs.Fields("HSCYear").value = Text11.Text
rs.Fields("HSCCollegeName").value = Text17.Text
rs.Fields("FirstYearPercentage").value = Text6.Text
rs.Fields("FirstYearYear").value = Text12.Text
rs.Fields("FirstYearCollegeName").value = Text18.Text
rs.Fields("SecondYearPercentage").value = Text7.Text
rs.Fields("SecondYearYear").value = Text13.Text
rs.Fields("SecondYearCollegeName").value = Text19.Text
rs.Fields("ThirdYearPercentage").value = Text8.Text
rs.Fields("ThirdYearYear").value = Text14.Text
rs.Fields("ThirdYearCollegeName").value = Text20.Text
rs.Fields("ForthYearPercentage").value = Text9.Text
rs.Fields("ForthYearYear").value = Text15.Text
rs.Fields("ForthYearCollegeName").value = Text21.Text
MsgBox "Data is saved successfully...!", vbInformation
rs.Update
End Sub

Private Sub updrecbtn_Click()
rs.Fields("RegistrationNumber").value = Text25.Text
rs.Fields("FullNameofTheStudent").value = Text26.Text
rs.Fields("LastAppearedExam").value = Text1.Text
rs.Fields("AcademicMedium").value = Text2.Text
rs.Fields("AdmissionAppliedFor").value = Text3.Text
rs.Fields("FutureGoals").value = Text22.Text
rs.Fields("CurrentInstitution").value = Text24.Text
rs.Fields("CoCurricularActivities").value = Text23.Text
rs.Fields("SSCPercentage").value = Text4.Text
rs.Fields("SSCYear").value = Text10.Text
rs.Fields("SSCSchool").value = Text16.Text
rs.Fields("HSCPercentage").value = Text5.Text
rs.Fields("HSCYear").value = Text11.Text
rs.Fields("HSCCollegeName").value = Text17.Text
rs.Fields("FirstYearPercentage").value = Text6.Text
rs.Fields("FirstYearYear").value = Text12.Text
rs.Fields("FirstYearCollegeName").value = Text18.Text
rs.Fields("SecondYearPercentage").value = Text7.Text
rs.Fields("SecondYearYear").value = Text13.Text
rs.Fields("SecondYearCollegeName").value = Text19.Text
rs.Fields("ThirdYearPercentage").value = Text8.Text
rs.Fields("ThirdYearYear").value = Text14.Text
rs.Fields("ThirdYearCollegeName").value = Text20.Text
rs.Fields("ForthYearPercentage").value = Text9.Text
rs.Fields("ForthYearYear").value = Text15.Text
rs.Fields("ForthYearCollegeName").value = Text21.Text
MsgBox "Data is Updated successfully...!", vbInformation
rs.Update
End Sub


Sub refreshdata()
rs.Close
rs.Open "select * from educationdetails", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "No Record Found"
End If
End Sub
