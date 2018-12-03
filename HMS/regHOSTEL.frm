VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form reghostel 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   9675
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   FillColor       =   &H80000000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   9960
      TabIndex        =   57
      Top             =   3000
      Width           =   5655
      Begin HMSLFMJD.jcbutton Feesreceipt 
         Height          =   615
         Left            =   4440
         TabIndex        =   69
         Top             =   3840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
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
         BackColor       =   16711680
         Caption         =   ""
         Picture         =   "regHOSTEL.frx":0000
         UseMaskCOlor    =   -1  'True
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   64
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   61
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   59
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "PRINT FEES    RECEIPT :"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   240
         TabIndex        =   68
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   66
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "4.Fees Remaining  :"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   65
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Shape Shape30 
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   3000
         Width           =   4695
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C000&
         BackStyle       =   0  'Transparent
         Caption         =   "3. Fees Paid :"
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
         Left            =   240
         TabIndex        =   62
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Shape Shape29 
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1560
         Width           =   5415
      End
      Begin VB.Shape Shape28 
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   2280
         Width           =   5415
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "2. Total Fees Alloted to       the Student"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   60
         Top             =   840
         Width           =   2055
      End
      Begin VB.Shape Shape27 
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "1.Total Fees of Hostel"
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
         Left            =   240
         TabIndex        =   58
         Top             =   240
         Width           =   1935
      End
      Begin VB.Shape Shape24 
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5415
      End
   End
   Begin HMSLFMJD.jcbutton btnunload 
      Height          =   495
      Left            =   16080
      TabIndex        =   56
      Top             =   8040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      ButtonStyle     =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16765357
      Caption         =   ""
      Picture         =   "regHOSTEL.frx":10DA
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnload 
      Height          =   495
      Left            =   15360
      TabIndex        =   55
      Top             =   8040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      ButtonStyle     =   10
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
      Picture         =   "regHOSTEL.frx":21B4
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text15 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   405
      Left            =   10920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   52
      Top             =   7440
      Width           =   6135
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   10920
      MultiLine       =   -1  'True
      TabIndex        =   50
      Top             =   6840
      Width           =   6135
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10920
      TabIndex        =   48
      Text            =   "Select Blood Group"
      Top             =   6240
      Width           =   6135
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   11160
      MultiLine       =   -1  'True
      TabIndex        =   46
      Top             =   5040
      Width           =   5775
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   11160
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   4320
      Width           =   5775
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   11160
      MultiLine       =   -1  'True
      TabIndex        =   42
      Top             =   3600
      Width           =   5775
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   11160
      MultiLine       =   -1  'True
      TabIndex        =   40
      Top             =   2880
      Width           =   5775
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   11160
      MultiLine       =   -1  'True
      TabIndex        =   38
      Top             =   2160
      Width           =   5775
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11160
      TabIndex        =   36
      Text            =   "Select Council"
      Top             =   1440
      Width           =   5775
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   7920
      Width           =   5655
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2760
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Top             =   7200
      Width           =   5655
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   525
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   6480
      Width           =   5655
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   5760
      Width           =   5295
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   5040
      Width           =   5655
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   4320
      Width           =   5655
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Female"
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
      Left            =   5760
      TabIndex        =   21
      Top             =   3720
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Male"
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
      Left            =   2760
      TabIndex        =   20
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   2880
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   2160
      Width           =   5655
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   1440
      Width           =   4095
   End
   Begin HMSLFMJD.jcbutton SEARCH 
      Height          =   615
      Left            =   7080
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
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
      BackColor       =   12582912
      Caption         =   "SEARCH"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":328E
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnlast 
      Height          =   615
      Left            =   19200
      TabIndex        =   10
      Top             =   8040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      ButtonStyle     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388608
      Caption         =   "LAST"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":4368
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnnext 
      Height          =   615
      Left            =   17520
      TabIndex        =   9
      Top             =   8040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ButtonStyle     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388608
      Caption         =   "NEXT"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":5442
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnprevious 
      Height          =   615
      Left            =   18840
      TabIndex        =   8
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ButtonStyle     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388608
      Caption         =   "PREVIOUS"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":651C
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnfirst 
      Height          =   615
      Left            =   17520
      TabIndex        =   7
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      ButtonStyle     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8388608
      Caption         =   "FIRST"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":75F6
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnupdate 
      Height          =   495
      Left            =   18000
      TabIndex        =   2
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      BackColor       =   8388608
      Caption         =   "UPDATE"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":86D0
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnupload 
      Height          =   495
      Left            =   17760
      TabIndex        =   1
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
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
      BackColor       =   8388608
      Caption         =   "UPLOAD PHOTO"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":97AA
      UseMaskCOlor    =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   17160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin HMSLFMJD.jcbutton btnsave 
      Height          =   495
      Left            =   18000
      TabIndex        =   3
      Top             =   5880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      BackColor       =   8388608
      Caption         =   "SAVE"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":A884
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btndelete 
      Height          =   495
      Left            =   18000
      TabIndex        =   4
      Top             =   4680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      BackColor       =   8388608
      Caption         =   "DELETE"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":B95E
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnexit 
      Height          =   495
      Left            =   18000
      TabIndex        =   5
      Top             =   6480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      BackColor       =   8388608
      Caption         =   "EXIT"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":CA38
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnaddnew 
      Height          =   495
      Left            =   18000
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
      BackColor       =   8388608
      Caption         =   "ADD NEW"
      ForeColor       =   16777215
      Picture         =   "regHOSTEL.frx":DB12
      UseMaskCOlor    =   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   8040
      TabIndex        =   28
      Top             =   5760
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      _Version        =   393216
      Format          =   110886913
      CurrentDate     =   43317
   End
   Begin VB.Label Label27 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Details"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   9360
      TabIndex        =   54
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Shape Shape26 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   9000
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Label Label26 
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
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   9000
      TabIndex        =   53
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Shape Shape25 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   8880
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "3. Any Major   Problem :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   51
      Top             =   7440
      Width           =   1935
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "2. height (In feet's)"
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
      Left            =   9000
      TabIndex        =   49
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1. Blood Group :"
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
      Left            =   9000
      TabIndex        =   47
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "16. Family Income p.a :"
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
      Left            =   8880
      TabIndex        =   45
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "15. Last Institution  Attended :"
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
      Left            =   8880
      TabIndex        =   43
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "14. Student Email ID :"
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
      Left            =   8880
      TabIndex        =   41
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "13. Student Contact no :"
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
      Left            =   8880
      TabIndex        =   39
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "12. Parent Contact no :"
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
      Left            =   8880
      TabIndex        =   37
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "11. Name of Council :"
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
      Left            =   8880
      TabIndex        =   35
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Shape Shape23 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   8880
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   8295
   End
   Begin VB.Shape Shape22 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   8880
      Shape           =   4  'Rounded Rectangle
      Top             =   7440
      Width           =   8295
   End
   Begin VB.Shape Shape21 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   8295
   End
   Begin VB.Shape Shape20 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   8295
   End
   Begin VB.Shape Shape19 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   8295
   End
   Begin VB.Shape Shape18 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   8295
   End
   Begin VB.Shape Shape17 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   8295
   End
   Begin VB.Shape Shape16 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   8880
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   8295
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "10. Name of Jamatkhana :"
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
      TabIndex        =   33
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "9. Permanent Address :"
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
      TabIndex        =   31
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "8. Age :"
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
      TabIndex        =   29
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "7. Date of Birth :"
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
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "6. Place of Birth :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "5. Caste (with Sub-Caste) :"
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
      TabIndex        =   22
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "4.Gender:"
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
      TabIndex        =   19
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "3. Name of the            Parents/Guardian :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "2. Full Name of the       Student :"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "1.Registration Number :"
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
      TabIndex        =   13
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Shape Shape15 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   7920
      Width           =   8295
   End
   Begin VB.Shape Shape14 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   8295
   End
   Begin VB.Shape Shape13 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   8295
   End
   Begin VB.Shape Shape12 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   8295
   End
   Begin VB.Shape Shape11 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   8295
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   8295
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   8295
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   8295
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   7200
      Width           =   8295
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   8295
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   6735
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   7335
      Left            =   120
      Top             =   1320
      Width           =   17175
   End
   Begin VB.Label Label25 
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
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   17880
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   2535
      Left            =   17760
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "                                                                                 REGISTRATION"
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
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20055
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   -240
      Top             =   0
      Width           =   20895
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   120
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "reghostel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim confirm As Integer
Dim Y As Integer

Sub clear()
Text19.Text = ""
Text1.Text = ""
Text2.Text = ""
Option1.value = False
Option2.value = False
Text3.Text = ""
Text4.Text = ""
DTPicker1.value = "10/03/2018"
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo1.Text = "Select Council"
Text9.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Combo2.Text = "Select Blood Group"
Text18.Text = ""
Text15.Text = ""
Text16.Text = ""
Text20.Text = ""
Text21.Text = ""
'Picture1.Picture = LoadPicture("")
Image1.Picture = LoadPicture("")
End Sub

Private Sub btnaddnew_Click()
rs.AddNew
clear
End Sub

Private Sub btndelete_Click()
confirm = MsgBox("Do you want to Delete the student profile", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox "Record has been Deleted successfully", vbInformation, "message"
rs.Update
refreshdata
Else
MsgBox "Profile Not Deleted..!", vbInformation, "message"
End If
End Sub

Private Sub btnexit_Click()
MDIForm1.Show
reghostel.Hide
End Sub

Private Sub btnsave_Click()
If Message = True Then
Exit Sub
Else
rs.Fields("registrationnumber").value = Text19.Text
rs.Fields("NameoftheStudent").value = Text1.Text
rs.Fields("NameoftheParent").value = Text2.Text
If Option1.value = True Then
rs.Fields("Gender") = Option1.Caption
Else
rs.Fields("Gender") = Option2.Caption
End If
rs.Fields("Caste").value = Text3.Text
rs.Fields("PlaceofBirth").value = Text4.Text
rs.Fields("DOB").value = DTPicker1.value
rs.Fields("Age").value = Text5.Text
rs.Fields("PermanentAddress").value = Text6.Text
rs.Fields("NameofJamatkhana").value = Text7.Text
rs.Fields("Nameofcouncil").value = Combo1.Text
rs.Fields("StudentContactNO").value = Text9.Text
rs.Fields("StudentEmailID").value = Text11.Text
rs.Fields("ParentContactNO").value = Text12.Text
rs.Fields("LastInstitutionAttended").value = Text13.Text
rs.Fields("FamilyIncome").value = Text14.Text
rs.Fields("BloodGroup").value = Combo2.Text
rs.Fields("Height").value = Text18.Text
rs.Fields("AnyMajorHealthIssue").value = Text15.Text
rs.Fields("TotalFeesofHostelAllotted").value = Text16.Text
rs.Fields("FeesPaid").value = Text20.Text
rs.Fields("FeesRemaining").value = Text21.Text
rs.Fields("Photo").value = str
MsgBox "Data is saved successfully ..!!!", vbInformation
rs.Update
End If
End Sub

Private Sub btnsearch_Click()
rs.Close
rs.Open "Select * from STUDENTregister where registrationnumber='" + Text19.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
display
reload
Else
MsgBox "Record Profile not found ..!!", vbInformation
End If
End Sub

Private Sub btnupdate_Click()
If Message = True Then
Exit Sub
Else
rs.Fields("registrationnumber").value = Text19.Text
rs.Fields("NameoftheStudent").value = Text1.Text
rs.Fields("NameoftheParent").value = Text2.Text
If Option1.value = True Then
rs.Fields("Gender") = Option1.Caption
Else
rs.Fields("Gender") = Option2.Caption
End If
rs.Fields("Caste").value = Text3.Text
rs.Fields("PlaceofBirth").value = Text4.Text
rs.Fields("DOB").value = DTPicker1.value
rs.Fields("Age").value = Text5.Text
rs.Fields("PermanentAddress").value = Text6.Text
rs.Fields("NameofJamatkhana").value = Text7.Text
rs.Fields("Nameofcouncil").value = Combo1.Text
rs.Fields("StudentContactNO").value = Text9.Text
rs.Fields("StudentEmailID").value = Text11.Text
rs.Fields("ParentContactNO").value = Text12.Text
rs.Fields("LastInstitutionAttended").value = Text13.Text
rs.Fields("FamilyIncome").value = Text14.Text
rs.Fields("BloodGroup").value = Combo2.Text
rs.Fields("Height").value = Text18.Text
rs.Fields("AnyMajorHealthIssue").value = Text15.Text
rs.Fields("TotalFeesofHostelAllotted").value = Text16.Text
rs.Fields("FeesPaid").value = Text20.Text
rs.Fields("FeesRemaining").value = Text21.Text
MsgBox "Data is updated successfully ..!!!", vbInformation
rs.Update
End If
End Sub


Function Message() As Boolean
   
   If Text19.Text = "" Then
       MsgBox "Please Enter Registration Number"
       Text19.SetFocus
       Message = True
    ElseIf Text1.Text = "" Then
       MsgBox "Please Enter Name of the Student"
       Text1.SetFocus
       Message = True
    ElseIf Text2.Text = "" Then
       MsgBox "Please Enter Name of the Parent"
       Text2.SetFocus
       Message = True
   ElseIf Text3.Text = "" Then
       MsgBox "Please Enter Caste"
       Text3.SetFocus
       Message = True
    ElseIf Text4.Text = "" Then
       MsgBox "Please Enter Place of Birth"
       Text4.SetFocus
       Message = True
   ElseIf Text5.Text = "" Then
       MsgBox "Please Enter Age"
       Text5.SetFocus
       Message = True
   ElseIf Text6.Text = "" Then
       MsgBox "Please Enter Permanent Address"
       Text6.SetFocus
       Message = True
   ElseIf Text7.Text = "" Then
       MsgBox "Please Enter Name of Jamatkhana"
       Text7.SetFocus
       Message = True
   ElseIf Combo1.Text = "" Then
       MsgBox "Please Enter Name of council"
       Combo1.SetFocus
       Message = True
    ElseIf Text13.Text = "" Then
       MsgBox "Please Enter Last Institution Attended"
       Text13.SetFocus
       Message = True
    ElseIf Text14.Text = "" Then
       MsgBox "Please Enter Family Income"
       Text14.SetFocus
       Message = True
    ElseIf Combo2.Text = "" Then
       MsgBox "Please Enter Blood Group"
       Combo2.SetFocus
       Message = True
    ElseIf Text18.Text = "" Then
       MsgBox "Please Enter Height"
       Text18.SetFocus
       Message = True
    ElseIf Text15.Text = "" Then
       MsgBox "Please Enter Any Major Health Issue"
       Text15.SetFocus
       Message = True
    ElseIf Text16.Text = "" Then
       MsgBox "Please Enter Total Fees of Hostel Allotted"
       Text16.SetFocus
       Message = True
    ElseIf Text20.Text = "" Then
       MsgBox "Please Enter FeesPaid"
       Text20.SetFocus
       Message = True
    ElseIf Text21.Text = "" Then
       MsgBox "Please Enter Fees Remaining"
       Text21.SetFocus
       Message = True
       End If
End Function
Private Sub btnupload_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
'Picture1.Picture = LoadPicture(str)
Image1.Picture = LoadPicture(str)
End Sub

Sub reload()
rs.Close
rs.Open "select * from STUDENTregister", con, adOpenDynamic, adLockPessimistic
End Sub
Sub display()
Text19.Text = rs!RegistrationNumber
Text1.Text = rs!NameoftheStudent
Text2.Text = rs!NameoftheParent
If rs!Gender = "MALE" Then
Option1.value = True
Else
Option2.value = True
End If
Text3.Text = rs!Caste
Text4.Text = rs!PlaceofBirth
DTPicker1.value = rs!DOB
Text5.Text = rs!Age
Text6.Text = rs!PlaceofBirth
Text7.Text = rs!NameofJamatkhana
Combo1.Text = rs!NameofCouncil
Text9.Text = rs!StudentContactNO
Text11.Text = rs!StudentEmailID
Text12.Text = rs!ParentContactNO
Text13.Text = rs!LastInstitutionAttended
Text14.Text = rs!FamilyIncome
Combo2.Text = rs!BloodGroup
Text18.Text = rs!Height
Text15.Text = rs!AnyMajorHealthIssue
Text16.Text = rs!TotalFeesofHostelAllotted
Text20.Text = rs!FeesPaid
Text21.Text = rs!FeesRemaining
'Picture1.Picture = LoadPicture(rs!photo)
Image1.Picture = LoadPicture(rs!photo)
End Sub


Sub refreshdata()
rs.Close
rs.Open "select * from STUDENTregister", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "No Record Found"
End If
End Sub


Private Sub btnfirst_Click()
rs.MoveFirst
display
End Sub

Private Sub btnlast_Click()
rs.MoveLast
display
End Sub

Private Sub btnnext_Click()
rs.MoveNext
If Not rs.EOF Then
display
Else
rs.MoveFirst
display
End If
End Sub

Private Sub btnprevious_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
display
Else
display
End If
End Sub

Private Sub DTPicker1_Change()
Text8.Text = DTPicker1
End Sub

Private Sub Feesreceipt_Click()
DataReportfeesReceipt.Show
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Desktop\ngo management system\ngo\database\login.mdb;Persist Security Info=False"
rs.Open " select * from STUDENTregister", con, adOpenDynamic, adLockPessimistic
Combo1.AddItem " Eastern council"
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

Private Sub Label18_Click()

End Sub

Private Sub Text11_LostFocus()
If Not IsNumeric(Text11.Text) Or Len(Trim(Text11.Text)) < 10 Then
    MsgBox "Invalid Mobile Number"
    Text11.Text = ""
    'Text11.SetFocus
    End If
End Sub

Private Sub Text12_LostFocus()
Dim myAt As Integer
Dim myDot As Integer
Dim myDotDot As Integer
isEmail = True
myAt = InStr(1, Text12.Text, "@", vbTextCompare)
myDot = InStr(myAt + 2, Text12.Text, ".", vbTextCompare)
myDotDot = InStr(myAt + 2, Text12.Text, "..", vbTextCompare)
If myAt = 0 Or myDot = 0 Or Not myDotDot = 0 Or Right(Text12.Text, 1) = "." Then
MsgBox ("Entered Email is Invalid!")
Text12.Text = ""
'Text12.SetFocus
End If
End Sub

Private Sub Text16_Change()

End Sub

Private Sub Text20_Change()

End Sub

Private Sub Text5_Click()
Y = (DateValue(Date) - DateValue(Text8.Text)) / 365
Text5.Text = Y
End Sub

Private Sub Text9_LostFocus()
If Not IsNumeric(Text9.Text) Or Len(Trim(Text9.Text)) < 10 Then
    MsgBox "Invalid Mobile Number"
    Text9.Text = ""
    'Text9.SetFocus
End If
End Sub
