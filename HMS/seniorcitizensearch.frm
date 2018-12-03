VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form mjdsearch 
   BackColor       =   &H000000FF&
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
   Begin HMSLFMJD.jcbutton jcbutton9 
      Height          =   615
      Left            =   11280
      TabIndex        =   15
      Top             =   8160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      ButtonStyle     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   128
      Caption         =   "SHOW ALL RECORD"
      Picture         =   "seniorcitizensearch.frx":0000
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btndelete 
      Height          =   615
      Left            =   6120
      TabIndex        =   14
      Top             =   8160
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      ButtonStyle     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   128
      Caption         =   "DELETE RECORD "
      Picture         =   "seniorcitizensearch.frx":10DA
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnsearch 
      Height          =   615
      Left            =   11040
      TabIndex        =   13
      Top             =   2160
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
      BackColor       =   128
      Caption         =   "SEARCH"
      Picture         =   "seniorcitizensearch.frx":21B4
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnsouthsearch 
      Height          =   615
      Left            =   17520
      TabIndex        =   12
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
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
      BackColor       =   128
      Caption         =   "SOUTHERN COUNCIL"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnnorthsearch 
      Height          =   615
      Left            =   14880
      TabIndex        =   11
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
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
      BackColor       =   128
      Caption         =   "NORTHERN COUNCIL"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnwestsearch 
      Height          =   615
      Left            =   12480
      TabIndex        =   10
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
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
      BackColor       =   128
      Caption         =   "WESTERN COUNCIL"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btneastsearch 
      Height          =   615
      Left            =   10080
      TabIndex        =   9
      Top             =   1200
      Width           =   2175
      _ExtentX        =   3836
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
      BackColor       =   128
      Caption         =   "EASTERN COUNCIL"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnfemalesearch 
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
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
      BackColor       =   128
      Caption         =   "FEMALE"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnmalesearch 
      Height          =   615
      Left            =   3600
      TabIndex        =   7
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
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
      BackColor       =   128
      Caption         =   "MALE"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnexit 
      Height          =   615
      Left            =   18720
      TabIndex        =   6
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
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
      BackColor       =   128
      Caption         =   ""
      Picture         =   "seniorcitizensearch.frx":328E
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   4815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   14880
      Top             =   9720
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Connect         =   $"seniorcitizensearch.frx":4368
      OLEDBString     =   $"seniorcitizensearch.frx":43F8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from MJDregister"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "seniorcitizensearch.frx":4488
      Height          =   4815
      Left            =   360
      TabIndex        =   1
      Top             =   3120
      Width           =   19695
      _ExtentX        =   34740
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   192
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   25
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   5055
      Left            =   240
      Top             =   3000
      Width           =   19935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Name / Registration Number"
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
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   -360
      Top             =   2040
      Width           =   13095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Record By Name            of the Council"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Record By Gender"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   -120
      Top             =   1080
      Width           =   20655
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   5040
      Picture         =   "seniorcitizensearch.frx":449D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000C0&
      Caption         =   "                                                         SEARCH FOR A SENIOR CITIZEN RECORD"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   -120
      Top             =   0
      Width           =   21015
   End
End
Attribute VB_Name = "mjdsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim confirm As Integer
Private Sub btndelete_Click()
confirm = MsgBox("Do you want to delete the Record", vbYesNo + vbExclamation, "Warning Message")
If confirm = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record Deleted Successfully", vbInformation, "Delete Record Confirmation"
Else
MsgBox "Record Not Deleted", vbInformation, "Record Not Deleted"
End If
End Sub

Private Sub btneastsearch_Click()
Adodc1.RecordSource = "Select * from MJDregister  where NameofCouncil='Eastern council'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnexit_Click()
jdsearch.Hide
MDIForm1.Show
End Sub

Private Sub btnfemalesearch_Click()
Adodc1.RecordSource = "Select * from MJDregister  where Gender='female'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnmalesearch_Click()
Adodc1.RecordSource = "Select * from MJDregister  where Gender='male'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnnorthsearch_Click()
Adodc1.RecordSource = "Select * from MJDregister  where NameofCouncil='Northern council'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnsearch_Click()
Adodc1.RecordSource = "Select * from MJDregister where registrationnumber='" + Text1.Text + "' or Nameoftheperson='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Record Not Found,Enter any other Registration Number or Name", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub btnshowall_Click()
Adodc1.RecordSource = "Select * from MJDregister "
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnsscsearch_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub btnsouthsearch_Click()
Adodc1.RecordSource = "Select * from MJDregister  where NameofCouncil='Southern council'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnwestsearch_Click()
Adodc1.RecordSource = "Select * from MJDregister  where NameofCouncil='Western council'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

