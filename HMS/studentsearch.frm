VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form search 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin HMSLFMJD.jcbutton btnshowall 
      Height          =   615
      Left            =   12120
      TabIndex        =   14
      Top             =   8040
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      Caption         =   "SHOW ALL RECORDS"
      Picture         =   "studentsearch.frx":0000
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btndelete 
      Height          =   615
      Left            =   6720
      TabIndex        =   13
      Top             =   8040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      Caption         =   "DELETE RECORD"
      Picture         =   "studentsearch.frx":10DA
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnfouyearsearch 
      Height          =   855
      Left            =   10680
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      Caption         =   "FOURTH YEAR"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnthiyearsearch 
      Height          =   855
      Left            =   9000
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      Caption         =   "THIRD YEAR"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnsecyearsearch 
      Height          =   855
      Left            =   7440
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      Caption         =   "SECOND YEAR"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnfiryearsearch 
      Height          =   855
      Left            =   6120
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      Caption         =   "FIRST YEAR"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnhscsearch 
      Height          =   855
      Left            =   5040
      TabIndex        =   8
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      Caption         =   "HSC"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnsscsearch 
      Height          =   855
      Left            =   4080
      TabIndex        =   7
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      ButtonStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      Caption         =   "SSC"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnsearch 
      Height          =   855
      Left            =   18840
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
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
      BackColor       =   8421376
      Caption         =   "SEARCH"
      UseMaskCOlor    =   -1  'True
   End
   Begin HMSLFMJD.jcbutton btnexit 
      Height          =   615
      Left            =   18960
      TabIndex        =   5
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
      BackColor       =   8421376
      Caption         =   ""
      Picture         =   "studentsearch.frx":21B4
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15960
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   17400
      Top             =   9840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Connect         =   $"studentsearch.frx":328E
      OLEDBString     =   $"studentsearch.frx":331E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *  from educationdetails"
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
      Bindings        =   "studentsearch.frx":33AE
      Height          =   5175
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   19695
      _ExtentX        =   34740
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483635
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   19
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   9.75
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
   Begin VB.Shape Shape5 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   5535
      Left            =   240
      Top             =   2400
      Width           =   19935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Name / Registration                       Number"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   3
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   12360
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Applying  for wise                  information"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   0
      Top             =   960
      Width           =   21090
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   5880
      Picture         =   "studentsearch.frx":33C3
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "                                                              SEARCH FOR A STUDENT RECORD"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   19695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   20730
   End
End
Attribute VB_Name = "search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim confirm As Integer

Private Sub btnexit_Click()
SEARCH.Hide
MDIForm1.Show
End Sub

Private Sub btnfiryearsearch_Click()
Adodc1.RecordSource = "Select * from educationdetails  where AdmissionAppliedFor='first year'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnsecyearsearch_Click()
Adodc1.RecordSource = "Select * from educationdetails  where AdmissionAppliedFor='second year'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnthiyearsearch_Click()
Adodc1.RecordSource = "Select * from educationdetails  where AdmissionAppliedFor='third year'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnfouyearsearch_Click()
Adodc1.RecordSource = "Select * from educationdetails  where AdmissionAppliedFor='fourth year'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnsscsearch_Click()
Adodc1.RecordSource = "Select * from educationdetails  where AdmissionAppliedFor='ssc'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub


Private Sub btnhscsearch_Click()
Adodc1.RecordSource = "Select * from educationdetails  where AdmissionAppliedFor='hsc'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btnsearch_Click()
Adodc1.RecordSource = "Select * from educationdetails where RegistrationNumber='" + Text1.Text + "' or FullNameofTheStudent='" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "Record Not Found,Enter any other Registration Number or Name", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End Sub

Private Sub btnshowall_Click()
Adodc1.RecordSource = "Select * from educationdetails "
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub btndelete_Click()
confirm = MsgBox("Do you want to delete the Record", vbYesNo + vbExclamation, "Warning Message")
If confirm = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record Deleted Successfully", vbInformation, "Delete Record Confirmation"
Else
MsgBox "Record Not Deleted", vbInformation, "Record Not Deleted"
End If
End Sub


Private Sub DataGrid1_Click()

End Sub

Private Sub Form_Load()

End Sub

Private Sub jcbutton5_Click()

End Sub
