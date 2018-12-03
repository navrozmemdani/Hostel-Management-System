VERSION 5.00
Begin VB.Form openingfrm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   20490
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "openingfrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin HMSLFMJD.mm_circle_progressbar mm_circle_progressbar1 
      Height          =   2175
      Left            =   9960
      TabIndex        =   4
      Top             =   7920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3836
      value           =   45
      ProgressbarColor=   192
      progressbar_frontof_color=   16777215
      BackColor       =   16777215
      ProgressBackcolor=   8421376
      prograssInactiveZone=   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Rockwell"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer3 
      Interval        =   900
      Left            =   13320
      Top             =   9600
   End
   Begin VB.Timer Timer2 
      Interval        =   900
      Left            =   13560
      Top             =   10320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12360
      Top             =   7920
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING...."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   10320
      TabIndex        =   3
      Top             =   10200
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING...."
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   10200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   20685
      Left            =   -360
      Picture         =   "openingfrm.frx":000C
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   21765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "KSA Company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   9420
      TabIndex        =   1
      Top             =   5205
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "L O A D I N G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   9000
      TabIndex        =   0
      Top             =   6240
      Width           =   2145
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   3375
   End
End
Attribute VB_Name = "openingfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim i As Integer
Private Sub Form_Load()
mm_circle_progressbar1.Advanced_Style_Animation = Not mm_circle_progressbar1.Advanced_Style_Animation


Timer1.Enabled = True
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub


Private Sub Timer1_Timer()
i = i + 1
If i > 101 Then i = 0 '

mm_circle_progressbar1.value = i
If i = 101 Then
  '  End
  Unload Me
   adminlogin.Show
  Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Label4.Visible = False
Label2.Visible = True
End Sub

Private Sub Timer3_Timer()
Label4.Visible = True
Label2.Visible = False
End Sub
