VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   540
   ClientWidth     =   20250
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIForm1.frx":10CA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   1588
      ButtonWidth     =   3704
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "STUDENT DETAILS"
            Object.ToolTipText     =   "Student Admission Registerstion"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SENIOR CITIZEN "
            Object.ToolTipText     =   "Senior CitizenRegisterstion"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EDUCATION"
            Object.ToolTipText     =   "Education Details "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "GALLERY"
            Object.ToolTipText     =   "Photos/Activities"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "VISION"
            Object.ToolTipText     =   "Our vision and goal"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ABOUT"
            Object.ToolTipText     =   "About us "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "HELP"
            Object.ToolTipText     =   "Find a help"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "STUDENT SIGN UP"
            Object.ToolTipText     =   "Student Sign Up "
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SENIOR CITIZEN SIGN UP"
            Object.ToolTipText     =   "Senior Citizen Sign Up "
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   14400
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":3B819
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":3C46B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":3D0BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":3DD0F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":3E961
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":3F5B3
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":40205
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":40E57
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":41AA9
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":426FB
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":4334D
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   10140
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2743
            MinWidth        =   2743
            Picture         =   "MDIForm1.frx":43F9F
            TextSave        =   "09:47 AM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Picture         =   "MDIForm1.frx":44BF1
            TextSave        =   "20-09-2018"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Picture         =   "MDIForm1.frx":45843
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnumenu 
      Caption         =   "&MENU"
      Begin VB.Menu mnuhostel 
         Caption         =   "&HOSTEL"
         Begin VB.Menu mnureghostel 
            Caption         =   "REGESTRATION"
            Shortcut        =   ^{F4}
         End
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnumjd 
         Caption         =   "&MJD"
         Begin VB.Menu mnuregmjd 
            Caption         =   "REGESTRATION"
         End
      End
      Begin VB.Menu mnusepr 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnueducation 
      Caption         =   "&EDUCATION"
      Begin VB.Menu mnudetails 
         Caption         =   "DETAILS"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnusearch 
      Caption         =   "&SEARCH"
      Begin VB.Menu mnustusearch 
         Caption         =   "&STUDENT SEARCH"
      End
      Begin VB.Menu mnusepre 
         Caption         =   "-"
      End
      Begin VB.Menu MNUSENCITSEARCH 
         Caption         =   "&SENIOR CITIZEN SEARCH"
      End
   End
   Begin VB.Menu mnugallary 
      Caption         =   "&GALLERY"
   End
   Begin VB.Menu mnuvision 
      Caption         =   "&VISION"
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&ABOUT US"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&HELP"
   End
   Begin VB.Menu mnusignup 
      Caption         =   "&SIGN UP"
      Begin VB.Menu mnuhossignup 
         Caption         =   "&HOSTEL SIGN UP"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnumjdsignup 
         Caption         =   "&MJDL SIGN UP"
         Shortcut        =   ^{F9}
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "&EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuabout_Click()
Unload adminlogin
Unload details
Unload gallary
Unload help
Unload hosregisteration
Unload mjdregistration
Unload regadmin
Unload reghostel
Unload regmjd
Unload SEARCH
Unload vision
about.Show
End Sub

Private Sub mnudetails_Click()
'Unload adminlogin
'Unload gallary
'Unload help
'Unload hosregisteration
'Unload mjdregistration
'Unload regadmin
'Unload reghostel
'Unload regmjd
'Unload search
'Unload vision
'Unload about
details.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnugallary_Click()
Unload adminlogin
Unload details
Unload hosregisteration
Unload mjdregistration
Unload regadmin
Unload reghostel
Unload help
Unload regmjd
Unload SEARCH
Unload vision
Unload about
gallary.Show
End Sub

Private Sub mnuhelp_Click()
Unload adminlogin
Unload details
Unload gallary
Unload hosregisteration
Unload mjdregistration
Unload regadmin
Unload reghostel
Unload regmjd
Unload SEARCH
Unload vision
Unload about
help.Show
End Sub

Private Sub mnuhossignup_Click()
Unload adminlogin
Unload details
Unload gallary
Unload mjdregistration
Unload regadmin
Unload reghostel
Unload regmjd
Unload SEARCH
Unload vision
Unload about
Unload help
hosregisteration.Show
End Sub

Private Sub mnumjdsignup_Click()
Unload adminlogin
Unload details
Unload gallary
Unload hosregisteration
Unload regadmin
Unload reghostel
Unload regmjd
Unload SEARCH
Unload vision
Unload about
Unload help
mjdregistration.Show
End Sub

Private Sub mnureghostel_Click()
'Unload adminlogin
'Unload details
'Unload gallary
'Unload hosregisteration
'Unload mjdregistration
'Unload regadmin
'Unload regmjd
'Unload search
'Unload vision
'Unload about
'Unload help
reghostel.Show
End Sub

Private Sub mnuregmjd_Click()
'Unload adminlogin
'Unload details
'Unload gallary
'Unload hosregisteration
'Unload mjdregistration
'Unload regadmin
'Unload reghostel
'Unload search
'Unload vision
'Unload about
'Unload help
regmjd.Show
End Sub

Private Sub MNUSENCITSEARCH_Click()
Unload adminlogin
Unload details
Unload gallary
Unload hosregisteration
Unload mjdregistration
Unload regadmin
Unload reghostel
Unload regmjd
Unload vision
Unload about
Unload help
Unload SEARCH
mjdsearch.Show
End Sub

Private Sub mnustusearch_Click()
Unload adminlogin
Unload details
Unload gallary
Unload hosregisteration
Unload mjdregistration
Unload regadmin
Unload reghostel
Unload regmjd
Unload vision
Unload about
Unload help
SEARCH.Show
End Sub

Private Sub mnuvision_Click()
Unload adminlogin
Unload details
Unload gallary
Unload hosregisteration
Unload mjdregistration
Unload regadmin
Unload reghostel
Unload regmjd
Unload SEARCH
Unload about
Unload help
vision.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
reghostel.Show
End If
If Button.Index = 2 Then
regmjd.Show
End If
If Button.Index = 3 Then
details.Show
End If
If Button.Index = 4 Then
gallary.Show
End If
If Button.Index = 5 Then
vision.Show
End If
If Button.Index = 6 Then
about.Show
End If
If Button.Index = 7 Then
help.Show
End If
If Button.Index = 8 Then
hosregisteration.Show
End If
If Button.Index = 9 Then
mjdregistration.Show
End If
End Sub














