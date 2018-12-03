VERSION 5.00
Begin VB.UserControl mm_circle_progressbar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   MaskColor       =   &H0000FF00&
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   ToolboxBitmap   =   "mon_circle_progressbar.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   270
   End
   Begin VB.Label lbl_progress 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   225
      TabIndex        =   0
      Top             =   0
      Width           =   510
   End
End
Attribute VB_Name = "mm_circle_progressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Const ALTERNATE As Long = 1
Private Const pi As Double = 3.14159265358979
Private Const WINDING As Long = 2

Private Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


Private UserControl_Parent As Control   'detect parent of usercontrol
Private ProgressValue As Single         'value of progressbar

'Declaration for GDI technique -----------------------------------------------
Private Const GdiplusVersion As Long = 1&
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GdiplusStartupOutput
    NotificationHook As Long
    NotificationUnhook As Long
End Type

' The SmoothingMode enumeration specifies the type of
' smoothing (antialiasing) that is applied to lines and curves.
Private Enum SmoothingMode
   SmoothingModeInvalid
   ' Reserved.
   SmoothingModeDefault = 0
   ' Specifies that smoothing is not applied.
   SmoothingModeHighSpeed = 1
   ' Specifies that smoothing is not applied.
   SmoothingModeHighQuality = 2
   ' Specifies that smoothing is applied using an 8 X 4 box filter.
   SmoothingModeNone = 3
   ' Specifies that smoothing is not applied.
   SmoothingModeAntiAlias8x4 = 4
   ' Specifies that smoothing is applied using an 8 X 4 box filter.
   SmoothingModeAntiAlias
   ' Specifies that smoothing is applied using an 8 X 4 box filter.
   SmoothingModeAntiAlias8x8
   ' Specifies that smoothing is applied using an 8 X 8 box filter.
End Enum

Private Enum GpUnit
   UnitWorld = 0
   ' World coordinate (non-physical unit)
   UnitDisplay = 1
   ' Variable -- for PageTransform only
   UnitPixel = 2
   ' Each unit is one device pixel.
   UnitPoint = 3
   ' Each unit is a printer's point, or 1/72 inch.
   UnitInch = 4
   ' Each unit is 1 inch.
   UnitDocument = 5
   ' Each unit is 1/300 inch.
   UnitMillimeter = 6
   ' Each unit is 1 millimeter.
End Enum

Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMode As SmoothingMode) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GdiplusStartupInput, ByRef lpOutput As GdiplusStartupOutput) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef graphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As OLE_COLOR, ByVal Width As Single, ByVal unit As GpUnit, ByRef pen As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal Color As OLE_COLOR, ByRef brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As Long
Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal path As Long) As Long
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDrawPie Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillPie Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long


'""""""""""""" new -------------
'for GDI Gradient Brush
''Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
''Private Declare Function GdipCreateLineBrushI Lib "gdiplus" (point1 As POINTAPI, point2 As POINTAPI, ByVal color1 As Colors, ByVal color2 As Colors, ByVal WrapMd As WrapMode, lineGradient As Long) As GpStatus
'used in blend color function
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'""""""""""""""""""""""""""""""----


Private m_Token As Long
' ---------------------------------------------------------------------
' translate color RGB (OLE_color) to GDI+ (includ Alpha) ---------------
Private Type COLORBYTES
   BlueByte As Byte
   GreenByte As Byte
   RedByte As Byte
   AlphaByte As Byte
End Type

Private Type COLORLONG
   longval As Long
End Type

Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
' -----------------------------------------------------------------------


'Private i As Integer
Private mbackcolor As OLE_COLOR
Private mbackcolor_border As OLE_COLOR
Private mprogressbar As OLE_COLOR
Private mprogressbar_border As OLE_COLOR
Private mprogressbar_back As OLE_COLOR
Private mprogressbar_back_border As OLE_COLOR
Private mfrontcircle As OLE_COLOR
Private mfrontcircle_border As OLE_COLOR

Const m_def_Max = 100
Const m_def_BorderColor = 5329233
Const m_def_Border = 0
Const m_def_Epaisseur = 10
Const m_def_Transparent = 0
Const m_def_CaptionVisible = 1
Const m_def_prograssInactiveZone = 0
Const m_def_ProgressBackcolor = 0
Const m_def_BackColor = 0
Const m_def_Progress_backcolor = 0
Const m_def_progressbar_color = 0
Const m_def_progressbar_frontof_color = 0
Const m_def_value = 0
Const m_def_CaptionSymbol = "%"
Const m_def_Use_Gradient = 0
Const m_def_ValueVisibility = 1
Private m_CaptionSymbol As String
Private m_Max As Integer
Private m_BorderColor As OLE_COLOR
Private m_Border As Boolean
Private m_Epaisseur As Integer
Private m_Transparent As Boolean
Private m_CaptionVisible As Boolean
Private m_prograssInactiveZone As OLE_COLOR
Private m_ProgressBackcolor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_progressbarcolor As OLE_COLOR
Private m_progressbar_frontof_color As OLE_COLOR
Private m_value As Long
Private m_Use_Gradient As Boolean
Private m_ValueVisibility As Boolean
Const m_def_Advanced_Style_Animation = 0
Const m_def_Advanced_Style = 1
Const m_def_Advanced_Style_BackGroundCircle1_Start = 150 '180
Const m_def_Advanced_Style_BackGroundCircle1_Value = 100 '90
Const m_def_Advanced_Style_BackGroundCircle2_Start = 0
Const m_def_Advanced_Style_BackGroundCircle2_Value = 45
Dim m_Advanced_Style_Animation As Boolean
Dim m_Advanced_Style As Boolean
Dim m_Advanced_Style_BackGroundCircle1_Start As Integer
Dim m_Advanced_Style_BackGroundCircle1_Value As Integer
Dim m_Advanced_Style_BackGroundCircle2_Start As Integer
Dim m_Advanced_Style_BackGroundCircle2_Value As Integer

Private Function blendcolor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
'for my gradient
    Dim lCFrom As Long
    Dim lCTo As Long
    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
   
    lCFrom = GetLngColor(oColorFrom)
    lCTo = GetLngColor(oColorTo)
    
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
   
    blendcolor = RGB( _
      ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
      ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
      ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))
      
End Function


Sub DrawGradientInPie(X As Integer, Y As Integer, w As Integer, h As Integer, color1 As OLE_COLOR, m_Graphics As Long, Optional invert As Boolean = False, Optional transparent1 As Boolean = False)

    Dim pt1 As POINTAPI, pt2 As POINTAPI
    Dim brush As Long ', pen As Long
    Dim blendedcolor As OLE_COLOR
    
    pt1.X = X
    pt1.Y = Y
    pt2.X = w
    pt2.Y = h
    
    blendedcolor = blendcolor(color1, vbBlack, 140)
    
    If transparent1 Then
        Call GdipCreateLineBrushI(pt1, pt2, color1, Transparent, WrapModeTileFlipXY, brush)
    Else
        If invert Then
            Call GdipCreateLineBrushI(pt1, pt2, GetRGB_VB2GDIP(blendedcolor), GetRGB_VB2GDIP(color1), WrapModeTileFlipXY, brush)
        Else
            Call GdipCreateLineBrushI(pt1, pt2, GetRGB_VB2GDIP(color1), GetRGB_VB2GDIP(blendedcolor), WrapModeTileFlipXY, brush)
        End If
    End If
    
    ' Fill Ellipse with gradient brush
    Call GdipFillEllipseI(m_Graphics, brush, pt1.X, pt1.Y, pt2.X, pt2.Y)
    
    Call GdipDeleteBrush(brush)


End Sub

Private Function GetLngColor(Color As Long) As Long
    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else
        GetLngColor = Color
    End If
End Function



Private Sub LngToRGB(LCul As Long, r As Byte, g As Byte, B As Byte)
   r = LCul And &HFF&
   g = (LCul And &HFF00&) \ &H100&
   B = (LCul And &HFF0000) \ &H10000
End Sub

Private Sub Detect_Parent_Of_Usercontrol()
    
On Error Resume Next
    
If UserControl.Parent.hWnd = UserControl.ContainerHwnd Then
    Set UserControl_Parent = Nothing
Else
    Set UserControl_Parent = UserControl.Extender.Container
End If


End Sub
Sub Update_Progressbar()

    Dim m_Graphics As Long
    Dim m_Brush_01 As Long
    Dim m_Brush_02 As Long
    Dim m_Brush_03 As Long
    Dim m_Brush_04 As Long
    Dim m_Brush_05 As Long
    Dim m_Pen_01 As Long
    Dim m_Width As Single
    Dim m_Height As Single
    Dim val As Long
    
    Dim radian As Long
    
    If m_value < 0 Then m_value = 0 'just in case
    
    lbl_progress.Visible = m_CaptionVisible
    
    ''UserControl.Cls
    UserControl.BackColor = m_BackColor
    
    lbl_progress.ForeColor = UserControl.ForeColor
    Set lbl_progress.Font = UserControl.Font
    
    m_Width = UserControl.Width / Screen.TwipsPerPixelX
    m_Height = UserControl.Height / Screen.TwipsPerPixelY 'm_Width
    
    Call GdipCreateFromHDC(UserControl.hDC, m_Graphics)
    Call GdipSetSmoothingMode(m_Graphics, SmoothingModeAntiAlias8x4)
    
    If Not m_Advanced_Style Then
        'First Circle : Background *************************************************
        ProgressValue = 360
        'Border
        If Border = True Then
            Call GdipCreatePen1(GetRGB_VB2GDIP(m_BorderColor), 1, UnitPixel, m_Pen_01)
            Call GdipDrawPie(m_Graphics, m_Pen_01, 5, 5, m_Width - 10, m_Height - 10, -90, ProgressValue)
        End If
        'background
        If Not Use_Gradient Then
            Call GdipCreateSolidFill(GetRGB_VB2GDIP(m_ProgressBackcolor), m_Brush_01)
            Call GdipFillPie(m_Graphics, m_Brush_01, 5, 5, m_Width - 10, m_Height - 10, -90, ProgressValue)
        Else
            Call DrawGradientInPie(5, 5, m_Width - 10, m_Height - 10, m_ProgressBackcolor, m_Graphics)
        End If
    Else
        'First Circle : Background *************************************************
        ProgressValue = Advanced_Style_BackGroundCircle1_Value
        'Border
        If Border = True Then
            Call GdipCreatePen1(GetRGB_VB2GDIP(m_BorderColor), 1, UnitPixel, m_Pen_01)
            Call GdipDrawPie(m_Graphics, m_Pen_01, 5, 5, m_Width - 10, m_Height - 10, Advanced_Style_BackGroundCircle1_Start, ProgressValue)
        End If
        'background
        'If Not Use_Gradient Then
            Call GdipCreateSolidFill(GetRGB_VB2GDIP(m_ProgressBackcolor), m_Brush_01)
            Call GdipFillPie(m_Graphics, m_Brush_01, 5, 5, m_Width - 10, m_Height - 10, Advanced_Style_BackGroundCircle1_Start, ProgressValue)
        'Else
        '    Call DrawGradientInPie(5, 5, m_Width - 10, m_Height - 10, m_ProgressBackcolor, m_Graphics)
        'End If
        
        'Second Circle : Background *************************************************
        ProgressValue = Advanced_Style_BackGroundCircle2_Value
        'Border
        If Border = True Then
            Call GdipCreatePen1(GetRGB_VB2GDIP(m_BorderColor), 1, UnitPixel, m_Pen_01)
            Call GdipDrawPie(m_Graphics, m_Pen_01, 5, 5, m_Width - 10, m_Height - 10, Advanced_Style_BackGroundCircle2_Start, ProgressValue)
        End If
        'background
        'If Not Use_Gradient Then
            Call GdipCreateSolidFill(GetRGB_VB2GDIP(m_ProgressBackcolor), m_Brush_01)
            Call GdipFillPie(m_Graphics, m_Brush_01, 5, 5, m_Width - 10, m_Height - 10, Advanced_Style_BackGroundCircle2_Start, ProgressValue)
        'Else
        '    Call DrawGradientInPie(5, 5, m_Width - 10, m_Height - 10, m_ProgressBackcolor, m_Graphics)
        'End If
        
    End If
        
    
    'Second Circle : inactive zone Progressbar *********************************
    ProgressValue = 360 '---90
    'background
    If Not Use_Gradient Then
        Call GdipCreateSolidFill(GetRGB_VB2GDIP(m_prograssInactiveZone), m_Brush_02)
        Call GdipFillPie(m_Graphics, m_Brush_02, 5 + (m_Epaisseur), 5 + (m_Epaisseur), m_Width - 10 - (m_Epaisseur * 2), m_Height - 10 - (m_Epaisseur * 2), 0, ProgressValue)
    Else
        Call DrawGradientInPie(5 + (m_Epaisseur), 5 + (m_Epaisseur), m_Width - 10 - (m_Epaisseur * 2), m_Height - 10 - (m_Epaisseur * 2), m_prograssInactiveZone, m_Graphics, True)
    End If
    
    
    
    'Third Circle : Progressbar Value ******************************************
'    ProgressValue = m_value * 360 / m_Max
    Call GdipCreateSolidFill(GetRGB_VB2GDIP(m_progressbarcolor), m_Brush_03)
    Call GdipFillPie(m_Graphics, m_Brush_03, 5 + (m_Epaisseur), 5 + (m_Epaisseur), m_Width - 10 - (m_Epaisseur * 2), m_Height - 10 - (m_Epaisseur * 2), -90, ProgressValue)


    'Fourth Circle : in front of progress **************************************
    ProgressValue = 360
    If Not Use_Gradient Then
        Call GdipCreateSolidFill(GetRGB_VB2GDIP(m_progressbar_frontof_color), m_Brush_04)
        Call GdipFillPie(m_Graphics, m_Brush_04, 5 + (m_Epaisseur * 2), 5 + (m_Epaisseur * 2), m_Width - 10 - (m_Epaisseur * 4), m_Height - 10 - (m_Epaisseur * 4), -90, ProgressValue)
    Else
        Call DrawGradientInPie(5 + (m_Epaisseur * 2), 5 + (m_Epaisseur * 2), m_Width - 10 - (m_Epaisseur * 4), m_Height - 10 - (m_Epaisseur * 4), m_progressbar_frontof_color, m_Graphics)

    End If
    
    'Show or not the Caption/Value
    
    If CaptionVisible Then 'if caption is visible show value (if visibility true) and/or caption symbol (if exist)
        If ValueVisibility Then 'show the value
            lbl_progress.Caption = Int(m_value)
        Else
            lbl_progress.Caption = ""
        End If
        
        If m_CaptionSymbol <> "" Then 'show symbol
            lbl_progress.Caption = lbl_progress.Caption & m_CaptionSymbol '"%"
        End If
    End If
    
   
    UserControl.Refresh
    
    'Free Mem
    Call GdipDeletePen(m_Pen_01)
    Call GdipDeleteBrush(m_Brush_01)
    Call GdipDeleteBrush(m_Brush_02)
    Call GdipDeleteBrush(m_Brush_03)
    Call GdipDeleteBrush(m_Brush_04)
    Call GdipDeleteBrush(m_Brush_05)
    Call GdipDeleteGraphics(m_Graphics)

Error:

End Sub

Private Function ColorARGB(ByVal Alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Long
   Dim bytestruct As COLORBYTES
   Dim result As COLORLONG
   With bytestruct
      .AlphaByte = Alpha
      .RedByte = Red
      .GreenByte = Green
      .BlueByte = Blue
   End With
   LSet result = bytestruct
   ColorARGB = result.longval
End Function
Private Function GetRGB_VB2GDIP(ByVal lColor As Long, Optional ByVal Alpha As Byte = 255) As Long
   'Convert RGB to GDI
   Dim rgbq As RGBQUAD
   CopyMemory rgbq, lColor, 4
   GetRGB_VB2GDIP = ColorARGB(Alpha, rgbq.rgbBlue, rgbq.rgbGreen, rgbq.rgbRed)
End Function

Private Function ShutdownGDIPlus() As Long
    ShutdownGDIPlus = GdiplusShutdown(m_Token)
End Function

Private Function StartUpGDIPlus(ByVal GdipVersion As Long) As Long
    Dim GdipStartupInput As GdiplusStartupInput
    Dim GdipStartupOutput As GdiplusStartupOutput
    GdipStartupInput.GdiplusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(m_Token, GdipStartupInput, GdipStartupOutput)
End Function


Private Sub PaintControl()

    Set lbl_progress.Font = UserControl.Font
    'center the lbl_progress label
    lbl_progress.Move (UserControl.ScaleWidth - lbl_progress.Width) / 2, (UserControl.ScaleHeight - lbl_progress.Height) / 2

    UserControl.Cls

    If m_Transparent Then 'if transparency activated
        Detect_Parent_Of_Usercontrol
        If UserControl_Parent Is Nothing Then 'the parent is directly the Form
            If UserControl.Parent.Picture <> 0 Then UserControl.PaintPicture UserControl.Parent.Picture, 0, 0, UserControl.Width, UserControl.Height, UserControl.Extender.Left / Screen.TwipsPerPixelX, UserControl.Extender.Top / Screen.TwipsPerPixelY, UserControl.Width, UserControl.Height
        Else 'the parent is Picturebox or other object having a picture
            If UserControl_Parent.Picture <> 0 Then UserControl.PaintPicture UserControl_Parent.Picture, 0, 0, UserControl.Width, UserControl.Height, UserControl.Extender.Left / Screen.TwipsPerPixelX, UserControl.Extender.Top / Screen.TwipsPerPixelY, UserControl.Width, UserControl.Height
        End If
    End If

    If Advanced_Style Then
        Timer1.Enabled = Advanced_Style_Animation
    Else
        Timer1.Enabled = False
    End If

    Update_Progressbar
    
    
    'where is usercontrol.left and usercontrol.top ?
    'you can acces them using usercontrol.extender.left and usercontrol.extender.top
    'why?
    'there is the answer (from internet) :
    ' -------------------------------------------------------------------------------------------------------------
    ' | Certain standard properties such as top and left are provided to a control by its container.
    ' | These properties are exposed to the UserControl developer through the Extender object:
    ' | UserControl.Extender.Top
    ' | Note: These properties are intended for the client's use. You should only read them, never change them.
    ' -------------------------------------------------------------------------------------------------------------
    
End Sub


Private Sub Timer1_Timer()

If Advanced_Style Then

    Advanced_Style_BackGroundCircle1_Start = Advanced_Style_BackGroundCircle1_Start - 1
    Advanced_Style_BackGroundCircle2_Start = Advanced_Style_BackGroundCircle2_Start - 1
    
    If Advanced_Style_BackGroundCircle1_Start <= 0 Then Advanced_Style_BackGroundCircle1_Start = 360
    If Advanced_Style_BackGroundCircle2_Start <= 0 Then Advanced_Style_BackGroundCircle2_Start = 360
    
    PaintControl
Else 'just in case
    Timer1.Enabled = False
End If
End Sub

Private Sub UserControl_Initialize()
    'init GDI+
    Call StartUpGDIPlus(GdiplusVersion)
End Sub

Private Sub UserControl_Paint()
PaintControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_value = PropBag.ReadProperty("value", m_def_value)
    m_progressbarcolor = PropBag.ReadProperty("ProgressbarColor", m_def_progressbar_color)
    m_progressbar_frontof_color = PropBag.ReadProperty("progressbar_frontof_color", m_def_progressbar_frontof_color)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ProgressBackcolor = PropBag.ReadProperty("ProgressBackcolor", m_def_ProgressBackcolor)
    m_prograssInactiveZone = PropBag.ReadProperty("prograssInactiveZone", m_def_prograssInactiveZone)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_CaptionVisible = PropBag.ReadProperty("CaptionVisible", m_def_CaptionVisible)
    m_Transparent = PropBag.ReadProperty("Transparent", m_def_Transparent)
    m_Epaisseur = PropBag.ReadProperty("Epaisseur", m_def_Epaisseur)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_Border = PropBag.ReadProperty("Border", m_def_Border)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    
    m_Use_Gradient = PropBag.ReadProperty("Use_Gradient", m_def_Use_Gradient)
    m_CaptionSymbol = PropBag.ReadProperty("CaptionSymbol", m_def_CaptionSymbol)
    m_ValueVisibility = PropBag.ReadProperty("ValueVisibility", m_def_ValueVisibility)
    m_Advanced_Style_BackGroundCircle1_Start = PropBag.ReadProperty("Advanced_Style_BackGroundCircle1_Start", m_def_Advanced_Style_BackGroundCircle1_Start)
    m_Advanced_Style_BackGroundCircle1_Value = PropBag.ReadProperty("Advanced_Style_BackGroundCircle1_Value", m_def_Advanced_Style_BackGroundCircle1_Value)
    m_Advanced_Style_BackGroundCircle2_Start = PropBag.ReadProperty("Advanced_Style_BackGroundCircle2_Start", m_def_Advanced_Style_BackGroundCircle2_Start)
    m_Advanced_Style_BackGroundCircle2_Value = PropBag.ReadProperty("Advanced_Style_BackGroundCircle2_Value", m_def_Advanced_Style_BackGroundCircle2_Value)
    m_Advanced_Style = PropBag.ReadProperty("Advanced_Style", m_def_Advanced_Style)
    m_Advanced_Style_Animation = PropBag.ReadProperty("Advanced_Style_Animation", m_def_Advanced_Style_Animation)
End Sub
Private Sub UserControl_Resize()
    
    UserControl.Height = UserControl.Width
    lbl_progress.Width = UserControl.Width
    
    PaintControl
End Sub

Private Sub UserControl_Show()

    PaintControl

End Sub

Private Sub UserControl_Terminate()
Call ShutdownGDIPlus 'Shutdown GDI
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("value", m_value, m_def_value)
    Call PropBag.WriteProperty("ProgressbarColor", m_progressbarcolor, m_def_progressbar_color)
    Call PropBag.WriteProperty("progressbar_frontof_color", m_progressbar_frontof_color, m_def_progressbar_frontof_color)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ProgressBackcolor", m_ProgressBackcolor, m_def_ProgressBackcolor)
    Call PropBag.WriteProperty("prograssInactiveZone", m_prograssInactiveZone, m_def_prograssInactiveZone)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("CaptionVisible", m_CaptionVisible, m_def_CaptionVisible)
    Call PropBag.WriteProperty("Transparent", m_Transparent, m_def_Transparent)
    Call PropBag.WriteProperty("Epaisseur", m_Epaisseur, m_def_Epaisseur)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("Border", m_Border, m_def_Border)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    
    Call PropBag.WriteProperty("Use_Gradient", m_Use_Gradient, m_def_Use_Gradient)
    Call PropBag.WriteProperty("CaptionSymbol", m_CaptionSymbol, m_def_CaptionSymbol)
    Call PropBag.WriteProperty("ValueVisibility", m_ValueVisibility, m_def_ValueVisibility)
    Call PropBag.WriteProperty("Advanced_Style_BackGroundCircle1_Start", m_Advanced_Style_BackGroundCircle1_Start, m_def_Advanced_Style_BackGroundCircle1_Start)
    Call PropBag.WriteProperty("Advanced_Style_BackGroundCircle1_Value", m_Advanced_Style_BackGroundCircle1_Value, m_def_Advanced_Style_BackGroundCircle1_Value)
    Call PropBag.WriteProperty("Advanced_Style_BackGroundCircle2_Start", m_Advanced_Style_BackGroundCircle2_Start, m_def_Advanced_Style_BackGroundCircle2_Start)
    Call PropBag.WriteProperty("Advanced_Style_BackGroundCircle2_Value", m_Advanced_Style_BackGroundCircle2_Value, m_def_Advanced_Style_BackGroundCircle2_Value)
    Call PropBag.WriteProperty("Advanced_Style", m_Advanced_Style, m_def_Advanced_Style)
    Call PropBag.WriteProperty("Advanced_Style_Animation", m_Advanced_Style_Animation, m_def_Advanced_Style_Animation)
End Sub
Private Sub UserControl_InitProperties()
    
    m_value = 45 'm_def_value
    m_ProgressBackcolor = 15395562 'm_def_Progress_backcolor
    m_progressbarcolor = 3320424  '&HE5E5E5 'm_def_progressbar_color
    m_progressbar_frontof_color = 15263976 ' m_def_progressbar_frontof_color
    m_CaptionVisible = True: lbl_progress.Visible = True
    m_BackColor = &HE5E5E5 'm_def_BackColor
    m_prograssInactiveZone = 13487565
    
    Set UserControl.Font = Ambient.Font
    Set lbl_progress.Font = UserControl.Font
    
    m_CaptionVisible = m_def_CaptionVisible
    m_Transparent = False 'm_def_Transparent
    m_Epaisseur = m_def_Epaisseur
    m_BorderColor = m_def_BorderColor
    m_Border = m_def_Border
    m_Max = m_def_Max
    m_Use_Gradient = m_def_Use_Gradient
    m_CaptionSymbol = m_def_CaptionSymbol
    m_ValueVisibility = m_def_ValueVisibility
    m_Advanced_Style_BackGroundCircle1_Start = m_def_Advanced_Style_BackGroundCircle1_Start
    m_Advanced_Style_BackGroundCircle1_Value = m_def_Advanced_Style_BackGroundCircle1_Value
    m_Advanced_Style_BackGroundCircle2_Start = m_def_Advanced_Style_BackGroundCircle2_Start
    m_Advanced_Style_BackGroundCircle2_Value = m_def_Advanced_Style_BackGroundCircle2_Value
    m_Advanced_Style = m_def_Advanced_Style
    m_Advanced_Style_Animation = m_def_Advanced_Style_Animation
    
    '
    Timer1.Enabled = m_Advanced_Style_Animation
    
End Sub

Public Property Get value() As Integer
    value = m_value
End Property

Public Property Let value(ByVal New_value As Integer)
    m_value = New_value
    PropertyChanged "value"
    
    If m_value < 0 Or m_value > 100 Then m_value = 0
    
    If m_value > m_Max Then m_value = m_Max
    
    'update
    PaintControl
    
End Property

Public Property Get ProgressbarColor() As OLE_COLOR
    ProgressbarColor = m_progressbarcolor
End Property

Public Property Let ProgressbarColor(ByVal New_progressbarcolor As OLE_COLOR)
    m_progressbarcolor = New_progressbarcolor
    PropertyChanged "ProgressbarColor"
    PaintControl
End Property

Public Property Get progressbar_frontof_color() As OLE_COLOR
    progressbar_frontof_color = m_progressbar_frontof_color
End Property

Public Property Let progressbar_frontof_color(ByVal New_progressbar_frontof_color As OLE_COLOR)
    m_progressbar_frontof_color = New_progressbar_frontof_color
    PropertyChanged "progressbar_frontof_color"
    PaintControl
End Property


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Renvoie ou définit la couleur d'arrière-plan utilisée pour afficher le texte et les graphiques d'un objet."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    PaintControl
    
End Property

Public Property Get ProgressBackcolor() As OLE_COLOR
    ProgressBackcolor = m_ProgressBackcolor
End Property

Public Property Let ProgressBackcolor(ByVal New_ProgressBackcolor As OLE_COLOR)
    m_ProgressBackcolor = New_ProgressBackcolor
    PropertyChanged "ProgressBackcolor"
    PaintControl
End Property

Public Property Get prograssInactiveZone() As OLE_COLOR
    prograssInactiveZone = m_prograssInactiveZone
End Property

Public Property Let prograssInactiveZone(ByVal New_prograssInactiveZone As OLE_COLOR)
    m_prograssInactiveZone = New_prograssInactiveZone
    PropertyChanged "prograssInactiveZone"
    PaintControl
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Renvoie un objet Font."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    
    Set lbl_progress.Font = New_Font
    lbl_progress.Move (UserControl.ScaleWidth - lbl_progress.Width) / 2, (UserControl.ScaleHeight - lbl_progress.Height) / 2

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Renvoie ou définit la couleur de premier plan utilisée pour afficher le texte et les graphiques d'un objet."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    lbl_progress.ForeColor() = New_ForeColor
End Property

Public Property Get CaptionVisible() As Boolean
    CaptionVisible = m_CaptionVisible
End Property

Public Property Let CaptionVisible(ByVal New_CaptionVisible As Boolean)
    m_CaptionVisible = New_CaptionVisible
    PropertyChanged "CaptionVisible"
    
    lbl_progress.Visible = m_CaptionVisible
    
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal New_Transparent As Boolean)
    m_Transparent = New_Transparent
    PropertyChanged "Transparent"
    
    PaintControl
    
End Property

Public Property Get Epaisseur() As Integer
    Epaisseur = m_Epaisseur
End Property

Public Property Let Epaisseur(ByVal New_Epaisseur As Integer)
    m_Epaisseur = New_Epaisseur
    PropertyChanged "Epaisseur"
    
    If m_Epaisseur < 1 Or m_Epaisseur > 50 Then m_Epaisseur = 10
    
    PaintControl
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    
    If m_Border = True Then PaintControl
End Property

Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    PropertyChanged "Border"
    
    PaintControl
    
End Property

Public Property Get Max() As Integer
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Integer)
    m_Max = New_Max
    PropertyChanged "Max"
    
    If m_Max < 10 Then m_Max = 10
    
    If m_value > m_Max Then m_value = m_Max

    PaintControl

    
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=0,0,0,0
Public Property Get Use_Gradient() As Boolean
    Use_Gradient = m_Use_Gradient
End Property

Public Property Let Use_Gradient(ByVal New_Use_Gradient As Boolean)
    m_Use_Gradient = New_Use_Gradient
    PropertyChanged "Use_Gradient"
    
    PaintControl

End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=13,0,0,%
Public Property Get CaptionSymbol() As String
    CaptionSymbol = m_CaptionSymbol
End Property

Public Property Let CaptionSymbol(ByVal New_CaptionSymbol As String)
    m_CaptionSymbol = New_CaptionSymbol
    PropertyChanged "CaptionSymbol"
    
    PaintControl
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=0,0,0,1
Public Property Get ValueVisibility() As Boolean
    ValueVisibility = m_ValueVisibility
End Property

Public Property Let ValueVisibility(ByVal New_ValueVisibility As Boolean)
    m_ValueVisibility = New_ValueVisibility
    PropertyChanged "ValueVisibility"
    
    PaintControl
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=7,0,0,0
Public Property Get Advanced_Style_BackGroundCircle1_Start() As Integer
    Advanced_Style_BackGroundCircle1_Start = m_Advanced_Style_BackGroundCircle1_Start
End Property

Public Property Let Advanced_Style_BackGroundCircle1_Start(ByVal New_Advanced_Style_BackGroundCircle1_Start As Integer)
    m_Advanced_Style_BackGroundCircle1_Start = New_Advanced_Style_BackGroundCircle1_Start
    PropertyChanged "Advanced_Style_BackGroundCircle1_Start"
    
    If m_Advanced_Style_BackGroundCircle1_Start < 0 Or m_Advanced_Style_BackGroundCircle1_Start > 360 Then m_Advanced_Style_BackGroundCircle1_Start = 0
    
    PaintControl
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=7,0,0,0
Public Property Get Advanced_Style_BackGroundCircle1_Value() As Integer
    Advanced_Style_BackGroundCircle1_Value = m_Advanced_Style_BackGroundCircle1_Value
End Property

Public Property Let Advanced_Style_BackGroundCircle1_Value(ByVal New_Advanced_Style_BackGroundCircle1_Value As Integer)
    m_Advanced_Style_BackGroundCircle1_Value = New_Advanced_Style_BackGroundCircle1_Value
    PropertyChanged "Advanced_Style_BackGroundCircle1_Value"
    
    If m_Advanced_Style_BackGroundCircle1_Value < 0 Or m_Advanced_Style_BackGroundCircle1_Value > 360 Then m_Advanced_Style_BackGroundCircle1_Value = 0
    
    PaintControl
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=7,0,0,0
Public Property Get Advanced_Style_BackGroundCircle2_Start() As Integer
    Advanced_Style_BackGroundCircle2_Start = m_Advanced_Style_BackGroundCircle2_Start
End Property

Public Property Let Advanced_Style_BackGroundCircle2_Start(ByVal New_Advanced_Style_BackGroundCircle2_Start As Integer)
    m_Advanced_Style_BackGroundCircle2_Start = New_Advanced_Style_BackGroundCircle2_Start
    PropertyChanged "Advanced_Style_BackGroundCircle2_Start"
    
    If m_Advanced_Style_BackGroundCircle2_Start < 0 Or m_Advanced_Style_BackGroundCircle2_Start > 360 Then m_Advanced_Style_BackGroundCircle2_Start = 0
    PaintControl
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=7,0,0,0
Public Property Get Advanced_Style_BackGroundCircle2_Value() As Integer
    Advanced_Style_BackGroundCircle2_Value = m_Advanced_Style_BackGroundCircle2_Value
End Property

Public Property Let Advanced_Style_BackGroundCircle2_Value(ByVal New_Advanced_Style_BackGroundCircle2_Value As Integer)
    m_Advanced_Style_BackGroundCircle2_Value = New_Advanced_Style_BackGroundCircle2_Value
    PropertyChanged "Advanced_Style_BackGroundCircle2_Value"
    
    If m_Advanced_Style_BackGroundCircle2_Value < 0 Or m_Advanced_Style_BackGroundCircle2_Value > 360 Then m_Advanced_Style_BackGroundCircle2_Value = 0
    PaintControl
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=0,0,0,0
Public Property Get Advanced_Style() As Boolean
    Advanced_Style = m_Advanced_Style
End Property

Public Property Let Advanced_Style(ByVal New_Advanced_Style As Boolean)
    m_Advanced_Style = New_Advanced_Style
    PropertyChanged "Advanced_Style"
    
    PaintControl
    
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=0,0,0,0
Public Property Get Advanced_Style_Animation() As Boolean
    Advanced_Style_Animation = m_Advanced_Style_Animation
End Property

Public Property Let Advanced_Style_Animation(ByVal New_Advanced_Style_Animation As Boolean)
    m_Advanced_Style_Animation = New_Advanced_Style_Animation
    PropertyChanged "Advanced_Style_Animation"
    
    If Advanced_Style Then
        Timer1.Enabled = Advanced_Style_Animation
    Else
        Timer1.Enabled = False
    End If
End Property

