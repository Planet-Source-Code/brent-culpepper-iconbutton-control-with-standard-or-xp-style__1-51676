VERSION 5.00
Begin VB.UserControl IconButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   300
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   20
   Begin VB.PictureBox picSource 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "IconButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************
'Project:       IconButton Control with standard or xp style
'Programmer:    Brent Culpepper (IDontKnow)
'*************************************************************
'Copyright© Brent Culpepper, All Rights Reserved

'The subclassing and mouse-tracking modules contain
'copyrights by Steve McMahon from vbAccelerator.

' AS ALWAYS:
'You may use this code in your projects. You may
'distribute this code to others if credit is given
'where credit is due. You may NOT sell this, either
'compiled or as source code. You may NOT post this
'code and claim it as your own work.

'This code is provided AS-IS and no warranty or
'guarantee is either expressed or implied. The author
'is not responsible if it starts dating your girlfriend,
'leaving the cap off your toothpaste, or making obscene
'phone calls to your mother-in-law! ;)



Option Explicit

' ### API Constants and Declares: ###
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

' Alternative EDGE styles (Combines the constants above)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

' Constants for the grfFlags parameter in the DrawEdge API
Private Const BF_FLAT = &H4000
Private Const BF_LEFT = &H1
Private Const BF_MONO = &H8000
Private Const BF_MIDDLE = &H800
Private Const BF_RIGHT = &H4
Private Const BF_SOFT = &H1000
Private Const BF_TOP = &H2
Private Const BF_ADJUST = &H2000
Private Const BF_BOTTOM = &H8
Private Const BF_DIAGONAL = &H10
Private Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)

Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4

Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000

Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8

Private Const CLR_INVALID = -1

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Sub DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long)
    
'Event Declarations:
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

' Property Types and Enums:
Public Enum BUTTON_STYLE
    [PopUp Button] = 0
    [XP Button] = 1
End Enum

Public Enum BUTTON_VALUE
    ibUnpressed = 0
    ibPressed = 1
End Enum

Const m_DefaultSize16x16 As Single = 360
Const m_DefaultSize32x32 As Single = 600

'Default Property Values:
Const m_def_Value = 0
Const m_def_Style = 0

Private m_Value As BUTTON_VALUE
Private m_Style As BUTTON_STYLE

'Private members:
Private m_bMouseOver As Boolean
Private m_bPressed As Boolean
Private m_bPrivateResize As Boolean
Private m_Icon As Long
Private m_IconSize As Long


Private WithEvents m_Rodent As cMouseTrack
Attribute m_Rodent.VB_VarHelpID = -1

'### PROPERTIES ###
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawButton
End Property

Public Property Get ButtonStyle() As BUTTON_STYLE
Attribute ButtonStyle.VB_Description = "Returns/sets the button style. Options are a standard popup button or a button drawn in the XP style. Property is Read Only at runtime."
Attribute ButtonStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ButtonStyle = m_Style
End Property

Public Property Let ButtonStyle(ByVal New_Style As BUTTON_STYLE)
    If Ambient.UserMode Then Err.Raise 382
    m_Style = New_Style
    PropertyChanged "ButtonStyle"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    DrawButton
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/sets the picture displayed. The picture must be an icon, either 16x16 or 32x32."
    Set Picture = picSource.Picture
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set picSource.Picture = New_Picture
    ValidatePicture
    PropertyChanged "Picture"
End Property

Public Property Get Value() As BUTTON_VALUE
Attribute Value.VB_Description = "Returns/sets the appearance of the button after it is pressed."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As BUTTON_VALUE)
    m_Value = New_Value
    PropertyChanged "Value"
    DrawButton
End Property

'### MOUSE TRACKING ###
Private Sub m_Rodent_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, X As Single, Y As Single)
    If Not (m_Rodent.Tracking) Then m_Rodent.StartMouseTracking
End Sub

Private Sub m_Rodent_MouseLeave()
    m_bMouseOver = False
    DrawButton
End Sub

'### USERCONTROL EVENTS ###
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bPressed = True
    DrawButton
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not (m_Rodent.Tracking) Then m_Rodent.StartMouseTracking
    If Not m_bMouseOver Then
        m_bMouseOver = True
        DrawButton
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not ((X >= 0 And X <= ScaleWidth) And (Y >= 0 And Y <= ScaleHeight)) Then m_bMouseOver = False
    m_bPressed = False
    DrawButton
    If Button = vbLeftButton And m_bMouseOver Then
        RaiseEvent Click
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    Static bInit As Boolean
    If Not bInit Then
        UserControl.Height = m_DefaultSize16x16
        UserControl.Width = m_DefaultSize16x16
        bInit = True
    End If
    If m_bPrivateResize Then
        If m_IconSize = 16 Then
            UserControl.Height = m_DefaultSize16x16
            UserControl.Width = m_DefaultSize16x16
        Else
            UserControl.Height = m_DefaultSize32x32
            UserControl.Width = m_DefaultSize32x32
        End If
    End If
    DrawButton
End Sub

Private Sub UserControl_Paint()
    DrawButton
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "BackColor" Then BackColor = Ambient.BackColor
End Sub

Private Sub UserControl_Show()
    InitializeTracking
End Sub

Private Sub UserControl_Terminate()
    ' Stop mouse tracking:
    If Not m_Rodent Is Nothing Then
        m_Rodent.DetachMouseTracking
        Set m_Rodent = Nothing
    End If
End Sub

'### PRIVATE ROUTINES ###
Private Sub InitializeTracking()
    If UserControl.Ambient.UserMode Then
        Set m_Rodent = New cMouseTrack
        m_Rodent.AttachMouseTracking UserControl.hwnd
    End If
End Sub

Private Sub DrawButton()
    Dim bar As RECT
    Dim xpos As Long
    Dim ypos As Long
    Dim lR As Long
    Dim brsh As Long
    Dim clr As Long
    Dim Left_x As Long, Top_y As Long
    Dim Right_x As Long, Bottom_y As Long
    Dim intState As Integer
        
    UserControl.Cls
    Left_x = ScaleLeft
    Top_y = ScaleTop
    Right_x = ScaleWidth
    Bottom_y = ScaleHeight
    
    ' Draw the icon in the current state
    If m_Icon <> 0 Then
        xpos = (ScaleWidth / 2) - (m_IconSize / 2)
        ypos = (ScaleHeight / 2) - (m_IconSize / 2)
    End If
    
    If Enabled Then
        If Value = 0 And Not (m_bMouseOver) And Not (m_bPressed) Then intState = 0
        If Value = 0 And m_bMouseOver And Not (m_bPressed) Then intState = 1
        If m_bMouseOver And m_bPressed Then intState = 2
        If Value = 1 And Not (m_bPressed) Then intState = 3
    Else
        intState = 4
    End If
    
    ' Draw either standard popup or xp-style button:
    Select Case m_Style
        Case 0  ' Standard style
            Select Case intState
                Case 0  ' Normal, no border
                    If m_Icon <> 0 Then
                        DrawIconEx hDC, xpos, ypos, m_Icon, m_IconSize, m_IconSize, 0, 0, DI_NORMAL
                    End If
                Case 1  ' MouseOver state:
                    If m_Icon <> 0 Then
                        brsh = CreateSolidBrush(RGB(136, 141, 157))
                        lR = DrawState(hDC, brsh, 0, m_Icon, 0, xpos, ypos, m_IconSize, m_IconSize, DST_ICON Or DSS_MONO)
                        DeleteObject brsh
                        lR = DrawState(hDC, 0, 0, m_Icon, 0, xpos - 1, ypos - 1, m_IconSize, m_IconSize, DST_ICON Or DSS_NORMAL)
                    End If
                    ' Draw the border raised:
                    SetRect bar, Left_x, Top_y, Right_x, Bottom_y
                    Call DrawEdge(hDC, bar, EDGE_RAISED, BF_RECT)
                        
                Case 2  ' MouseDown state:
                    If m_Icon <> 0 Then
                        DrawIconEx hDC, xpos, ypos, m_Icon, m_IconSize, m_IconSize, 0, 0, DI_NORMAL
                    End If
                    ' Draw the border pressed:
                    SetRect bar, Left_x, Top_y, Right_x, Bottom_y
                    Call DrawEdge(hDC, bar, EDGE_SUNKEN, BF_RECT)
                        
                Case 3  ' Draw the Value state when set to ibPressed:
                    If m_Icon <> 0 Then
                        DrawIconEx hDC, xpos, ypos, m_Icon, m_IconSize, m_IconSize, 0, 0, DI_NORMAL
                    End If
                    ' Draw the border pressed:
                    SetRect bar, Left_x, Top_y, Right_x, Bottom_y
                    Call DrawEdge(hDC, bar, EDGE_SUNKEN, BF_RECT)
                        
                Case 4  ' Draw disabled:
                    If m_Icon <> 0 Then
                        lR = DrawState(hDC, 0, 0, m_Icon, 0, xpos, ypos, m_IconSize, m_IconSize, DST_ICON Or DSS_DISABLED)
                    End If
            End Select
            
        Case 1  'XP style
            Select Case intState
                Case 0  ' Draw icon normal with no border:
                    If m_Icon <> 0 Then
                        DrawIconEx hDC, xpos, ypos, m_Icon, m_IconSize, m_IconSize, 0, 0, DI_NORMAL
                    End If
                    
                Case 1  ' Draw MouseOver state:
                    SetRect bar, Left_x, Top_y, Right_x, Bottom_y
                    brsh = CreateSolidBrush(BlendColor(vbHighlight, vbWindowBackground, 80))
                    FillRect hDC, bar, brsh
                    DeleteObject brsh
                    SetRect bar, Left_x, Top_y, Right_x, Bottom_y
                    clr = TranslateColor(vbHighlight)
                    brsh = CreateSolidBrush(clr)
                    FrameRect hDC, bar, brsh
                    DeleteObject brsh
                    If m_Icon <> 0 Then
                        brsh = CreateSolidBrush(RGB(136, 141, 157))
                        lR = DrawState(hDC, brsh, 0, m_Icon, 0, xpos + 1, ypos + 1, m_IconSize, m_IconSize, DST_ICON Or DSS_MONO)
                        DeleteObject brsh
                        lR = DrawState(hDC, 0, 0, m_Icon, 0, xpos - 1, ypos - 1, m_IconSize, m_IconSize, DST_ICON Or DSS_NORMAL)
                    End If
                
                Case 2  ' Draw MouseDown state:
                    SetRect bar, Left_x, Top_y, Right_x, Bottom_y
                    brsh = CreateSolidBrush(BlendColor(vbHighlight, vbButtonFace, 80))
                    FillRect hDC, bar, brsh
                    DeleteObject brsh
                    SetRect bar, Left_x, Top_y, Right_x, Bottom_y
                    clr = TranslateColor(vbHighlight)
                    brsh = CreateSolidBrush(clr)
                    FrameRect hDC, bar, brsh
                    DeleteObject brsh
                    If m_Icon <> 0 Then
                        DrawIconEx hDC, xpos, ypos, m_Icon, m_IconSize, m_IconSize, 0, 0, DI_NORMAL
                    End If
                    
                Case 3  ' Draw the Value state when set to ibPressed:
                    SetRect bar, Left_x, Top_y, Right_x, Bottom_y
                    brsh = CreateSolidBrush(BlendColor(vbButtonFace, vbWindowBackground))
                    FillRect hDC, bar, brsh
                    DeleteObject brsh
                    SetRect bar, Left_x, Top_y, Right_x, Bottom_y
                    clr = TranslateColor(vbHighlight)
                    brsh = CreateSolidBrush(clr)
                    FrameRect hDC, bar, brsh
                    DeleteObject brsh
                    If m_Icon <> 0 Then
                        DrawIconEx hDC, xpos, ypos, m_Icon, m_IconSize, m_IconSize, 0, 0, DI_NORMAL
                    End If
                    
                Case 4  ' Draw disabled
                    If m_Icon <> 0 Then
                        lR = DrawState(hDC, 0, 0, m_Icon, 0, xpos, ypos, m_IconSize, m_IconSize, DST_ICON Or DSS_DISABLED)
                    End If
            End Select
    End Select

End Sub

Private Sub ValidatePicture()
    Dim bValid As Boolean
    If Not (picSource.Picture = 0) Then
        bValid = (picSource.Height = 16) Or (picSource.Height = 32)
        If picSource.Picture.Type <> vbPicTypeIcon Then bValid = False
        If bValid Then
            m_Icon = picSource.Picture
            m_IconSize = picSource.Height
        Else
            MsgBox "Valid picture must be an icon, size 16 x 16 or 32 x 32", vbExclamation, "IconButton Control"
            Set picSource.Picture = LoadPicture("")
            m_Icon = 0
            m_IconSize = 0
        End If
        m_bPrivateResize = True
        Call UserControl_Resize
        m_bPrivateResize = False
    Else
        m_Icon = 0
        m_IconSize = 0
    End If
End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, _
                                ByVal oColorTo As OLE_COLOR, _
                                Optional ByVal alpha As Long = 128) As Long
Dim lCFrom As Long
Dim lCTo As Long
   lCFrom = TranslateColor(oColorFrom)
   lCTo = TranslateColor(oColorTo)
Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000
     
   
   BlendColor = RGB( _
      ((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), _
      ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), _
      ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255) _
      )
      
End Property

'### INITIALIZE/READ PROPERTY VALUES ###
Private Sub UserControl_InitProperties()
    m_Style = m_def_Style
    m_Value = m_def_Value
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    Set picSource.Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Style = PropBag.ReadProperty("ButtonStyle", m_def_Style)
    
    ValidatePicture
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Picture", picSource.Picture, Nothing)
    Call PropBag.WriteProperty("ButtonStyle", m_Style, m_def_Style)
End Sub

Sub ShowAboutBox()
Attribute ShowAboutBox.VB_Description = "Show the About Box for the control."
Attribute ShowAboutBox.VB_UserMemId = -552
Attribute ShowAboutBox.VB_MemberFlags = "40"
    Dim Msg As String
    Msg = "IconButton Control with standard or XP style" & vbNewLine & vbNewLine & _
        "By Brent Culpepper  (IDontKnow)" & vbNewLine & _
        "Copyright© 2004 All rights reserved" & vbNewLine & vbNewLine & _
        "Thanks to Steve McMahon and vbAccelerator for the " & vbNewLine & _
        "mouse tracking class used in this control. This control " & vbNewLine & _
        "also uses the vbAccelerator subclassing method."

    Call MsgBox(Msg, vbInformation, "About IconButton")
End Sub


