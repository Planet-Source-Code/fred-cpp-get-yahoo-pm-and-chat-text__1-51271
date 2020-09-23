VERSION 5.00
Begin VB.UserControl ISButton 
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4260
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   284
   ToolboxBitmap   =   "ISButton.ctx":0000
   Begin VB.Timer timUpdate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   180
   End
End
Attribute VB_Name = "ISButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''  Fred.cpp Alfredo Córdova Pérez 27-julio-2001
'****************************************************************'
'*                                                              *'
'*     Control:     ISButton                                    *'
'*                                                              *'
'*     Author:      Fred cpp
'*                  fred_cpp@msn.com
'*                                                              *'
'*     Description:
'*     ISButton Is a Multi -Style Button that has some Extra
'*     properties:
'*
'*     * Style                 ' Select your Style !
'*     * BackColor
'*     * HoverColor
'*     * FontColor
'*     * FontHoverColor
'*     * Icon
'*     * HoverIcon
'*     * SmallIcon
'*     * IconAlign
'*     * Caption Align
'*     * ShowFocus
'*     *
'*     *
'*  Note:   Not all styles support all properties.
'*          I don't like the MSOXP Style with Focus rect, or
'*          the Standard Style with HoverColor.
'****************************************************************'
'* You Can Use and distribute this control but please give credit


Option Explicit

Private Type POINTAPI
        X As Long
        y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum AlignPosition
    ISAlignLeft
    ISAlignRight
    ISAlignTop
    ISAlignBottom
    ISAlignCenter
End Enum

Public Enum ISButtonStyle
    ISBTNStandard
    ISBTNSoft
    ISBTNFlat
    ISBTNOficceXP
    ISBTNWINXP
    ISBTNVGradient
    ISBTNHGradient
End Enum

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


'' Image type
Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4

' ' State type
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4

Private Const DTA_LEFT = DT_SINGLELINE Or DT_LEFT Or DT_VCENTER
Private Const DTA_RIGHT = DT_SINGLELINE Or DT_RIGHT Or DT_VCENTER
Private Const DTA_TOP = DT_SINGLELINE Or DT_TOP Or DT_CENTER
Private Const DTA_BOTTOM = DT_SINGLELINE Or DT_BOTTOM Or DT_CENTER
Private Const DTA_CENTER = DT_SINGLELINE Or DT_CENTER Or DT_VCENTER

Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_MENU = 4
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_WINDOWTEXT = 8

Private OnClicking As Boolean
Private bFocused As Boolean
Private OnFocus As Boolean
Private InOut As Boolean
Private rtFocus As RECT
Private rtIcon As RECT
Private iState As Integer
'Private PE As ascPaintEffects
'Dim MB As New ascMemoryBitmap

    ' State Values Are:
    '   0 = Normal
    '   1 = Hover
    '   2 = Pressed
    '   3 = Disabled?<<PENDIENTE: Agregar Estado Disabled a todos los estilos
    
'Default Property Values:
Const m_def_Style = 0
Const m_def_ShowFocus = False
Const m_def_SmallIcon = True
Const m_def_FontColor = 0
Const m_def_BackColor = &HE0E0E0
Const m_def_HoverColor = &HD1ADAD
Const m_def_FontHoverColor = &HC00000
Const m_def_Caption = "Caption"
Const m_def_CaptionAlign = 0
Const m_def_IconAlign = 0
Const m_def_IconSize = 16
Const m_def_GColor1 = &HFF0000
Const m_def_GColor2 = &HFFFF
Const m_def_XPDisabledColor = &HE0F0F0
Const m_def_XPDisabledBorderColor = &HB0C0C0


'Property Variables:
Dim m_Style As Integer
Dim m_Backcolor As OLE_COLOR
Dim m_HoverColor As OLE_COLOR
Dim m_FontColor As OLE_COLOR
Dim m_FontHoverColor As OLE_COLOR
Dim m_ShowFocus As Boolean
Dim m_MaskColor As OLE_COLOR
Dim m_UseMaskColor As Boolean
Dim m_Icon As Picture
Dim m_HoverIcon As Picture
Dim m_SmallIcon As Boolean
Dim m_Caption As String
Dim m_CaptionAlign As Integer
Dim m_IconAlign As Single
Dim m_IconSize As Integer
Dim m_GColor1 As OLE_COLOR
Dim m_GColor2 As OLE_COLOR

'Event Declarations:
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Event Resize()
Event MouseHover()
Event MouseOut()

Private Sub APILine(x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As OLE_COLOR)
    'Use the API LineTo for Fast Drawing
    Dim pt As POINTAPI
    UserControl.ForeColor = lColor
    MoveToEx UserControl.hdc, x1, y1, pt
    LineTo UserControl.hdc, x2, y2
End Sub

'Make Soft a color
Function SoftColor(lColor As OLE_COLOR) As OLE_COLOR
    Dim lRed As OLE_COLOR
    Dim lGreen As OLE_COLOR
    Dim lBlue As OLE_COLOR
    Dim lr As OLE_COLOR, lg As OLE_COLOR, lb As OLE_COLOR
    lr = (lColor And &HFF)
    lg = ((lColor And 65280) \ 256)
    lb = ((lColor) And 16711680) \ 65536
    lRed = (76 - Int(((lColor And &HFF) + 32) \ 64) * 19)
    lGreen = (76 - Int((((lColor And 65280) \ 256) + 32) \ 64) * 19)
    lBlue = (76 - Int((((lColor And &HFF0000) \ &H10000) + 32) / 64) * 19)
    SoftColor = RGB(lr + lRed, lg + lGreen, lb + lBlue)
End Function

'Offset a color
Function OffsetColor(lColor As OLE_COLOR, lOffset As Long) As OLE_COLOR
    Dim lRed As OLE_COLOR
    Dim lGreen As OLE_COLOR
    Dim lBlue As OLE_COLOR
    Dim lr As OLE_COLOR, lg As OLE_COLOR, lb As OLE_COLOR
    lr = (lColor And &HFF)
    lg = ((lColor And 65280) \ 256)
    lb = ((lColor) And 16711680) \ 65536
    lRed = (lOffset + lr)
    lGreen = (lOffset + lg)
    lBlue = (lOffset + lb)
    If lRed > 255 Then lRed = 255
    If lRed < 0 Then lRed = 0
    If lGreen > 255 Then lGreen = 255
    If lGreen < 0 Then lGreen = 0
    If lBlue > 255 Then lBlue = 255
    If lBlue < 0 Then lBlue = 0
    OffsetColor = RGB(lRed, lGreen, lBlue)
End Function

''Draw Button Image.
Private Sub fDrawPicture( _
    ByRef m_Picture As StdPicture, _
    ByVal X As Long, _
    ByVal y As Long, ByVal W As Long, ByVal H As Long, _
    Optional ByVal bShadow As Boolean = False, Optional ByVal Disabled As Boolean = False)
On Error Resume Next
     
    Dim lFlags As Long
    Dim hBrush As Long
         
    Select Case m_Picture.Type
        Case vbPicTypeBitmap
            lFlags = DST_BITMAP
        Case vbPicTypeIcon
            lFlags = DST_ICON
        Case Else
            lFlags = DST_COMPLEX
    End Select

    If bShadow Then
        hBrush = CreateSolidBrush(&H9C8181)  'RGB(128, 128, 128))
    End If
    If Disabled Then
     DrawState UserControl.hdc, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, X, y, W, _
        H, _
        lFlags Or DSS_DISABLED
    Else
     DrawState UserControl.hdc, IIf(bShadow, hBrush, 0), 0, m_Picture.Handle, 0, X, y, W, _
        H, _
        lFlags Or IIf(bShadow, DSS_MONO, DSS_NORMAL)
     End If
    If bShadow Then
        DeleteObject hBrush
    End If
End Sub

Private Sub DrawVGradient(lEndColor As Long, lStartcolor As Long)
    ''Draw a Vertical Gradient in the current HDC
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim lh As Long, lw As Long
    Dim ni As Long
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / lh
    dG = (sG - eG) / lh
    dB = (sB - eB) / lh
    For ni = 0 To lh
        APILine 0, ni, lw, ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next ni
End Sub

Private Sub DrawHGradient(lEndColor As Long, lStartcolor As Long)
    ''Draw a Horizontal Gradient in the current HDC
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim lh As Long, lw As Long
    Dim ni As Long
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / lw
    dG = (sG - eG) / lw
    dB = (sB - eB) / lw
    For ni = 0 To lw
        APILine ni, 0, ni, lh, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next ni
End Sub

' Commons Functions Support
Private Function InBox(ObjectHWnd As Long) As Boolean
    Dim mpos As POINTAPI
    Dim oRect As RECT
    GetCursorPos mpos
    GetWindowRect ObjectHWnd, oRect
    If mpos.X >= oRect.Left And mpos.X <= oRect.Right And _
        mpos.y >= oRect.Top And mpos.y <= oRect.Bottom Then
        InBox = True
    Else
        InBox = False
   End If
End Function


Private Sub timUpdate_Timer()
    If InBox(UserControl.hWnd) Then
        If InOut = False Then
            iState = 1
            'bFocused = True
            UserControl_Paint
            RaiseEvent MouseHover
        End If
    InOut = True
    Else
        If InOut Then
            iState = 0
            UserControl_Paint
            timUpdate.Enabled = False
            RaiseEvent MouseOut
        End If
        InOut = False
    End If

End Sub


'****************************************************************'
'*                                                              *'
'*     Control:     ISButton                                    *'
'*                                                              *'
'*     Section:     Properties                                  *'
'*                                                              *'
'****************************************************************'

'''Backcolor Property
Public Property Get Backcolor() As OLE_COLOR
    Backcolor = m_Backcolor
End Property

Public Property Let Backcolor(ByVal New_BackColor As OLE_COLOR)
    m_Backcolor = New_BackColor
    UserControl_Paint
    PropertyChanged "BackColor"
End Property
'''GColor1 Property
Public Property Get GColor1() As OLE_COLOR
    GColor1 = m_GColor1
End Property

Public Property Let GColor1(ByVal New_GColor1 As OLE_COLOR)
    m_GColor1 = New_GColor1
    UserControl_Paint
    PropertyChanged "GColor1"
End Property

'''GColor2 Property
Public Property Get GColor2() As OLE_COLOR
    GColor2 = m_GColor2
End Property

Public Property Let GColor2(ByVal New_GColor2 As OLE_COLOR)
    m_GColor2 = New_GColor2
    UserControl_Paint
    PropertyChanged "GColor2"
End Property

'''''''Set Icon Size
Public Property Get IconSize() As Integer
    IconSize = m_IconSize
End Property

Public Property Let IconSize(ByVal New_IconSize As Integer)
    m_IconSize = New_IconSize
    UserControl_Resize
    PropertyChanged "IconSize"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    If New_Enabled Then
        iState = 0
    Else
        iState = 3
    End If
    UserControl_Paint
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    UserControl_Paint
    PropertyChanged "Font"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Devuelve un controlador (de Microsoft Windows) a la ventana de un objeto."
    hWnd = UserControl.hWnd
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Establece un icono personalizado para el mouse."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Devuelve o establece el tipo de puntero del mouse mostrado al pasar por encima de un objeto."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Style() As ISButtonStyle
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As ISButtonStyle)
    m_Style = New_Style
    UserControl_Paint
    PropertyChanged "Style"
End Property

Public Property Get HoverColor() As OLE_COLOR
    HoverColor = m_HoverColor
End Property

Public Property Let HoverColor(ByVal New_HoverColor As OLE_COLOR)
    m_HoverColor = New_HoverColor
    PropertyChanged "HoverColor"
End Property

Public Property Get FontColor() As OLE_COLOR
    FontColor = m_FontColor
End Property

Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
    m_FontColor = New_FontColor
    UserControl_Paint
    PropertyChanged "FontColor"
End Property

Public Property Get FontHoverColor() As OLE_COLOR
    FontHoverColor = m_FontHoverColor
End Property

Public Property Let FontHoverColor(ByVal New_FontHoverColor As OLE_COLOR)
    m_FontHoverColor = New_FontHoverColor
    PropertyChanged "FontHoverColor"
End Property

Public Property Get ShowFocus() As Boolean
    ShowFocus = m_ShowFocus
End Property

Public Property Let ShowFocus(ByVal New_ShowFocus As Boolean)
    m_ShowFocus = New_ShowFocus
    UserControl_Paint
    PropertyChanged "ShowFocus"
End Property

Public Property Get Icon() As Picture
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set m_Icon = New_Icon
    UserControl_Paint
    PropertyChanged "Icon"
End Property

Public Property Get HoverIcon() As Picture
    Set HoverIcon = m_HoverIcon
End Property

Public Property Set HoverIcon(ByVal New_HoverIcon As Picture)
    Set m_HoverIcon = New_HoverIcon
    PropertyChanged "HoverIcon"
End Property

Public Property Get SmallIcon() As Boolean
    SmallIcon = m_SmallIcon
End Property

Public Property Let SmallIcon(ByVal New_SmallIcon As Boolean)
    m_SmallIcon = New_SmallIcon
    UserControl_Paint
    PropertyChanged "SmallIcon"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    UserControl_Paint
    PropertyChanged "Caption"
End Property

Public Property Get CaptionAlign() As AlignPosition
    CaptionAlign = m_CaptionAlign
End Property

Public Property Let CaptionAlign(ByVal New_CaptionAlign As AlignPosition)
    m_CaptionAlign = New_CaptionAlign
    UserControl_Paint
    PropertyChanged "CaptionAlign"
End Property

Public Property Get IconAlign() As AlignPosition
    IconAlign = m_IconAlign
End Property

Public Property Let IconAlign(ByVal New_IconAlign As AlignPosition)
    m_IconAlign = New_IconAlign
    UserControl_Resize
    PropertyChanged "IconAlign"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    'initialize Drawing objects
'    MB.CreateByResource "DITHER"
'    Set PE = New ascPaintEffects
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    'If PE Is Nothing Then Set PE = New ascPaintEffects
    m_MaskColor = vbMagenta
    m_Style = m_def_Style
    m_Backcolor = GetSysColor(COLOR_BTNFACE)
    m_HoverColor = m_def_HoverColor
    m_FontColor = m_def_FontColor
    m_FontHoverColor = m_def_FontHoverColor
    m_ShowFocus = m_def_ShowFocus
    Set m_Icon = LoadPicture("")
    Set m_HoverIcon = LoadPicture("")
    m_SmallIcon = m_def_SmallIcon
    m_Caption = m_def_Caption
    m_CaptionAlign = m_def_CaptionAlign
    m_IconAlign = m_def_IconAlign
    m_IconSize = m_def_IconSize
    m_CaptionAlign = 4
    m_GColor1 = m_def_GColor1
    m_GColor2 = m_def_GColor2
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Backcolor = PropBag.ReadProperty("BackColor", GetSysColor(COLOR_BTNFACE))
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_HoverColor = PropBag.ReadProperty("HoverColor", m_def_HoverColor)
    m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    m_FontHoverColor = PropBag.ReadProperty("FontHoverColor", m_def_FontHoverColor)
    m_ShowFocus = PropBag.ReadProperty("ShowFocus", m_def_ShowFocus)
    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
    Set m_HoverIcon = PropBag.ReadProperty("HoverIcon", Nothing)
    m_SmallIcon = PropBag.ReadProperty("SmallIcon", m_def_SmallIcon)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_CaptionAlign = PropBag.ReadProperty("CaptionAlign", m_def_CaptionAlign)
    m_IconAlign = PropBag.ReadProperty("IconAlign", m_def_IconAlign)
    m_IconSize = PropBag.ReadProperty("IconSize", m_def_IconSize)
    m_GColor1 = PropBag.ReadProperty("GColor1", m_def_GColor1)
    m_GColor2 = PropBag.ReadProperty("GColor2", m_def_GColor2)
End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_Backcolor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("HoverColor", m_HoverColor, m_def_HoverColor)
    Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
    Call PropBag.WriteProperty("FontHoverColor", m_FontHoverColor, m_def_FontHoverColor)
    Call PropBag.WriteProperty("ShowFocus", m_ShowFocus, m_def_ShowFocus)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("HoverIcon", m_HoverIcon, Nothing)
    Call PropBag.WriteProperty("SmallIcon", m_SmallIcon, m_def_SmallIcon)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("CaptionAlign", m_CaptionAlign, m_def_CaptionAlign)
    Call PropBag.WriteProperty("IconAlign", m_IconAlign, m_def_IconAlign)
    Call PropBag.WriteProperty("IconSize", m_IconSize, m_def_IconSize)
    Call PropBag.WriteProperty("GColor1", m_GColor1, m_def_GColor1)
    Call PropBag.WriteProperty("GColor2", m_GColor2, m_def_GColor2)
End Sub

'****************************************************************'
'*                                                              *'
'*     Control:     ISButton                                    *'
'*                                                              *'
'*     Section:     Events                                      *'
'*                                                              *'
'****************************************************************'

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    iState = 2
    UserControl_Paint
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent Click
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
        iState = 1
        UserControl_Paint
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then
        iState = 2
        UserControl_Paint
    End If
    RaiseEvent MouseDown(Button, Shift, X, y)
End Sub

'Private Sub UserControl_Click()
'    RaiseEvent Click
'End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If UserControl.Enabled And Button = 0 Then
        timUpdate.Enabled = True
    End If
    RaiseEvent MouseMove(Button, Shift, X, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If UserControl.Enabled And Button = vbLeftButton Then
        iState = 1
        UserControl_Paint
    End If
    RaiseEvent MouseUp(Button, Shift, X, y)
    timUpdate.Enabled = True
End Sub


Private Sub UserControl_GotFocus()
    'OnFocus
    OnFocus = True
    UserControl_Paint
End Sub

Private Sub UserControl_LostFocus()
    '
    OnFocus = False
    UserControl_Paint
End Sub

'****************************************************************'
'*       Resize Efects.                                         *'
Private Sub UserControl_Resize()
    rtFocus.Left = 4
    rtFocus.Right = UserControl.ScaleWidth - 4
    rtFocus.Top = 4
    rtFocus.Bottom = UserControl.ScaleHeight - 4
    '''LSet rtIcon = rtFocus
    Select Case m_IconAlign
    Case 0: 'Left
        rtIcon.Top = (UserControl.ScaleHeight - m_IconSize) / 2
        rtIcon.Left = rtFocus.Left
    Case 1: 'Right
        rtIcon.Top = (UserControl.ScaleHeight - m_IconSize) / 2
        rtIcon.Left = (UserControl.ScaleWidth - m_IconSize) - 4
    Case 2: 'Top
        rtIcon.Top = 4 '(UserControl.ScaleHeight - m_IconSize) / 2
        rtIcon.Left = (UserControl.ScaleWidth - m_IconSize) / 2
    Case 3: 'Bottom
        rtIcon.Top = (UserControl.ScaleHeight - m_IconSize) - 4
        rtIcon.Left = (UserControl.ScaleWidth - m_IconSize) / 2
    Case 4: 'Center
        rtIcon.Top = (UserControl.ScaleHeight - m_IconSize) / 2
        rtIcon.Left = (UserControl.ScaleWidth - m_IconSize) / 2
    End Select
    UserControl_Paint
    RaiseEvent Resize
End Sub

'****************************************************************'
'*       Painting Efects.                                       *'
'*  All Painting Efects Are In This Sub!                        *'
'****************************************************************'
Private Sub UserControl_Paint()
    Dim lh As Long, lw As Long
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
    Dim ni As Long
    Dim lStep As Long 'single
    Dim tempColor As OLE_COLOR
    Dim ParentBackColor As OLE_COLOR
    Dim imgOffset As Integer
    ParentBackColor = UserControl.Extender.Parent.Backcolor
    tempColor = OffsetColor(m_Backcolor, &H30)
    lStep = 25 / lh
    Select Case m_Style
        Case 0  'Standard Button
            Select Case iState
            Case 0: 'Normal
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine lw - 2, 2, lw - 2, lh, GetSysColor(COLOR_BTNSHADOW)
                APILine 2, lh - 2, lw, lh - 2, GetSysColor(COLOR_BTNSHADOW)
                APILine lw - 1, 1, lw - 1, lh, vbBlack
                APILine 1, lh - 1, lw, lh - 1, vbBlack
                If m_ShowFocus And OnFocus Then
                    APILine lw - 3, 3, lw - 3, lh - 2, GetSysColor(COLOR_BTNSHADOW)
                    APILine 3, lh - 3, lw - 2, lh - 3, GetSysColor(COLOR_BTNSHADOW)
                    DrawFocusRect UserControl.hdc, rtFocus
                End If
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 1: 'Hover
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor 'm_HoverColor
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine lw - 2, 2, lw - 2, lh, GetSysColor(COLOR_BTNSHADOW)
                APILine 2, lh - 2, lw, lh - 2, GetSysColor(COLOR_BTNSHADOW)
                APILine lw - 1, 1, lw - 1, lh, vbBlack
                APILine 1, lh - 1, lw, lh - 1, vbBlack
                If m_ShowFocus And OnFocus Then
                    APILine lw - 3, 3, lw - 3, lh - 2, GetSysColor(COLOR_BTNSHADOW)
                    APILine 3, lh - 3, lw - 2, lh - 3, GetSysColor(COLOR_BTNSHADOW)
                    DrawFocusRect UserControl.hdc, rtFocus
                End If
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 2: 'Pressed
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor ' m_HoverColor
                APILine 0, 0, lw - 1, 0, vbBlack
                APILine 0, 0, 0, lh - 1, vbBlack
                APILine lw - 1, 0, lw - 1, lh, vbBlack
                APILine 0, lh - 1, lw, lh - 1, vbBlack
                APILine 1, 1, lw - 2, 1, GetSysColor(COLOR_BTNSHADOW)
                APILine 1, 1, 1, lh - 2, GetSysColor(COLOR_BTNSHADOW)
                APILine lw - 2, 1, lw - 2, lh - 1, GetSysColor(COLOR_BTNSHADOW)
                APILine 1, lh - 2, lw - 1, lh - 2, GetSysColor(COLOR_BTNSHADOW)
                If m_ShowFocus And OnFocus Then
                    DrawFocusRect UserControl.hdc, rtFocus
                End If
                'fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, False, False
            Case 3 ' Disabled
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine lw - 2, 2, lw - 2, lh, GetSysColor(COLOR_BTNSHADOW)
                APILine 2, lh - 2, lw, lh - 2, GetSysColor(COLOR_BTNSHADOW)
                APILine lw - 1, 1, lw - 1, lh, vbBlack
                APILine 1, lh - 1, lw, lh - 1, vbBlack
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, True
            End Select
        Case 1  ' Soft Style
            Select Case iState
            Case 0 ' Normal
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine lw - 1, 0, lw - 1, lh, GetSysColor(COLOR_BTNSHADOW)
                APILine 0, lh - 1, lw, lh - 1, GetSysColor(COLOR_BTNSHADOW)
                If m_ShowFocus And OnFocus Then
                    DrawFocusRect UserControl.hdc, rtFocus
                End If
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 1 ' Hover
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor 'm_HoverColor
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine lw - 1, 0, lw - 1, lh, vbBlack
                APILine 0, lh - 1, lw, lh - 1, vbBlack
                APILine lw - 2, 0, lw - 2, lh - 1, GetSysColor(COLOR_BTNSHADOW)
                APILine 0, lh - 2, lw - 1, lh - 2, GetSysColor(COLOR_BTNSHADOW)
                If m_ShowFocus And OnFocus Then
                    DrawFocusRect UserControl.hdc, rtFocus
                End If
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 2 ' Pressed
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNSHADOW)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNSHADOW)
                APILine lw - 1, 0, lw - 1, lh, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, lh - 1, lw, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                If m_ShowFocus And OnFocus Then
                    DrawFocusRect UserControl.hdc, rtFocus
                End If
                'fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, False, False
            Case 3 ' Disabled
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine lw - 1, 0, lw - 1, lh, GetSysColor(COLOR_BTNSHADOW)
                APILine 0, lh - 1, lw, lh - 1, GetSysColor(COLOR_BTNSHADOW)
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, True
            End Select
        Case 2  ' Flat Style
            Select Case iState
            Case 0 ' Normal
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor
                If m_ShowFocus And OnFocus Then
                    DrawFocusRect UserControl.hdc, rtFocus
                End If
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 1 ' Hover
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine lw - 1, 0, lw - 1, lh, SoftColor(vbBlack)
                APILine 0, lh - 1, lw, lh - 1, SoftColor(vbBlack)
                If m_ShowFocus And OnFocus Then
                    DrawFocusRect UserControl.hdc, rtFocus
                End If
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 2 ' Pressed
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNSHADOW)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNSHADOW)
                APILine lw - 1, 0, lw - 1, lh, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, lh - 1, lw, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                If m_ShowFocus And OnFocus Then
                    DrawFocusRect UserControl.hdc, rtFocus
                End If
                'fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, False, False
            Case 3 ' Disabled
                UserControl.Cls
                UserControl.Backcolor = m_Backcolor
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, True
            End Select
        Case 3  ' OficeXP
            Dim lBorderColor As Long
            lBorderColor = RGB(0, 0, 128)
            Select Case iState
            Case 0 ' Normal
                UserControl.Cls
                UserControl.Backcolor = SoftColor(m_Backcolor)
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 1 ' Hover
                UserControl.Cls
                UserControl.Backcolor = m_HoverColor
                APILine 0, 0, lw - 1, 0, lBorderColor
                APILine 0, 0, 0, lh - 1, lBorderColor
                APILine lw - 1, 0, lw - 1, lh, lBorderColor
                APILine 0, lh - 1, lw, lh - 1, lBorderColor
                'APILine lw - 2, 1, lw - 2, lh - 1, lBorderColor
                'APILine 1, lh - 2, lw - 1, lh - 2, lBorderColor
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, True, False
                'fDrawPicture m_Icon, rtIcon.Left - 2, rtIcon.Top - 2, m_IconSize, m_IconSize, False, False
            Case 2 ' Pressed
                UserControl.Cls
                UserControl.Backcolor = RGB(128, 128, 192)
                APILine 0, 0, lw - 1, 0, lBorderColor
                APILine 0, 0, 0, lh - 1, lBorderColor
                APILine lw - 1, 0, lw - 1, lh, lBorderColor
                APILine 0, lh - 1, lw, lh - 1, lBorderColor
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 3 ' Disabled
                lBorderColor = RGB(198, 198, 198)
                UserControl.Cls
                UserControl.Backcolor = SoftColor(m_Backcolor)
                APILine 0, 0, lw - 1, 0, lBorderColor
                APILine 0, 0, 0, lh - 1, lBorderColor
                APILine lw - 1, 0, lw - 1, lh, lBorderColor
                APILine 0, lh - 1, lw, lh - 1, lBorderColor
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, True
            End Select
        Case 4  ' Win XP Style
            Select Case iState
            Case 0 ' Normal
                UserControl.Cls
                tempColor = OffsetColor(m_Backcolor, &H30)
                lStep = 25 / lh
                tempColor = OffsetColor(m_Backcolor, &H30)
                For ni = 0 To lh
                    APILine 0, ni, lw, ni, OffsetColor(tempColor, -ni * lStep)
                Next ni
                APILine 1, lh - 2, lw - 1, lh - 2, OffsetColor(tempColor, -64)
                APILine 1, lh - 3, lw - 1, lh - 3, OffsetColor(tempColor, -32)
                APILine lw - 2, 1, lw - 2, lh - 1, OffsetColor(tempColor, -64)
                APILine lw - 3, 2, lw - 3, lh - 2, OffsetColor(tempColor, -32)
                APILine 2, 1, lw - 2, 1, OffsetColor(tempColor, 19)
                APILine 1, 2, lw - 2, 2, OffsetColor(tempColor, 32)
                APILine 1, 2, 1, lh - 2, OffsetColor(tempColor, 4)
                APILine 2, 3, 2, lh - 3, OffsetColor(tempColor, -4)

                APILine 1, 0, lw - 1, 0, &H733C00
                APILine 0, 1, 0, lh - 1, &H733C00
                APILine lw - 1, 1, lw - 1, lh - 1, &H733C00
                APILine 1, lh - 1, lw - 1, lh - 1, &H733C00
                UserControl.PSet (0, 0), ParentBackColor
                UserControl.PSet (0, lh - 1), ParentBackColor
                UserControl.PSet (lw - 1, 0), ParentBackColor
                UserControl.PSet (lw - 1, lh - 1), ParentBackColor

                UserControl.PSet (1, 1), &H7B4D10
                UserControl.PSet (1, lh - 2), &H7B4D10
                UserControl.PSet (lw - 2, 1), &H7B4D10
                UserControl.PSet (lw - 2, lh - 2), &H7B4D10
                If OnFocus Then
                    ' Top
                    APILine 1, 1, lw - 1, 1, &HFFE7CE
                    APILine 1, 2, lw - 1, 2, &HF7D7BD
                    'Bottom
                    APILine 1, lh - 2, lw - 1, lh - 2, &HEF826B
                    APILine 2, lh - 3, lw - 2, lh - 3, &HE7AE8C

                    APILine lw - 2, 2, lw - 2, lh - 2, &HE7AE8C
                    APILine lw - 3, 3, lw - 3, lh - 3, &HF0D1B5

                    APILine 2, 3, 2, lh - 3, &HF0D1B5
                    APILine 1, 3, 1, lh - 2, &HE7AE8C
                End If
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 1 ' Hover
                tempColor = OffsetColor(m_Backcolor, &H30)
                lStep = 25 / lh
                For ni = 0 To lh
                    APILine 0, ni, lw, ni, OffsetColor(tempColor, -ni * lStep)
                Next ni
                APILine 1, 0, lw - 1, 0, &H733C00
                APILine 0, 1, 0, lh - 1, &H733C00
                APILine lw - 1, 1, lw - 1, lh - 1, &H733C00
                APILine 1, lh - 1, lw - 1, lh - 1, &H733C00
                APILine 1, lh - 2, lw - 1, lh - 2, OffsetColor(tempColor, -64)
                APILine 1, lh - 3, lw - 1, lh - 3, OffsetColor(tempColor, -32)
                APILine lw - 2, 1, lw - 2, lh - 1, OffsetColor(tempColor, -64)
                APILine lw - 3, 2, lw - 3, lh - 2, OffsetColor(tempColor, -32)
                APILine 2, 1, lw - 2, 1, OffsetColor(tempColor, 32)
                APILine 1, 2, lw - 2, 2, OffsetColor(tempColor, 64)
                APILine 1, 2, 1, lh - 2, OffsetColor(tempColor, 64)
                APILine 2, 3, 2, lh - 3, OffsetColor(tempColor, 32)
                UserControl.PSet (0, 0), ParentBackColor
                UserControl.PSet (0, lh - 1), ParentBackColor
                UserControl.PSet (lw - 1, 0), ParentBackColor
                UserControl.PSet (lw - 1, lh - 1), ParentBackColor
                'Top
                APILine 1, 1, lw - 1, 1, &HCEF3FF
                APILine 1, 2, lw - 1, 2, &H6BCBFF
                'Bottom
                APILine 1, lh - 2, lw - 1, lh - 2, &H96E7&
                APILine 2, lh - 3, lw - 2, lh - 3, &H31B2FF
                'Right
                APILine lw - 2, 2, lw - 2, lh - 2, &H31B2FF
                APILine lw - 3, 3, lw - 3, lh - 3, &H6BCBFF
                'Left
                APILine 2, 2, 2, lh - 3, &H6BCBFF
                APILine 1, 2, 1, lh - 2, &H31B2FF
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 2 ' Pressed
                tempColor = OffsetColor(m_Backcolor, &H50) ' &H30)
                tempColor = OffsetColor(tempColor, -48)
                lStep = 25 / lh
                For ni = 0 To lh
                    APILine 0, ni, lw, ni, OffsetColor(tempColor, ni * lStep)
                Next ni
                APILine 1, 0, lw - 1, 0, &H733C00
                APILine 0, 1, 0, lh - 1, &H733C00
                APILine lw - 1, 1, lw - 1, lh - 1, &H733C00
                APILine 1, lh - 1, lw - 1, lh - 1, &H733C00

                APILine 1, lh - 2, lw - 1, lh - 2, OffsetColor(tempColor, 64)
                APILine 1, lh - 3, lw - 1, lh - 3, OffsetColor(tempColor, 16)
                APILine lw - 2, 1, lw - 2, lh - 1, OffsetColor(tempColor, 32)
                APILine lw - 3, 2, lw - 3, lh - 2, OffsetColor(tempColor, 16)
                APILine 1, 1, lw - 1, 1, OffsetColor(tempColor, -16)
                APILine 1, 2, lw - 2, 2, OffsetColor(tempColor, -8)
                APILine 1, 2, 1, lh - 2, OffsetColor(tempColor, -16)
                APILine 2, 3, 2, lh - 3, OffsetColor(tempColor, -8)

                UserControl.PSet (0, 0), ParentBackColor
                UserControl.PSet (0, lh - 1), ParentBackColor
                UserControl.PSet (lw - 1, 0), ParentBackColor
                UserControl.PSet (lw - 1, lh - 1), ParentBackColor
                'fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, False, False
            Case 3 ' Disabled
                UserControl.Backcolor = m_def_XPDisabledColor
                APILine 1, 0, lw - 1, 0, m_def_XPDisabledBorderColor '&H733C00
                APILine 0, 1, 0, lh - 1, m_def_XPDisabledBorderColor
                APILine lw - 1, 1, lw - 1, lh - 1, m_def_XPDisabledBorderColor
                APILine 1, lh - 1, lw - 1, lh - 1, m_def_XPDisabledBorderColor

                UserControl.PSet (1, 1), m_def_XPDisabledBorderColor
                UserControl.PSet (1, lh - 2), m_def_XPDisabledBorderColor
                UserControl.PSet (lw - 2, 1), m_def_XPDisabledBorderColor
                UserControl.PSet (lw - 2, lh - 2), m_def_XPDisabledBorderColor

                UserControl.PSet (0, 0), ParentBackColor
                UserControl.PSet (0, lh - 1), ParentBackColor
                UserControl.PSet (lw - 1, 0), ParentBackColor
                UserControl.PSet (lw - 1, lh - 1), ParentBackColor
                'fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, False, True
            End Select
        Case 5  'VGradient
            Select Case iState
            Case 0 ' Normal
            DrawVGradient m_GColor1, m_GColor2
'                For ni = 0 To lh
'                    APILine 0, ni, lw, ni, OffsetColor(m_BackColor, -ni)
'                Next ni
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 1 ' Hover
'                For ni = 0 To lh
'                    APILine 0, ni, lw, ni, OffsetColor(m_BackColor, -ni)
'                Next ni
                DrawVGradient m_GColor1, m_GColor2
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine lw - 1, 0, lw - 1, lh, GetSysColor(COLOR_BTNSHADOW)
                APILine 0, lh - 1, lw, lh - 1, GetSysColor(COLOR_BTNSHADOW)
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 2 ' Pressed
                DrawVGradient m_GColor1, m_GColor2
'                For ni = 0 To lh
'                    APILine 0, ni, lw, ni, OffsetColor(m_BackColor, -ni)
'                Next ni
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNSHADOW)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNSHADOW)
                APILine lw - 1, 0, lw - 1, lh, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, lh - 1, lw, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                'fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, False, False
            Case 3 ' Disabled
                lBorderColor = RGB(198, 198, 198)
                DrawVGradient m_GColor1, m_GColor2
'                For ni = 0 To lh
'                    APILine 0, ni, lw, ni, OffsetColor(m_BackColor, -ni)
'                Next ni
                APILine 0, 0, lw - 1, 0, lBorderColor
                APILine 0, 0, 0, lh - 1, lBorderColor
                APILine lw - 1, 0, lw - 1, lh, lBorderColor
                APILine 0, lh - 1, lw, lh - 1, lBorderColor
                'fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, False, True
            End Select
        Case 6 ' HGradient
            Select Case iState
            Case 0 ' Normal
                DrawHGradient m_GColor1, m_GColor2
'                For ni = 0 To lh
'                    APILine 0, ni, lw, ni, OffsetColor(m_BackColor, -ni)
'                Next ni
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 1 ' Hover
                DrawHGradient m_GColor1, m_GColor2
'                For ni = 0 To lh
'                    APILine 0, ni, lw, ni, OffsetColor(m_BackColor, -ni)
'                Next ni
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine lw - 1, 0, lw - 1, lh, GetSysColor(COLOR_BTNSHADOW)
                APILine 0, lh - 1, lw, lh - 1, GetSysColor(COLOR_BTNSHADOW)
                'fDrawPicture m_Icon, rtIcon.Left, rtIcon.Top, m_IconSize, m_IconSize, False, False
            Case 2 ' Pressed
                DrawHGradient m_GColor1, m_GColor2
'                For ni = 0 To lh
'                    APILine 0, ni, lw, ni, OffsetColor(m_BackColor, -ni)
'                Next ni
                APILine 0, 0, lw - 1, 0, GetSysColor(COLOR_BTNSHADOW)
                APILine 0, 0, 0, lh - 1, GetSysColor(COLOR_BTNSHADOW)
                APILine lw - 1, 0, lw - 1, lh, GetSysColor(COLOR_BTNHIGHLIGHT)
                APILine 0, lh - 1, lw, lh - 1, GetSysColor(COLOR_BTNHIGHLIGHT)
                'fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, False, False
            Case 3 ' Disabled
                lBorderColor = RGB(198, 198, 198)
                DrawHGradient m_GColor1, m_GColor2
'                For ni = 0 To lh
'                    APILine 0, ni, lw, ni, OffsetColor(m_BackColor, -ni)
'                Next ni
                APILine 0, 0, lw - 1, 0, lBorderColor
                APILine 0, 0, 0, lh - 1, lBorderColor
                APILine lw - 1, 0, lw - 1, lh, lBorderColor
                APILine 0, lh - 1, lw, lh - 1, lBorderColor
                'fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, False, True
            End Select
    End Select
        ''''''''''''''''''''''''
    '' Draw the Icon on button
    If m_Icon Is Nothing Then GoTo NoPicture
    If Enabled = 0 Then
        
        fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize
    Else
        'Is not disabled
        'If MSOXPStyle, then adjust the image offfset
        If m_Style = ISBTNOficceXP Then If iState = 1 Then imgOffset = 2
        If m_Icon.Type = vbPicTypeIcon Then
            'DrawTransparentBitmap doesn't support icons
                If m_Style = ISBTNOficceXP Then fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, True
                 fDrawPicture m_Icon, rtIcon.Left + 1 - imgOffset, rtIcon.Top + 1 - imgOffset, m_IconSize, m_IconSize
                'hdc, P, (PX - imgOffset), (PY - imgOffset), PW, PH, 0, 0
        Else
'            If m_UseMaskColor Then
                If m_Style = ISBTNOficceXP Then fDrawPicture m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, True
                fDrawPicture m_Icon, (rtIcon.Left + 1 - imgOffset), (rtIcon.Top + 1 - imgOffset), m_IconSize, m_IconSize
'            Else
'                If m_Style = ISBTNOficceXP Then PE.PaintMonoPicture hdc, m_Icon, rtIcon.Left + 1, rtIcon.Top + 1, m_IconSize, m_IconSize, 0, 0, m_MaskColor
'                PE.PaintStandardPicture hdc, m_Icon, (rtIcon.Left + 1 - imgOffset), (rtIcon.Top + 1 - imgOffset), m_IconSize, m_IconSize, 0, 0
'            End If
        End If
    End If
NoPicture:
    'Draw Text
    Dim capAlign As Long
    Select Case m_CaptionAlign
        Case 0
            capAlign = DTA_LEFT
        Case 1
            capAlign = DTA_RIGHT
        Case 2
            capAlign = DTA_TOP
        Case 3
            capAlign = DTA_BOTTOM
        Case 4
            capAlign = DTA_CENTER
    End Select
    Select Case iState
    Case 0
        UserControl.ForeColor = m_FontColor
    Case 1
        UserControl.ForeColor = m_FontHoverColor
    Case 2
            UserControl.ForeColor = m_FontHoverColor
    End Select
    If iState = 2 Then
        Dim rtText As RECT
        rtText.Top = rtText.Top + 3
        rtText.Left = rtFocus.Left + 2
        rtText.Bottom = rtFocus.Bottom + 3
        rtText.Right = rtFocus.Right + 2
        DrawText UserControl.hdc, m_Caption, Len(m_Caption), rtText, capAlign
    Else
        DrawText UserControl.hdc, m_Caption, Len(m_Caption), rtFocus, capAlign
    End If
End Sub

