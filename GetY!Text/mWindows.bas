Attribute VB_Name = "mWindows"
Option Explicit

' Find wheather a window is visible or not
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
'

' Determine the Caption of a open window
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) _
As Long
'

' Determine the length of Caption of a open window
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
    (ByVal hWnd As Long) _
As Long
'

' Enumerates all top-level windows on the screen
' by passing the handle to each window, in turn,
' to an application-defined callback function.
Public Declare Function EnumWindows Lib "user32" _
    (ByVal lpEnumFunc _
    As Long, _
    ByVal lparam As Long) _
As Long
'

' The GetClassName function retrieves the name of the class to which the specified window belongs.
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
    (ByVal hWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) _
As Long
'

' The FindWindow function retrieves a handle to the top-level window whose
' class name and window name match the specified strings
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) _
As Long
'


Public sysWnd As Long
Private m_cSink As IEnumWindowsSink

' The The SetLayeredWindowAttributes function sets the opacity and transparency color key of a layered window.
'
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal crKey As Long, _
     ByVal bAlpha As Byte, _
     ByVal dwFlags As Long) _
As Long
'


' The GetWindowLong function retrieves information about the specified window
'
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, _
    ByVal nIndex As Long) _
As Long
'

' The SetWindowLong function changes an attribute of the specified window
'
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) _
As Long

Private Const GWL_EXSTYLE = (-20)   ' Sets a new extended window style
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Public Function Transparent(ByVal hWnd As Long, Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
      Transparent = 1
    Else
      Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
      Msg = Msg Or WS_EX_LAYERED
      SetWindowLong hWnd, GWL_EXSTYLE, Msg
      SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
      Transparent = 0
    End If
    If Err Then
      Transparent = 2
    End If
End Function

Public Function Opaque(ByVal hWnd As Long) As Long
    Dim Msg As Long
    On Error Resume Next
    Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
    Msg = Msg And Not WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
    Opaque = 0
    If Err Then
      Opaque = 2
    End If
End Function

Public Function EnumerateWindows(ByRef cSink As IEnumWindowsSink) As Boolean
    If Not (m_cSink Is Nothing) Then Exit Function
    Set m_cSink = cSink
    EnumWindows AddressOf EnumWindowsProc, cSink.Identifier
    Set m_cSink = Nothing
End Function

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lparam As Long) As Long
    Dim bStop As Boolean
    
    bStop = False
    m_cSink.EnumWindow hWnd, bStop
    If (bStop) Then
        EnumWindowsProc = 0
    Else
        EnumWindowsProc = 1
    End If
End Function

Public Function WindowTitle(ByVal lHwnd As Long) As String
Dim lLen As Long
Dim sBuf As String

    ' Get the Window Title:
    lLen = GetWindowTextLength(lHwnd)
    If (lLen > 0) Then
        sBuf = String$(lLen + 1, 0)
        lLen = GetWindowText(lHwnd, sBuf, lLen + 1)
        WindowTitle = Left$(sBuf, lLen)
    End If
    
End Function

Public Function ClassName(ByVal lHwnd As Long) As String
Dim lLen As Long
Dim sBuf As String
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lHwnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If
End Function


