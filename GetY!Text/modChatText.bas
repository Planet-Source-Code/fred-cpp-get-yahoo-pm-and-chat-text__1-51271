Attribute VB_Name = "modChatText"
Option Explicit
'A demonstration on how to get the Yahoo Chat or PM HTML or Text
'Based on Get YM! Chat Text & HTML from "Internet Explorer_Server" object By: James Johnston
'I cleaned up his method for ease of use, and also reduced the amount of code by 1/10th
'I added the last line and got rid of the extra scroll bar in the HTML version
'Now, I can go back to makin chat bots >:)
'c0ldfyr3 www.EliteProdigy.com

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function ObjectFromLresult Lib "oleacc" (ByVal lResult As Long, riid As UUID, ByVal wParam As Long, ppvObject As Any) As Long
Private Const SMTO_ABORTIFHUNG      As Long = &H2
Private Const GW_HWNDFIRST          As Long = 0
Private Const GW_HWNDNEXT           As Long = 2
Private Const GW_CHILD              As Long = 5
Private Type UUID
   Data1                            As Long
   Data2                            As Integer
   Data3                            As Integer
   Data4(0 To 7)                    As Byte
End Type
Private Type typWindows
    ClassName                       As String
    hWnd                            As Long
End Type
Private Type typWinFinal
    ChildWindows()                  As typWindows
    Count                           As Integer
End Type
Public Type YIMType
    Text                            As String
    HTML                            As String
End Type
Public Const c_ChatStyle            As String = "<STYLE>" & vbCrLf & _
                                                ".sendername { font-size:10pt;font-family:Arial;font-weight:bold;color:#000000;text-decoration:none };" & vbCrLf & _
                                                ".recvername { font-size:10pt;font-family:Arial;font-weight:bold;color:#0000FF;text-decoration:none };" & vbCrLf & _
                                                ".ymsgrname { font-size:10pt;font-family:Arial;font-weight:bold;color:#FF0000;text-decoration:none };" & vbCrLf & _
                                                ".chatusername { font-size:10pt;font-family:Arial;color:#FF0000;text-decoration:none };" & vbCrLf & _
                                                ".usertext { font-size:10pt;font-family:Arial; };" & vbCrLf & _
                                                ".redstatus { font-size:10pt;font-family:Arial;font-weight:bold;color:#FF0000;text-decoration:none };" & vbCrLf & _
                                                ".greenstatus { font-size:10pt;font-family:Arial;font-weight:bold;color:#008800;text-decoration:none };" & vbCrLf & _
                                                ".graystatus { font-size:10pt;font-family:Arial;font-weight:bold;color:#888888;text-decoration:none };" & vbCrLf & _
                                                ".chatrecver { font-size:10pt;font-family:Arial;font-weight:bold;color:#880000;text-decoration:none };" & vbCrLf & _
                                                ".chatsender { font-size:10pt;font-family:Arial;font-weight:bold;color:#0000FF;text-decoration:none };" & vbCrLf & _
                                                ".chataction { font-size:10pt;font-family;Arial;color:#880088;text-decoration:none };" & vbCrLf & _
                                                "a { color:#0000FF; };" & vbCrLf & _
                                                "p { text-indent:-7;margin-left:10;margin-top:0;margin-bottom:0 };" & vbCrLf & _
                                                "</STYLE>"
Public Const c_PMStyle              As String = "<STYLE>" & vbCrLf & _
                                                ".sendername { font-size:10pt;font-family:Arial;font-weight:bold;color:#000000; }" & vbCrLf & _
                                                ".recvername { font-size:10pt;font-family:Arial;font-weight:bold;color:#0000FF; }" & vbCrLf & _
                                                ".ymsgrname { font-size:10pt;font-family:Arial;font-weight:bold;color:#FF0000; }" & vbCrLf & _
                                                ".usertext { font-size:10pt;font-family:Arial; }" & vbCrLf & _
                                                ".redstatus { font-size:10pt;font-family:Arial;font-weight:bold;color:#FF0000; }" & vbCrLf & _
                                                ".greenstatus { font-size:10pt;font-family:Arial;font-weight:bold;color:#008800; }" & vbCrLf & _
                                                ".graystatus { font-size:10pt;font-family:Arial;font-weight:bold;color:#888888; }" & vbCrLf & _
                                                ".imvnotify { font-size:10pt;font-family:Arial;font-weight:bold;color:#000088; }" & vbCrLf & _
                                                "a { color:#0000FF; }" & vbCrLf & _
                                                "p { text-indent:-7;margin-left:10;margin-top:0;margin-bottom:0 }" & vbCrLf & _
                                                "</STYLE>"

Public Function GetClassN(ByVal hWnd As Long) As String
    Dim ParentClassName             As String
    Dim Z                           As Long
    ParentClassName = String(100, Chr(0))
    Z = GetClassName(hWnd, ParentClassName, 100)
    GetClassN = Left(ParentClassName, Z)
End Function
Private Function GetChildWindows(hWnd As Long) As typWinFinal
    Dim ChildP                      As Long
    Dim LastChild                   As String
    Dim MainP                       As Long
    Dim WinDetails                  As String
    Dim First                       As Boolean
    Dim AdWin                       As Long
    
    GetChildWindows.Count = -1
    MainP = GetWindow(hWnd, GW_CHILD)
    ChildP = GetWindow(MainP, GW_HWNDFIRST)
    Do While ChildP <> 0
        ChildP = GetWindow(ChildP, GW_HWNDNEXT)
        If ChildP = 0 Then Exit Do
        WinDetails = GetClassN(ChildP)
        
        GetChildWindows.Count = GetChildWindows.Count + 1
        ReDim Preserve GetChildWindows.ChildWindows(GetChildWindows.Count)
        
        With GetChildWindows.ChildWindows(GetChildWindows.Count)
            .ClassName = WinDetails
            .hWnd = ChildP
        End With
        DoEvents
    Loop
End Function
Public Function GetIMText() As YIMType
    Dim IMClass                     As Long
    Dim MidWin                      As Long
    Dim InternetExplorerServer      As Long
    Dim Something                   As typWinFinal
    Dim X                           As Integer
    Dim sTmp                        As String
    Dim yTmp                        As YIMType
    
    'Loop through all the windows finding their handles from predefined classnames.
    IMClass = FindWindow("imclass", vbNullString)
    Something = GetChildWindows(IMClass)
    For X = 0 To Something.Count
        sTmp = Something.ChildWindows(X).ClassName
        If Len(sTmp) > 4 Then
            If StrComp(Left(sTmp, 3), "atl", vbTextCompare) = 0 Then
                InternetExplorerServer = FindWindowEx(Something.ChildWindows(X).hWnd, 0&, "internet explorer_server", vbNullString)
                yTmp = GetIEText(InternetExplorerServer)
                If InStr(1, yTmp.HTML, "function RestoreStyles()") > 0 Then
                    GetIMText = yTmp
                End If
            End If
        End If
    Next
End Function
Private Function GetIEText(ByVal hWnd As Long) As YIMType
    Dim doc                         As IHTMLDocument2
    Dim col                         As IHTMLElementCollection2
    Dim EL                          As IHTMLElement
    Dim l                           As Long
    Dim v1                          As Variant
    Dim v2                          As Variant
    Set doc = IEDOMFromhWnd(hWnd)
    'Pass the data back through the function
    On Error GoTo Ender:
    GetIEText.Text = doc.body.innerText
    GetIEText.HTML = doc.body.innerHTML
    Exit Function
Ender:
    GetIEText.Text = "No Chat Or Pm Open ?"
    Err.Clear
End Function
Private Function IEDOMFromhWnd(ByVal hWnd As Long) As IHTMLDocument
    Dim IID_IHTMLDocument           As UUID
    Dim hWndChild                   As Long
    Dim spDoc                       As IUnknown
    Dim lRes                        As Long
    Dim lMsg                        As Long
    Dim hr                          As Long
    If hWnd <> 0 Then
        'If the Handle is not 0, that means if the window is open .........
        'We tell windows we are going in for the kill, and grabbing the data
        lMsg = RegisterWindowMessage("WM_HTML_GETOBJECT")
        Call SendMessageTimeout(hWnd, lMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, lRes)
        If lRes Then
            With IID_IHTMLDocument
                .Data1 = &H626FC520
                .Data2 = &HA41E
                .Data3 = &H11CF
                .Data4(0) = &HA7
                .Data4(1) = &H31
                .Data4(2) = &H0
                .Data4(3) = &HA0
                .Data4(4) = &HC9
                .Data4(5) = &H8
                .Data4(6) = &H26
                .Data4(7) = &H37
            End With
            hr = ObjectFromLresult(lRes, IID_IHTMLDocument, 0, IEDOMFromhWnd)
            'We pass the data back from the function.
        End If
    End If
End Function


Public Function GetIMWindowText(lWindow As Long) As YIMType
    Dim IMClass                     As Long
    Dim MidWin                      As Long
    Dim InternetExplorerServer      As Long
    Dim Something                   As typWinFinal
    Dim X                           As Integer
    Dim sTmp                        As String
    Dim yTmp                        As YIMType
    
    'Loop through all the windows finding their handles from predefined classnames.
    'IMClass = FindWindow("imclass", vbNullString)
    Something = GetChildWindows(lWindow)
    For X = 0 To Something.Count
        sTmp = Something.ChildWindows(X).ClassName
        If Len(sTmp) > 4 Then
            If StrComp(Left(sTmp, 3), "atl", vbTextCompare) = 0 Then
                InternetExplorerServer = FindWindowEx(Something.ChildWindows(X).hWnd, 0&, "internet explorer_server", vbNullString)
                yTmp = GetIEText(InternetExplorerServer)
                If InStr(1, yTmp.HTML, "function RestoreStyles()") > 0 Then
                    GetIMWindowText = yTmp
                End If
            End If
        End If
    Next
End Function



