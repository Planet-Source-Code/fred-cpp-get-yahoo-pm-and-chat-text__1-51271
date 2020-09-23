VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yahoo!Text - Get Yahoo Messages and Chat Text"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox listhWnd 
      Height          =   2595
      Left            =   -1200
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin YahooText.ISButton cmdUpdate 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   4
      Icon            =   "frmMain.frx":0E42
      Caption         =   "Update List"
      CaptionAlign    =   4
      GColor1         =   -2147483633
      GColor2         =   16761024
   End
   Begin SHDocVwCtl.WebBrowser WebChat 
      Height          =   2895
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   3495
      ExtentX         =   6165
      ExtentY         =   5106
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin YahooText.CoolList clWindows 
      Height          =   2880
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   5080
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSelected    =   8388608
      BackNormal      =   -2147483633
      BackSelected    =   16764057
      BackSelectedG2  =   16761024
      BoxBorder       =   -2147483642
      BoxRadius       =   8
      Focus           =   0   'False
      ItemHeight      =   32
      ItemHeightAuto  =   0   'False
      SelectModeStyle =   4
   End
   Begin YahooText.ISButton cmdAbout 
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   4
      Icon            =   "frmMain.frx":0F9C
      Caption         =   "About"
      CaptionAlign    =   4
      GColor1         =   -2147483633
      GColor2         =   16761024
   End
   Begin YahooText.ISButton cmdSaveHTML 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   4
      Icon            =   "frmMain.frx":1DEE
      Caption         =   "Save"
      CaptionAlign    =   4
      GColor1         =   -2147483633
      GColor2         =   16761024
   End
   Begin YahooText.ISButton cmdSave 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   4
      Icon            =   "frmMain.frx":1F48
      Caption         =   "Save"
      CaptionAlign    =   4
      GColor1         =   -2147483633
      GColor2         =   16761024
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This project Is Based on previous submisions in PSC
' Original Author: c0ldfyr3 www.EliteProdigy.com
' Edited and Improved by: Fred.cpp fred_cpp@msn.com
'
' If You Find Bugs or want to suggest Improvements,
' do It on the page (Link In the About Screen)
'
' Fred.cpp

'A demonstration on how to get the Yahoo Chat or PM HTML or Text
'Based on Get YM! Chat Text & HTML from "Internet Explorer_Server" object By: James Johnston
'I cleaned up his method for ease of use, and also reduced the amount of code by 1/10th
'I added the last line and got rid of the extra scroll bar in the HTML version
'This made it alot smoother and faster to execute.
'Now, I can go back to makin chat bots >:)
'c0ldfyr3 www.EliteProdigy.com

Option Explicit
Implements IEnumWindowsSink


Private Function ParseHTML(HTML As String) As String
    Dim Pos                         As Long
    Dim bChat                       As Boolean
    If Len(HTML) = 0 Then Exit Function
    Pos = InStrRev(HTML, "<DIV id=$im ")
    If Pos = 0 Then
        Pos = InStrRev(HTML, "<DIV id=imbody")
        bChat = True
    End If
    HTML = Mid(HTML, Pos)
    Pos = InStrRev(HTML, "<SCRIPT")
    HTML = Mid(HTML, 1, Pos - 1)
    If bChat = True Then
        HTML = Replace(HTML, "<DIV id=imbody style=""OVERFLOW-Y: scroll; BACKGROUND: none transparent scroll repeat 0% 0%; LEFT: 0px; OVERFLOW-X: auto; OVERFLOW: auto; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%; WORD-WRAP: break-word"" onscroll=$HandleScroll()>", "")
    Else
        HTML = Replace(HTML, "<DIV id=$im style=""OVERFLOW-Y: scroll; LEFT: 0px; OVERFLOW-X: auto; OVERFLOW: auto; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%; WORD-WRAP: break-word""><SPAN id=$rh style=""DISPLAY: none""></SPAN>", "")
    End If
    HTML = Replace(HTML, "<IFRAME style=""DISPLAY: none"" name=ymsgr></IFRAME>", "")
    If bChat = True Then
        HTML = c_ChatStyle & HTML
                    
    Else
        HTML = c_PMStyle & HTML
    End If
    ParseHTML = HTML
End Function

Private Sub cmdHTMLSave_Click()
    WebChat.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_SHOWHELP
End Sub

Private Sub cmdSave_Click()
    WebChat.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
    WebChat.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
    'docChat. = Clipboard.GetData
    On Error GoTo NoWordObjectCreated
    Dim wObj As New Word.Application
    Dim wDoc As New Word.Document
    On Error GoTo NoActiveDocument
    'Clipboard.GetData
    wObj.NewDocument.Add "doc1.doc"
    wDoc.Application.selection.PasteAndFormat (wdPasteDefault)
    wObj.ActiveDocument.Save
    wDoc.Close False
    wObj.Quit
    Set wDoc = Nothing
    Set wObj = Nothing
Exit Sub
NoWordObjectCreated:
    On Error Resume Next
    MsgBox "Unable to start Microsoft Word", vbCritical
    Set wDoc = Nothing
    Set wObj = Nothing
Exit Sub
NoActiveDocument:
On Error Resume Next
    MsgBox "Unable to Save", vbCritical
    Set wDoc = Nothing
    Set wObj = Nothing
End Sub


Private Sub clWindows_DblClick()
    'Get Especified Windows Text
    Dim YIM                         As YIMType
    Dim sHTML                       As String
    'Add the plain text
    WebChat.GoHome ' .Navigate "about:Blank"
    YIM = GetIMWindowText(Me.listhWnd.List(clWindows.ListIndex))
    'Add the HTML part, replacing the HTML that adds an extra scrollbar in out web applet.
    sHTML = ParseHTML(YIM.HTML)
    Open App.Path & "\tmp.htm" For Output As #1
    Print #1, sHTML
    Close #1
    WebChat.Navigate2 App.Path & "\tmp.htm"
    'WebChat.Document.write sHTML
End Sub

Private Sub cmdAbout_Click()
    AddAbout
End Sub

Private Sub cmdSaveHTML_Click()
    WebChat.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub IEnumWindowsSink_EnumWindow(ByVal hWnd As Long, bStop As Boolean)
    'Search Y!Messenger Windows
    If Len(Trim(WindowTitle(hWnd))) > 0 And IsWindowVisible(hWnd) Then
        If (StrComp(ClassName(hWnd), "IMClass") <> 0) Then Exit Sub
        clWindows.AddItem WindowTitle(hWnd)
        listhWnd.AddItem hWnd
    End If
End Sub

Private Property Get IEnumWindowsSink_Identifier() As Long
    IEnumWindowsSink_Identifier = Me.hWnd
End Property

Private Sub cmdUpdate_Click()
    clWindows.Clear
    EnumerateWindows Me
End Sub

Private Sub Form_Load()
    EnumerateWindows Me
    AddAbout
End Sub

Sub AddAbout()
    Dim tmpStr As String
    tmpStr = "<html><base target='_blank'><body text='#000080' bgcolor='#99CCFF'><p align='center'><font face='Verdana' color='#000000'><b>Yahoo!Text<br></b></font><font face='Verdana' size='1'><a href='http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=51271&lngWId=1'>Get Yahoo Messages and Chat Text</a><br>By:<br></font><font face='Verdana' size='2'><b><a href='mailto:fred_cpp@msn.com?subject=About Yahoo!Text'>Fred.cpp</a></b></font></p><p align='center'><font face='Verdana' size='1'>Based On: <br>Lots Of Submitions from: <br><a href='http://www.planet-source-code.com'>www.planet-source-code.com&nbsp;</a></font></p><p align='center'><font face='Verdana' size='1'>Blue Eyes<br>(<a href='http://www.hope-tech.net'>http://www.hope-tech.net</a>)<br><a href='http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=41053&lngWId=1'>(Windows Enumerate)</a></font></p>" _
    & "<p align='center'><font face='Verdana' size='1'>CharlesPV <br>(<a href='http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=29586&lngWId=1'>ucCoolList</a>)</font></p><p align='center'><font face='Verdana' size='1'>c0ldfyr3<br>(<a href='http://www.EliteProdigy.com'>http://www.EliteProdigy.com</a>)<br><a href='http://www.planet-source-code.com/vb/scripts/showcode.asp?txtCodeId=50201&lngWId=1'>(Get IMText)</a></font></p><p align='center'>&nbsp;</p></body></html>"
    Open App.Path & "\tmp.htm" For Output As #1
    Print #1, tmpStr
    Close #1
    WebChat.Navigate2 App.Path & "\tmp.htm"
End Sub

