VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "webview2内核的网页浏览器"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13515
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   565
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   901
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   11160
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "H5特效测试"
      Height          =   330
      Left            =   9720
      TabIndex        =   7
      Top             =   90
      Width           =   1215
   End
   Begin VB.CommandButton cmdCaptureWV 
      Caption         =   "网页截图"
      Height          =   330
      Left            =   5535
      TabIndex        =   5
      Top             =   90
      Width           =   1815
   End
   Begin VB.CommandButton cmdNavigate 
      Caption         =   "打开百度贴吧"
      Height          =   330
      Left            =   3330
      TabIndex        =   4
      Top             =   90
      Width           =   2130
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2070
      TabIndex        =   2
      Text            =   "666"
      Top             =   90
      Width           =   735
   End
   Begin VB.CommandButton cmdAssignNewText 
      Caption         =   "给输入框赋值"
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   1935
   End
   Begin VB.PictureBox picWV 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   8055
      TabIndex        =   3
      Top             =   540
      Width           =   8055
   End
   Begin VB.CommandButton cmdOpenDevTools 
      Caption         =   "打开调试工具DevTools"
      Height          =   330
      Left            =   7425
      TabIndex        =   0
      Top             =   90
      Width           =   2160
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   45
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Note, that this Demo requires the properly registered RC6-Binaries
'and in addition an installed "Chromium-Edge" (in its "evergreen" WebView2-incarnation)
'installable from its official MS-Download-URL: https://go.microsoft.com/fwlink/p/?LinkId=2124703
 
Private WithEvents WV As cWebView2 'declare a WebView-variable WithEvents
Attribute WV.VB_VarHelpID = -1

Private Sub Command1_Click()
    WV.Navigate "http://www.htmlsucai.com/demo-56200.html"
End Sub

Private Sub Form_Load()
    Visible = True '<- it's important, that the hosting TopLevel-Form is visible...
                '...(and thus the Child-PicBox indirectly as well) - before we Bind the PicBox to the WebView
    
    Set WV = New_c.WebView2 'create the instance
    If WV.BindTo(picWV.hWnd) = 0 Then MsgBox "couldn't initialize WebView-Binding": Exit Sub
    
    '  Set WV = New_c.WebView2(picWV.hWnd) 'create the instance
    '  If WV Is Nothing Then MsgBox "couldn't initialize WebView-Binding": Exit Sub
    LocalWebViewInit 'initialize the WebView for local usage here in our Form
End Sub

Private Sub LocalWebViewInit()
    'we can predefine our own set of js-functions, before any document gets loaded
    WV.AddScriptToExecuteOnDocumentCreated "function test(a,b){ return a+b }"
    WV.AddScriptToExecuteOnDocumentCreated "function btn1_click_test(){ vbH().RaiseMessageEvent('btn1_click','') }"
    
    'so, above we've added two small functions (the latter btn1_click() function being used inside the following HTML-init-string)
    WV.NavigateToString "<!DOCTYPE html><html><head><title>AppTitle</title></head><body>" & _
                          "<div>Hello World...</div>" & _
                          "<input id='txt1' value='foo'>" & _
                          "<button id='btn1' onclick='btn1_click_test()' >Button1</button>" & _
                      "</body></html>"
    
    'this follow-up line shows already an interaction with the just loaded Document
    Dim btn1Caption As String 'reading the current caption-text out of the 'btn1' element
    btn1Caption = WV.jsProp("document.getElementById('btn1').innerHTML")
    Debug.Print btn1Caption, WV.DocumentTitle, WV.DocumentURL
    
    'and this shows, that the WV.jsProp("...") also works in Property-Let-Mode (at the left-hand-side)
    btn1Caption = "Click Me..." 'change the Caption-String
    WV.jsProp("document.getElementById('btn1').innerHTML") = btn1Caption 'and assign it to the Browser-Element as the new Caption via WV.jsProp() = ...
    'just for fun, we can change the style of the btn1-Element to color='red' as well this way
    WV.jsProp("document.getElementById('btn1').style.color") = "red"
    
    'and here we make first use, of our (at the very top) predefined js-test() function
    WV.jsRunAsync "test", 2, 3 'run the above added javascript test()-function asynchronously
    Debug.Print "async jsRun-started" 'so this PrintOut should come immediately after the call above  (and before the WV_JSAsyncResult-Event delivers the result)
End Sub

'*** VB-Command-Button-Handlers
Private Sub cmdAssignNewText_Click()
    WV.jsProp("document.getElementById('txt1').value") = Text1.Text 'assign a VB-Value to a WV-text-field
End Sub

Private Sub cmdNavigate_Click()
    WV.Navigate "https://tieba.baidu.com/index.html" '<- alternatively WV.jsProp("location.href") = "https://google.com" would also work
    
    'the call below, just to show that our initially added js-functions, remain "in place" - even when we re-navigate to something else
    WV.jsRunAsync "test", 2, 3
End Sub

Private Sub cmdCaptureWV_Click()
    Dim Srf As cCairoSurface
    Set Srf = WV.CapturePreview(CaptureAs_PNG) 'capture the current WV-Window as a Cairo-Image-Surface
    Srf.WriteContentToPngFile App.Path & "\WV_Capture.png" 'which we can now visualize, or just write out as a PNG-file
End Sub

Private Sub cmdOpenDevTools_Click()
    WV.OpenDevToolsWindow
End Sub

Private Sub Form_Resize()
    On Error Resume Next 'adjust the hosting VB-PicBox, according to the Form-size
    picWV.Move 0, picWV.Top, ScaleWidth, ScaleHeight - picWV.Top - lblStatus.Height - 10
    lblStatus.Top = picWV.Top + picWV.Height + 8
End Sub

Private Sub picWV_Resize() 'when the hosting picBox got resized, we have to call a syncSize-method on the WebView
    If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub
Private Sub picWV_GotFocus() 'same thing here... when the hosting picBox got the focus, we tell the WebView about it
    If Not WV Is Nothing Then WV.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then WV.Navigate Text2.Text
End Sub

'*** the above two EventHandlers (of the hosting VB-PicBox-Container-Ctl) are basically all what's needed "GUI-Binding-wise"

'*** the rest of the EventHandlers below, are raised by the WebView-instance itself

Private Sub WV_AcceleratorKeyPressed(ByVal KeyState As RC6.eWebView2AccKeyState, ByVal IsExtendedKey As Boolean, ByVal WasKeyDown As Boolean, ByVal IsKeyReleased As Boolean, ByVal IsMenuKeyDown As Boolean, ByVal RepeatCount As Long, ByVal ScanCode As Long, IsHandled As Boolean)
    Debug.Print "WV_AcceleratorKeyPressed"
End Sub

Private Sub WV_InitComplete()
    Debug.Print "WV_InitComplete"
End Sub

Private Sub WV_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
    Debug.Print "WV_NavigationCompleted"
End Sub

Private Sub WV_DocumentComplete()
    Debug.Print "WV_DocumentComplete"
    lblStatus.Caption = WV.DocumentURL
End Sub

Private Sub WV_GotFocus(ByVal Reason As eWebView2FocusReason)
    Debug.Print "WV_GotFocus", Reason
End Sub

Private Sub WV_JSAsyncResult(Result As Variant, ByVal Token As Currency, ByVal ErrString As String)
    Debug.Print "WV_JSAsyncResult "; Result, Token, ErrString
End Sub

Private Sub WV_JSMessage(ByVal sMsg As String, ByVal sMsgContent As String, oJSONContent As cCollection)
    Debug.Print sMsg, sMsgContent
    Select Case sMsg
        Case "btn1_click": MsgBox "txt1.value: " & WV.jsProp("document.getElementById('txt1').value")
    End Select
End Sub

Private Sub WV_LostFocus(ByVal Reason As eWebView2FocusReason)
    Debug.Print "WV_LostFocus", Reason
End Sub

Private Sub WV_NewWindowRequested(ByVal IsUserInitiated As Boolean, IsHandled As Boolean, ByVal URI As String, NewWindowFeatures As RC6.cCollection)
    'IsUserInitiated = False
    IsHandled = True
    WV.Navigate URI, 0
End Sub

Private Sub WV_UserContextMenu(ByVal ScreenX As Long, ByVal SreenY As Long)
    Debug.Print "WV_UserContextMenu", ScreenX, SreenY
End Sub

