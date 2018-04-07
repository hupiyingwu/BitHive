VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "HTTPserver"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   5550
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5220
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3069
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3069
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3069
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   3240
   End
   Begin MSWinsockLib.Winsock sckWS 
      Index           =   0
      Left            =   4320
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "end"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4920
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Download by http://www.NewXing.com
'用户请求页面名
Private requestedPage As String

Private Sub Command1_Click()
load_defaults '开始服务
End Sub

Private Sub Command2_Click()
Call stop_server
End Sub

Private Sub Form_Load()
ttlConnections = 0
End Sub

Private Sub sckWS_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Index = 0 Then
      '总用户数加一
      ttlConnections = ttlConnections + 1
      '当前用户数加一
      numConnections = numConnections + 1
      '用户满
      If numConnections = maxConnections Then GoTo done
      '为这个用户新装载一个WINSOCK控件实例
      Load sckWS(ttlConnections)
      sckWS(ttlConnections).LocalPort = 0
      sckWS(ttlConnections).Accept requestID
      '加入日志
      List1.AddItem sckWS(ttlConnections).RemoteHostIP & " connected"
      
StartOver:
      
      DoEvents
      '如果还不知道用户请求哪个页面，则循环
      If requestedPage$ = "" Then GoTo StartOver
      List1.AddItem "requested page: " & requestedPage$
      '如果为"/"，则将请求页面设为默认主页
      If requestedPage$ = "/" Then requestedPage$ = htmlIndexPage$
      '判断请求页面文件是否存在
      If FileExists(htmlPageDir & "\" & requestedPage$) Then
      '存在时的处理
      htmldata$ = text_read(htmlPageDir & "\" & requestedPage$)
      htmldata$ = ReplaceStr(htmldata$, "$ip", sckWS(0).LocalIP)
      sckWS(ttlConnections).SendData htmldata$ & vbCrLf
      Else 'if it doesn't exist, then...
       '文件不存在时的处理
       If requestedPage$ = htmlIndexPage$ Then
       sckWS(ttlConnections).SendData "<html><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1""><b>Please create an index html page.  It was not found.</font></html>" & vbCrLf
       requestedPage$ = ""
       End If
    
      requestedPage$ = "/a"
      sckWS(ttlConnections).SendData html_404$ & vbCrLf '发送出错文件
      End If
   End If
   
done:
      numConnections = numConnections - 1
End Sub

Private Sub sckWS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strdata As String
sckWS(Index).GetData strdata$
If Mid$(strdata$, 1, 3) = "GET" Then
findget = InStr(strdata$, "GET ")
spc2 = InStr(findget + 5, strdata$, " ")
pagetoget$ = Mid$(strdata$, findget + 4, spc2 - (findget + 4))
requestedPage$ = pagetoget$
End If
End Sub

Private Sub sckWS_SendComplete(Index As Integer)
If requestedPage$ <> "" Then
      requestedPage$ = ""
      sckWS(ttlConnections).Close
End If
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(1).Text = " current status:" & sckWS(0).State
StatusBar1.Panels(2).Text = "online nodes:" & numConnections
StatusBar1.Panels(3).Text = "logs" & ttlConnections
End Sub
