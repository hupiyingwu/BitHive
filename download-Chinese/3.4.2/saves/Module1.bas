Attribute VB_Name = "Module1"
'Download by http://www.NewXing.com
'开始变量定义
'定义端口，HTTP的缺省端口是80
Global http_port As Long
'已有的连接用户总数
Global ttlConnections As Long
'最大支持的连接用户数
Global maxConnections As Long
'当前连接的用户数
Global numConnections As Long
'HTML文件所在路径
Global htmlPageDir As String
'没有所请求页面时的网页
Global html_404 As String
'默认页面文件名
Global htmlIndexPage As String
Sub load_defaults()
'开始服务
Dim tport As String
'变量赋初值
http_port = 80
maxConnections = 100
ttlConnections = 0
numConnections = 0
htmlPageDir$ = App.Path
htmlIndexPage$ = "index.html"
html_404$ = html_404error()
'变量初值结束
tport$ = ""
If http_port = 80 Then tport$ = "" Else: tport$ = ":" & http_port
With frmMain
    .sckWS(0).Close
    .sckWS(0).LocalPort = http_port
    .sckWS(0).Listen
    '.Text1.Text = "http://" & .sckWS(0).LocalIP & tport$
    .Command1.Enabled = False
    .Command2.Enabled = True
End With

End Sub
Public Sub stop_server()
'关闭服务器函数
With frmMain
.Command1.Enabled = True
.Command2.Enabled = False
.List1.Clear
'.Text1.Text = ""
.sckWS(0).Close
End With

Call unloadControls
Call unset_vars

End Sub
Public Sub unloadControls()
'卸载我们所转载的所有WINSOCK控件
With frmMain
For i = 1 To ttlConnections
Unload .sckWS(i)
Next i
End With
End Sub
Public Sub unset_vars()
'所有的变量重置
http_port = 0
ttlConnections = 0
maxConnections = 0
numConnections = 0
htmlPageDir = 0
html_404 = ""
htmlIndexPage = ""
End Sub
