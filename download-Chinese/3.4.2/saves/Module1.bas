Attribute VB_Name = "Module1"
'Download by http://www.NewXing.com
'��ʼ��������
'����˿ڣ�HTTP��ȱʡ�˿���80
Global http_port As Long
'���е������û�����
Global ttlConnections As Long
'���֧�ֵ������û���
Global maxConnections As Long
'��ǰ���ӵ��û���
Global numConnections As Long
'HTML�ļ�����·��
Global htmlPageDir As String
'û��������ҳ��ʱ����ҳ
Global html_404 As String
'Ĭ��ҳ���ļ���
Global htmlIndexPage As String
Sub load_defaults()
'��ʼ����
Dim tport As String
'��������ֵ
http_port = 80
maxConnections = 100
ttlConnections = 0
numConnections = 0
htmlPageDir$ = App.Path
htmlIndexPage$ = "index.html"
html_404$ = html_404error()
'������ֵ����
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
'�رշ���������
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
'ж��������ת�ص�����WINSOCK�ؼ�
With frmMain
For i = 1 To ttlConnections
Unload .sckWS(i)
Next i
End With
End Sub
Public Sub unset_vars()
'���еı�������
http_port = 0
ttlConnections = 0
maxConnections = 0
numConnections = 0
htmlPageDir = 0
html_404 = ""
htmlIndexPage = ""
End Sub
