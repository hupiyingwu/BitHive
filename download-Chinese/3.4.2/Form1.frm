VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "BitHive"
   ClientHeight    =   6270
   ClientLeft      =   9045
   ClientTop       =   4605
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8010
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6840
      Top             =   2280
   End
   Begin VB.Frame Frame3 
      Caption         =   "指令"
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   5775
      Begin VB.CommandButton Command5 
         Caption         =   "加入"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "添加加密文本"
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox incommand 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "已知节点"
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "加入"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
      Begin VB.ListBox clientlist 
         Height          =   960
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "用户数据"
      Height          =   1695
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   5775
      Begin VB.ListBox list 
         Height          =   1320
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5055
      End
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   360
      Width           =   495
      ExtentX         =   873
      ExtentY         =   873
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label back 
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label doing 
      Caption         =   "ready"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Menu file 
      Caption         =   "文件"
      Begin VB.Menu file1 
         Caption         =   "创建密钥"
      End
      Begin VB.Menu file2 
         Caption         =   "加载数据"
      End
      Begin VB.Menu file3 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu wallet 
      Caption         =   "钱包"
      Begin VB.Menu wallet1 
         Caption         =   "发送代币"
      End
      Begin VB.Menu wallet2 
         Caption         =   "接受代币"
      End
      Begin VB.Menu wallet3 
         Caption         =   "挖矿"
      End
   End
   Begin VB.Menu sa 
      Caption         =   "脚本"
      Begin VB.Menu sa1 
         Caption         =   "创建"
         Begin VB.Menu sa2 
            Caption         =   "脚本"
         End
         Begin VB.Menu sa3 
            Caption         =   "应用程序"
         End
         Begin VB.Menu sa4 
            Caption         =   "文件"
         End
      End
      Begin VB.Menu sa5 
         Caption         =   "结束当前任务"
      End
   End
   Begin VB.Menu network 
      Caption         =   "网络"
      WindowList      =   -1  'True
      Begin VB.Menu network1 
         Caption         =   "加入网络"
      End
      Begin VB.Menu network2 
         Caption         =   "添加节点"
      End
   End
   Begin VB.Menu ud 
      Caption         =   "用户数据"
      Begin VB.Menu ud1 
         Caption         =   "区块链"
      End
      Begin VB.Menu ud2 
         Caption         =   "社区"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim scripthash As String, scriptname As String, mypeaple As Integer
Dim dignum As Integer

Const a As Byte = 20
Const b As Byte = 40
Dim myblock As String
Dim tasked As String, username As String, prikey As String, userdata() As String, usernum As Integer, checklist() As String, cln As Integer, clients As String, errrownum As Integer, timepoint As String, tmp() As String
Private Function bstr(code As String, str1 As String, str2 As String) As String

    bstr = Split(Split(code, clear(str1, " "))(1), clear(str2, " "))(0)
End Function
Private Function creatdata(basename As String, data As String) As String
    creatdata = "[" + basename + "]" + data + "[/" + basename + "]"
End Function
Private Function readata(base As String, basename As String) As String
    readata = bstr(base, "[" + basename + "]", "[/" + basename + "]")
End Function
Private Function StrJiaMi(ByVal strSource As String, ByVal Key1 As Byte, _
ByVal Key2 As Integer) As String
Dim bLowData As Byte
Dim bHigData As Byte
Dim i As Integer
Dim strEncrypt As String
Dim strChar As String
For i = 1 To Len(strSource)

strChar = Mid(strSource, i, 1)

bLowData = AscB(MidB(strChar, 1, 1)) Xor Key1

bHigData = AscB(MidB(strChar, 2, 1)) Xor Key2

If Len(Hex(bLowData)) = 1 Then
strEncrypt = strEncrypt & "0" & Hex(bLowData)
Else
strEncrypt = strEncrypt & Hex(bLowData)
End If
If Len(Hex(bHigData)) = 1 Then
strEncrypt = strEncrypt & "0" & Hex(bHigData)
Else
strEncrypt = strEncrypt & Hex(bHigData)
End If
Next
StrJiaMi = strEncrypt
End Function

Private Function StrJiMi(ByVal strSource As String, ByVal Key1 As Byte, _
ByVal Key2 As Integer) As String
Dim bLowData As Byte
Dim bHigData As Byte
Dim i As Integer
Dim strEncrypt As String
Dim strChar As String
For i = 1 To Len(strSource) Step 4

strChar = Mid(strSource, i, 4)

bLowData = "&H" & Mid(strChar, 1, 2)
bLowData = bLowData Xor Key1

bHigData = "&H" & Mid(strChar, 3, 2)
bHigData = bHigData Xor Key2

strEncrypt = strEncrypt & ChrB(bLowData) & ChrB(bHigData)
Next
StrJiMi = strEncrypt
End Function

Private Function aut(code As String) As String
    Dim h As String
    h = hexing(Hash(code))
    Dim mychar() As String
    ReDim mychar(30) As String
    Dim i As Integer
    For i = 1 To Len(h)
        mychar(i - 1) = Mid(h, i, 1)
    Next
    If Len(h) < 31 Then
        For i = Len(h) To 31
            mychar(i - 1) = 2
        Next
    End If
    For i = 1 To 30
        Dim num As Integer
        num = Val(mychar(i))
        If num < 5 Then mychar(i) = 2
        If num >= 5 Then mychar(i) = 3
        If num = 0 Then mychar(i) = 4
        Dim password As String, p As Integer, q As Integer, key As Integer
        key = Val(readata(username, "key" & i))
        p = Val(readata(username, "p" & i))
        q = Val(readata(username, "q" & i))
        password = password + creatdata("h" & i, rsajia(p, q, key, Val(mychar(i))))
    Next
    aut = password
End Function
Private Function checkaut(user As String, code As String, h As String) As Boolean
    Dim i As Integer, sh As String, mychar() As String
    sh = hexing(Hash(code))
    
    ReDim mychar(30) As String
    For i = 1 To Len(sh)
        mychar(i - 1) = Mid(sh, i, 1)
    Next
    If Len(sh) < 31 Then
        For i = Len(sh) To 31
            mychar(i - 1) = 2
        Next
    End If
    For i = 1 To 30
        Dim p As Integer, q As Integer, key As Integer, r As String
        p = Val(readata(user, "p" & i))
        q = Val(readata(user, "q" & i))
        key = Val(readata(user, "key" & i))
        r = rsajie(p, q, key, Val(readata(h, "h" & i)))
        Dim num As Integer
        num = Val(mychar(i))
        If num < 5 Then mychar(i) = 2
        If num >= 5 Then mychar(i) = 3
        If num = 0 Then mychar(i) = 4
        If r <> mychar(i) Then
        
            checkaut = False
            Exit Function
        End If
    Next
    checkaut = True
End Function
Private Function clear(base As String, code As String) As String
    clear = Replace(base, code, "")
End Function

Public Sub adduser(user As String, txt As String)
    Dim i As Integer, change As Boolean
    change = True
    For i = 1 To usernum
        If InStr(userdata(i), "<" + user + ">") Then
            change = False
            userdata(i) = "<" + user + ">" + txt
        End If
    Next
    If change = True And usernum <> 0 Then
        
        Dim rtmp() As String
        ReDim rtmp(1 To usernum) As String
        For i = 1 To usernum
            rtmp(i) = userdata(i)
        Next
        usernum = usernum + 1
        ReDim userdata(1 To usernum) As String
        For i = 1 To usernum - 1
            userdata(i) = rtmp(i)
        Next

        
        userdata(usernum) = "<" + user + ">" + txt
    ElseIf change = True And usernum = 0 Then
        ReDim userdata(1 To 1) As String
        usernum = usernum + 1
        userdata(1) = "<" + user + ">" + txt
    End If
End Sub
Private Function rsajia(p As Integer, q As Integer, key As Integer, x As Integer) As Integer
    Dim n As Integer
    n = p * q
    rsajia = x ^ key Mod n

End Function
Private Function rsajie(p As Integer, q As Integer, key As Integer, y As Integer) As Integer
    Dim n As Integer
    n = q * p
    
    
    rsajie = y ^ key Mod n
    
End Function
Private Function hexing(Sj As String) As String
    Dim strHexSj As String
    Dim i As Long
    Dim bytSj() As Byte
    
    bytSj = StrConv(Sj, vbFromUnicode)
    For i = 0 To UBound(bytSj)
        strHexSj = strHexSj & Right("0" & Hex(bytSj(i)), 2)
    Next
    hexing = strHexSj
End Function
Public Sub savefile(data As String, file As String)
    Open file For Output As #1
        Print #1, data
    Close #1
End Sub
Public Sub savedata()
    savefile Str(usernum), "usernum"
    Dim i As Integer
    For i = 1 To usernum
        savefile userdata(i), "userdata(" & i & ")"
        DoEvents
    Next
    savefile tasked, App.Path + "/saves/" + "tasked"
    savefile clients, App.Path + "/saves/" + "clients"
    savefile myblock, "myblock"
    savefile scripthash, "scripthash"
    savefile scriptname, "scriptname"
    savefile Str(mypeaple), "mypeaple"
    savefile username, "username"
    savefile prikey, "prikey"
End Sub
Public Sub getdata()
    Dim n As Integer
    n = Len(clients) - Len(clear(clients, ";"))
    Dim i As Integer
    For i = 1 To n
        doing.Caption = "Connecting clients"
        DoEvents
        '==============================================================
        Dim client As String
        client = Split(clients, ";")(i)
        Dim num As Integer
        On Error GoTo e
        num = Val(openURL(client + "/usernum"))
        Dim j As Integer
        For j = 1 To num
            doing.Caption = "Downloading data"
            DoEvents
            '==========================================================
            Dim paylist As String, user As String, data As String
            On Error GoTo e
            paylist = openURL(client + "/userdata(" & j & ")")
            user = bstr(paylist, "<", ">")
            On Error GoTo e
            data = Split(paylist, "<" + user + ">")(1)
            If checkuser(user, data) <> "errow" Then
                Dim f As Integer, creatnew As Boolean
                creatnew = True
                For f = 1 To usernum
                    If InStr(userdata(f), "<" + user + ">") Then
                        creatnew = False
                        If Len(Split(userdata(f), "<" + user + ">")(1)) < Len(data) Then
                            adduser user, data
                        End If
                    End If
                Next
                If creatnew = True Then adduser user, data
            End If
        Next
        On Error GoTo e
        clients = clients + ";" + openURL(clients + "/clients")
e:
    Next
    username = readf("username")
    prikey = readf("prikey")
    myblock = readf("myblock")
    scripthash = readf("scripthash")
    scriptname = readf("scriptname")
    mypeaple = Val(readf("mypeaple"))
End Sub
Private Function readf(p As String) As String
    Dim nHandle, l As String
    nHandle = FreeFile
    On Error GoTo e
    Open p For Input As #nHandle
    
    Do Until EOF(nHandle)
        Line Input #nHandle, l
        readf = readf + l
    Loop
    Close nHandle
e:
End Function

Private Function hz(n1 As Integer, n2 As Integer) As Boolean
    Dim i As Integer
    For i = 2 To n1 - 1
        If n1 Mod i = 0 And n2 Mod i = 0 Then
            hz = False
            Exit Function
        End If
    Next
    For i = 2 To n2 - 1
        If n2 Mod i = 0 And n1 Mod i = 0 Then
            hz = False
            Exit Function
        End If
    Next
    hz = True
End Function
Public Sub getkey()
    prikey = ""
    username = ""
    Dim i As Integer, p(1 To 5) As Integer, q(1 To 5) As Integer
    p(1) = 2
    p(2) = 3
    p(3) = 5
    p(4) = 7
    p(5) = 11
    q(1) = 13
    q(2) = 17
    q(3) = 19
    q(4) = 23
    q(5) = 29
    For i = 1 To 30
e:
        Dim r1 As Integer, r2 As Integer
        Randomize
        r1 = Int(Rnd * 5) + 1
        Randomize
        r2 = Int(Rnd * 5) + 1
        Dim n As Integer, m As Integer
        n = p(r1) * q(r2)
        m = (p(r1) - 1) * (q(r2) - 1)
        Dim e As Integer, d As Integer
        d = 1
        e = 2
again:
        While hz(e, m) = False
            e = e + 1
        Wend
        While e * d Mod m <> 1
            d = d + 1
        Wend
        Randomize
        If Int(Rnd * 100) + 3 Mod 5 < 3 Then GoTo again
        Dim g As String
        If e > 10 Or d > 10 Then GoTo e
        g = creatdata("p" & i, Str(p(r1))) & creatdata("q" & i, Str(q(r2)))
        prikey = prikey + g + creatdata("key" & i, Str(e))
        username = username + g + creatdata("key" & i, Str(d))
    Next
End Sub
Public Sub addchecklist(user As String, data As String)
news.Show
news.list.AddItem user + ":" + data
    Dim rtmp() As String
    If cln = 0 Then GoTo creat
    ReDim rtmp(1 To cln) As String
    Dim i As Integer
    For i = 1 To cln
        rtmp(i) = checklist(i)
    Next
    cln = cln + 1
    ReDim checklist(1 To cln) As String
    checklist(cln) = user + ":" + data
    For i = 1 To cln - 1
        checklist(i) = rtmp(i)
    Next
    Exit Sub
creat:
    cln = 1
    ReDim checklist(1) As String
    checklist(1) = user + ":" + data
    Debug.Print "用户数据:" + checklist(1)
End Sub
Private Function searchcklst(user As String, keywords As String, kill As Boolean) As String
    Dim i As Integer, r As String
    For i = 1 To cln
    Debug.Print "当前数据库:" + checklist(i)
        Dim Ob As String, bar As String
        Ob = Split(checklist(i), ":")(0)
        bar = Split(checklist(i), ":")(1)
        Debug.Print "查询shuj:" + bar
        If InStr(bar, keywords) Or InStr(keywords, bar) Then
            If user = Ob Then r = checklist(i)
            If kill = True Then checklist(i) = ":"
            searchcklst = clear(r, user + ":")
            Debug.Print "数据查询成功"
            Exit Function
        End If
    Next
    searchcklst = "null"
End Function
Public Sub clearlist()
    Dim i As Integer, j As Integer
    For i = 1 To usernum
    
        Dim user As String, data As String
        user = bstr(userdata(i), "<", ">")
       
        data = Split(userdata(i), "<" + user + ">")(1)
        
        doing.Caption = "checking " + user
        DoEvents
        '======================================================
        Dim aa As String
        'On Error Resume Next
        aa = checkuser(user, data)
        If aa = "False" Then
            Dim errownum As Integer
            errownum = errownum + 1
            Dim rtmp() As String
            
            For j = 1 To errownum - 1
                rtmp(j) = tmp(j)
            Next
            ReDim tmp(1 To errownum) As String
            tmp(errownum) = userdata(i)
            For j = 1 To errownum - 1
                tmp(j) = rtmp(j)
            Next
        ElseIf aa = "errow" Then
            adduser user, "null"
        End If
        For j = 1 To errownum
            user = bstr(tmp(j), "<", ">")
            
            If checkuser(user, "" + Split(tmp(j), "<" + user + ">")(1)) = "True" Then
                Dim k As Integer
                For k = j To errownum - 1
                    tmp(k) = tmp(k + 1)
                Next
                tmp(errownum) = ""
                ReDim tmp(1 To errownum - 1) As String
                errownum = errownum - 1
            End If
        Next
    Next
End Sub
    Public Function Hash(ByVal ET As String) As String
        Dim BitLenString As String, KeyString As String, FileText As String
        BitLenString = "12345678"
        KeyString = ET & BitLenString
        Call Initialize(KeyString)
        '根据KeyString产生随机数序列
        FileText = ET & BitLenString
        Call DoXor(FileText)
        '根据上述随机数序列对FileText加密
        KeyString = FileText
        Call Initialize(KeyString)
        '根据上述的加密结果产生新的随机数序列
        FileText = BitLenString
        Call DoXor(FileText)
        '根据上述随机数序列对FileText加密，8位字符
        Hash = FileText
        '8位字符送作HASH值
    End Function





Private Sub Command3_Click()
Dim i As String
i = InputBox("Input client(for example:127.0.0.1;http://example)")
clients = clients + ";" + i
clientlist.AddItem i

Dim n As Integer
n = Len(clients) - Len(clear(clients, ";"))
Debug.Print clients

End Sub

Private Sub Command4_Click()
incommand.Text = incommand.Text + StrJiaMi(InputBox("Input your code"), a, b)
End Sub

Private Sub Command5_Click()
list.clear
Dim i As Integer

For i = 1 To usernum
   ' MsgBox "now we read these data" & i & " :" + userdata(i)
Next
incommand.Text = Replace(incommand.Text, "me", getAddress(username))
addmylist incommand.Text
'MsgBox "addmylist sucessful"

clearlist
doing = "ready"

savedata
incommand.Text = ""
MsgBox "OK!"
End Sub



Private Sub file1_Click()
getkey
Command5.Enabled = True
adduser username, ""
MsgBox "密钥创建成功"
End Sub

Private Sub file2_Click()
If clients = "" Then
MsgBox "You haven't input clients yet!"
Exit Sub
End If
getdata
Command5.Enabled = True
adduser username, ""
doing = "ready"
End Sub

Private Sub file3_Click()
savedata
End
End Sub

Private Sub Form_Load()
Debug.Print "____________________________________________________________________________"
clients = ";http://hupiyingwu.github.io"

myblock = "BitHive"
'MsgBox openURL("https://www.baidu.com")
End Sub
Public Sub addmylist(bar As String)
    Dim list As String
    list = bar + "//" + myblock
    'MsgBox "区块链签名为：" + aut(list)
    myblock = hexing(Hash(list))
    Dim i As Integer
    For i = 1 To usernum
        If InStr(userdata(i), "<" & username & ">") Then userdata(i) = "<" + username + ">" + Split(userdata(i), "<" + username + ">")(1) + ";" + list + "//" + aut(list)
    Next
End Sub
Private Function checkuser(user As String, data As String) As String
    If data = "null" Or data = "" Then
        checkuser = "True"
        Exit Function
    End If
    Dim maxnum As Integer, coding() As String, data2 As String
    data2 = clear(data, ";")
    maxnum = Len(data) - Len(data2)
    ReDim coding(1 To maxnum) As String
    Dim i As Integer
    For i = 1 To maxnum
        coding(i) = Split(data, ";")(i)
    Next i
    Dim block As String
    block = "BitHive"
    Dim m As Single, sys As Single, myclient As String
    m = 0
    sys = 8000
    myclient = "null"
    Dim dengji As Integer
    dengji = 0
    Dim numkey As String
    numkey = "34036"
    For i = 1 To maxnum
    'MsgBox "正在解析第" & i & "个区块，总共" & maxnum & "个区块"
    'MsgBox coding(i)
        Dim command As String, lasthash As String, auting As String
        command = Split(coding(i), "//")(0)
        'MsgBox "指令为：" + command
        lasthash = Split(coding(i), "//")(1)
        'MsgBox "哈希为" + lasthash
        auting = Split(coding(i), "//")(2)
        'MsgBox "签名为" + auting
        'MsgBox "代码以获取，正在验证签名"
        If checkaut(user, command + "//" + lasthash, auting) = False Then
            'MsgBox "签名错误"
            GoTo sus
        Else:
        'MsgBox "签名验证完成,正在验证哈希链"
        'MsgBox lasthash + "与" + block
            If lasthash = block Then
                block = hexing(Hash(command + "//" + block))
            Else:
         '   MsgBox "哈希链错误"
                GoTo sus
            End If
        End If
        '处理区块=============================================================================================================
        Dim head As String, him As String
        him = getAddress(user)
        Randomize
        head = Int(Rnd * 100) & "="
        Dim fringe As Boolean
        fringe = False
f:
        Dim key As String
        key = "wallet-pay"
        If checkstr(command, key, head) = True Then
        Debug.Print "检测到正在使用支付功能"
            Dim bar As String, address As String, money As Single, pro As String
            bar = clear2(command, key, head)
            Debug.Print "解析指令" + bar
            address = sp(bar, 1)
            Debug.Print "收款地址" + address
            money = Val(sp(bar, 2))
            Debug.Print money
            pro = sp(bar, 3)
            Debug.Print pro
            Dim prolist As String
            Debug.Print "正在检测"
            'If InStr(prolist, pro) Then MsgBox "sb"
            If InStr(prolist, pro) Then
            Debug.Print "项目已存在"
            Else:
                If money <= m Then
                    If fringe = True Then m = m + money
                    m = m - money
                    addchecklist him, command
                    prolist = prolist + pro + ";"
                    Debug.Print "支付完成" + pro
                    If address = getAddress(username) Then addmylist "wallet-rsv " + him + "," & money & "," + pro
                    
                Else:
                    Debug.Print "账户余额不足" & money & ">" & m
                End If
            End If
        End If
        key = "wallet-rsv"
        If checkstr(command, key, head) = True Then
            bar = clear2(command, key, head)
            address = sp(bar, 1)
            money = Val(sp(bar, 2))
            pro = sp(bar, 3)
            Debug.Print "data is found:" + searchcklst(address, "wallet-pay " + him + "," & money & "," + pro, False)
            If searchcklst(address, "wallet-pay " + him + "," & money & "," + pro, True) = "null" Then '==================================================
                checkuser = "False"
                Debug.Print "没有找到该用户"
                Exit Function
            Else:
                m = m + money
                Debug.Print "账户余额以增加"
            End If
        End If
        key = "wallet-mining"
        If checkstr(command, key, head) = True Then
        Debug.Print "正在挖矿，区块为" + lasthash + "账单号为" & i
            bar = clear2(command, key, head)
            Dim mn As String
            Dim hs As String
            hs = "==" + hexing(Hash(him + lasthash & Int(Val(bar))))
            If dengji > 4 Then numkey = "34"
            If InStr(hs, "==" + numkey) Then
            Debug.Print "hash正确"
                If sys > 40 Then
                    m = m + 40
                    sys = sys - 40
                    dengji = dengji + 1
                Else:
                    m = m + sys
                    sys = 0
                End If
            Else:
            Debug.Print "The hash is error:" + hs
                GoTo sus
            End If
        End If
        key = "union"
        If checkstr(command, key, head) = True Then
            bar = clear2(command, key, head)
            myclient = bar
            Debug.Print "以确定上级"
        End If
        
        
        
        
        '-----------------------------------------------------------------------------------
        If checkstr(command, "script-creat", head) = True Then
        addchecklist him, command
        Debug.Print "脚本创建成功"
        End If
        key = "script-excuse"
        
        If checkstr(command, key, head) = True Then
        Debug.Print "正在请求项目"
            bar = clear2(command, key, head)
            Dim author As String, peaple As Integer, kw As String
            author = sp(bar, 1)
            pro = sp(bar, 2)
            kw = sp(bar, 6)
            peaple = Val(sp(bar, 4))
            money = Val(sp(bar, 5))
            If user = username Then
                If scripthash = "" Then scripthash = InputBox("输入正确运行结果")
                scriptname = pro
                mypeaple = peaple
            End If
            Dim stmp As Single
            stmp = money
            Dim g As Single
            If fringe = False Then
                g = 2.2 * money + 200
                
            Else:
                g = 0
            End If
            If m >= g Then
                m = m - g
                sys = sys + g
                Dim code As String
                code = searchcklst(author, "script-creat " + pro, False)
                If code = "null" Then
                Debug.Print "未找到该用户"
                    checkuser = "False"
                    Exit Function
                ElseIf user <> username Then
                    Dim al As Boolean
                    al = False
                    Debug.Print "正在确认项目是否已经结束"
                    Dim x As Integer, abc As Boolean
                    abc = True
                    For x = i To maxnum
                        If InStr(coding(x), "script-done") Then abc = False
                    Next x
                    If abc = True Then
                    Debug.Print "项目未结束，准备执行"
                    Debug.Print "解析前" + code
                        code = clear(code, "script-creat " + pro + ",")
                        Debug.Print "解析后" + code
                        Dim part As Single
                        part = 1
                        If peaple > 0 Then
                            While InStr(code, "[" & part & "]")
                            Debug.Print "正在执行第" & part
                                If InStr(tasked, him + author + pro & i) Then
                                Debug.Print "项目已经执行"
                                Else:
                                    Dim html As String, mes As String
                                    html = StrJiMi(readata(code, Str(part)), a, b)
                                    Debug.Print "html-file" + html
                                    If InStr(html, "<MEScript>") <= 0 Then html = html + "<MEScript></MEScript>"
                                    If InStr(html, "<MEScript>") And InStr(html, "</MEScript>") Then
                                    Debug.Print "正在执行MES"
                                        mes = bstr(html, "<MEScript>", "</MEScript>")
                                        mes = clear(mes, "" & vbCrLf)
                                        Dim l As Integer, s As String
                                        l = 1
                                        While sp(mes, l) <> "null"
                                            s = sp(mes, l)
                                            Dim ok As Boolean
                                            ok = False
                                            If InStr(s, "=file.read ") Then
                                                Dim var2 As String, data3 As String
                                                var2 = Split(s, "=")(0)
                                                data3 = Split(s, " ")(1)
                                                Dim pan As String, r As String
                                            If InStr(pan, "[" + data + "]") Then r = readata(pan, data)
                                            ok = True
                                            ElseIf InStr(s, "=getURL") Then
                                                var2 = Split(s, "=")(0)
                                                data3 = Split(s, " ")(1)
                                                ok = True
                                                r = openURL(data3)
                                            ElseIf InStr(s, "=getCookie ") Then
                                                var2 = Split(s, "=")(0)
                                                data3 = Split(s, " ")(1)
                                                ok = True
                                                Dim cookie As String
                                                If InStr(cookie, "[" + data3 + "]") Then r = readata(cookie, data3)
                                            End If
                                            If ok = True Then html = Replace(html, var2, r)
                                            If InStr(s, "api") Then
                                                data3 = clear(s, "api ")
                                                addmylist data3
                                            ElseIf InStr(s, "Cookie ") Then
                                                var2 = bstr(s, " ", ",")
                                                data3 = Split(s, ",")(1)
                                                cookie = creatdata(var2, data3) + cookie
                                            ElseIf InStr(s, "shell ") Then
                                                data3 = clear(s, "shell ")
                                                al = True
                                                
                                                
                                                run data3, kw
                                            End If
                                        Wend
                                        If al = False Then
                                        Debug.Print "正在生成文件"
                                            savefile html, "index.html"
                                            Debug.Print "正在运行网页"
                                            run App.Path + "/index.html", kw
                                            While back.Caption = ""
                                                DoEvents
                                            Wend
                                            Debug.Print "正在确定版权"
                                            If InStr(back.Caption, "copyright") Then
                                                author = bstr(back.Caption, "copyright ", ".")
                                                If him <> author And searchcklst(address, "wallet-pay " + author + ",10,buy", False) <> "null" Then
                                                    addmylist "script-return " + him + "," + pro + "," & part & "," + hexing(Hash(getAddress(username) + back.Caption))
                                                    tasked = tasked + him + author & i & ","
                                                End If
                                            End If
                                        End If
                                    End If
                                    If InStr(tasked, "correct" + tasked + him + author & i & ",") Then
                                        part = part + 1
                                    Else:
                                        part = 1.1
                                    End If
                                End If
                            Wend
                        End If
                    End If
                End If
            End If
        End If
    '---------------------------------------------------------------------------------------------------------------
    key = "script-return"
            If checkstr(command, key, head) = True Then
                bar = clear2(command, key, head)
                Dim scriptr As String, scriptu As String
                address = sp(bar, 1)
                pro = sp(bar, 2)
                Dim p As Integer, h As String
                p = Int(Val((sp(bar, 3))))
                h = sp(bar, 4)
                scriptu = scriptu + address + ","
                scriptr = scriptr + creatdata(address & p, h)
                If scriptname = pro And address = getAddress(username) And mypeaple > 0 Then
                    addmylist "script-correct " + address + "," + pro + "," & p
                    mypeaple = mypeaple - 1
                    Debug.Print "hash已成功提交"
                End If
            End If
        key = "script-correct"
        If checkstr(command, key, head) = True Then
            bar = clear2(command, key, head)
            address = sp(bar, 1)
            pro = sp(bar, 2)
            p = Int(Val(sp(bar, 3)))
            If peaple > 0 Then
                peaple = peaple - 1
                sys = sys - (1.2 * stmp)
                m = m + (1.2 * stmp)
                If address = getAddress(username) Then tasked = tasked + "correct " + him + author & i & ";"
            End If
        End If
        key = "script-done"
        If checkstr(command, key, head) = True Then
            bar = clear2(command, key, head)
            Dim k As Integer
            p = 1
            While InStr(bar, "[h" & p & "]")
                p = p + 1
            Wend
            k = 1
            While sp(scriptu, k) <> "null"
                Dim ch As Integer
                For ch = 1 To p
                    If readata(scriptr, sp(scriptu, k) & p) <> hexing(Hash(sp(scriptu, k) + readata(bar, "h" & p))) Then GoTo sus
                Next
                k = k + 1
            Wend
            sys = sys - 10
            m = m + 10
            If user = username Then
                scripthash = ""
                scriptname = ""
                scriptu = ""
                scriptr = ""
                peaple = 0
            End If
            cookie = ""
        End If
    Dim fringelist(1 To 5) As String, fnum As Integer
    key = "fringe"
    If checkstr(command, key, head) = True Then
                bar = clear2(command, key, head)
                Debug.Print "检测到条件" + bar
                ok = False
                Dim newcommand As String, keywords As String
                newcommand = StrJiMi(Split(command, " then ")(1), a, b)
                Debug.Print newcommand
                key = "wallet-pay"
                If checkstr(newcommand, key, head) = True Then
                    money = Val(sp(clear2(newcommand, key, head), 2))
                    If m >= money Then
                        ok = True
                        m = m - money
                    End If
                End If
                key = "script-excuse"
                If checkstr(newcommand, key, head) = True Then
                    bar = clear2(newcommand, key, head)
                    money = Val(sp(bar, 3)) * Val(sp(bar, 4))
                    If m >= money Then
                        m = m - money
                        ok = True
                    End If
                End If
                
                If ok = True Then
                Debug.Print "开始创建"
                    If fnum = 5 Then
                        fnum = 1
                    Else:
                        fnum = fnum + 1
                    End If
                    fringelist(fnum) = command
                End If
    End If
            Dim ff As Integer
            For ff = 1 To fnum
                If InStr(fringelist(ff), "fringe ") Then
                
                    bar = bstr(fringelist(ff), "fringe ", " then ")
                    address = sp(bar, 1)
                    keywords = sp(bar, 2)
                    For k = 1 To usernum
                        If Len(userdata(k)) > 30 Then
                            If getAddress(bstr(userdata(k), "<", ">")) = address Then
                                If InStr(userdata(k), keywords) Then
                                    fringe = True
                                    command = StrJiMi(Split(fringelist(ff), "then ")(1), a, b)
                                    fringelist(ff) = ""
                                    GoTo f
                                End If
                            End If
                        End If
                    Next k
                End If
            Next ff
            If myclient <> "null" Then
                If i > 10 And sys > 7900 Then
                    g = (sys - 7900) * 0.1
                    sys = sys - g
                    addchecklist "system", "wallet-pay " + him + "," + g + ",union"
                    myclient = "null"
                End If
            End If
            key = "file-write"
            If checkstr(command, key, head) = True Then
                bar = clear2(command, key, head)
                pan = creatdata(sp(bar, 1), sp(bar, 2)) + pan
            End If
                    
    
                
    
    Next i

sus:
list.AddItem him + ":" & m
    'MsgBox "开始为timepoint赋值"
    timepoint = "<" + user + ">" + Str(maxnum) + "</" + user + ">" + timepoint
    If user = username Then Me.Caption = "Your account balance is " & m
    checkuser = "True"
End Function
Private Sub Initialize(ByVal vKeyString As String)
        Dim intI As Integer, intJ As Integer
        Randomize (Rnd(-1)) '得到初始值（种子值）
        '每次调用初始值均相同
        '根据初始值（种子值）得到随机数序列，每次调用Initialize时，初始值均相同。只要vKeyString相同，所产生的随机数序列一定相同
        For intI = 1 To Len(vKeyString)
            intJ = Rnd(-Rnd * AscW(Mid(vKeyString, intI, 1)))
            Randomize (intJ)
        Next intI
    End Sub
    Public Sub DoXor(ByRef msFileText As String)
        Dim intC As Integer
        Dim intB As Integer
        Dim lngI As Long
        For lngI = 1 To Len(msFileText)
            intC = AscW(Mid(msFileText, lngI, 1))
            intB = Int(Rnd() * 2 ^ 7)
            '选用< =127可正确处理汉字，ChrW(n):n 有一个范围
            Mid(msFileText, lngI, 1) = ChrW(intC Xor intB)
        Next lngI
 
    End Sub
Private Function openURL(url As String) As String
    web.Navigate url
    Do While web.Busy
        DoEvents
    Loop
    openURL = web.Document.documentElement.outerHTML
End Function
Private Function clear2(command As String, code As String, head As String) As String
    clear2 = Replace(head + command, head + code + " ", "")

End Function
Private Function sp(bar As String, num As Integer) As String
    On Error GoTo e
    sp = Split(bar, ",")(num - 1)
    Exit Function
e:
    sp = "null"
End Function
Private Function checkstr(command As String, code As String, head As String) As Boolean
    If InStr(head + command, head + code + " ") Then
        checkstr = True
    Else:
        checkstr = False
    End If

End Function
Private Function getAddress(user As String) As String
    getAddress = hexing(Hash(user))
End Function
Public Sub run(url As String, kw As String)
    Dim ads As New showads
    ads.Show
    ads.web.Navigate url
    ads.Caption = kw
End Sub

Private Sub network1_Click()
Dim RetVal
RetVal = Shell(App.Path + "\saves\server.exe", 1)
End Sub

Private Sub network2_Click()
Dim i As String
i = InputBox("请输入节点地址")
clients = clients + ";" + i
clientlist.AddItem i

Dim n As Integer
n = Len(clients) - Len(clear(clients, ";"))

End Sub

Private Sub sa2_Click()
writecode.Show

End Sub

Private Sub sa3_Click()
incommand.Text = "script-excuse " + InputBox("输入脚本作者的收款地址") + ","
incommand.Text = incommand.Text + InputBox("输入脚本作者的项目名称") + ","
incommand.Text = incommand.Text + InputBox("输入代码片段总数") + ","
incommand.Text = incommand.Text + InputBox("输入每个代码片段所需支付的佣金") + ","
incommand.Text = incommand.Text + InputBox("输入脚本网址跳转关键词")
End Sub

Private Sub sa4_Click()
incommand.Text = "file-write " + InputBox("输入文件名") + "," + InputBox("输入文件内容")

End Sub

Private Sub sa5_Click()
incommand.Text = "script-done " + scripthash
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
getdata
End Sub

Private Sub ud1_Click()
blockchain.Show

blockchain.list.list(0) = Me.list.list(0)

End Sub

Private Sub ud2_Click()
news.Show

End Sub

Private Sub wallet1_Click()
Dim address As String, money As String
address = InputBox("请输入对方收款地址")
money = InputBox("请输入支付金额")
incommand.Text = "wallet-pay " + address + "," + money + "," + InputBox("请输入注释")
MsgBox "指令已生成"

End Sub

Private Sub wallet2_Click()
showaddress.Show
showaddress.address.Text = getAddress(username)
End Sub

Private Sub wallet3_Click()
Dim lucknum As Long
lucknum = 0
Debug.Print "你的区块为" + myblock
Debug.Print "hash:" + hexing(Hash(getAddress(username) + myblock & Int(lucknum)))
Dim h As String
h = "==" + hexing(Hash(getAddress(username) + myblock & Int(lucknum)))
Dim numkey As String
numkey = "34036"
If dignum > 4 Then numkey = "34"
Debug.Print "这是第" & dignum & "次挖矿，关键词为" & numkey
While InStr(h, "==" & numkey) <= 0
    doing.Caption = "正在尝试" & lucknum
    DoEvents
    lucknum = lucknum + 1
Wend
dignum = dignum + 1
incommand.Text = "wallet-mining " & lucknum
doing.Caption = "准备就绪"
End Sub
