VERSION 5.00
Begin VB.Form writecode 
   Caption         =   "code edicter"
   ClientHeight    =   7110
   ClientLeft      =   7830
   ClientTop       =   3690
   ClientWidth     =   12105
   LinkTopic       =   "Form2"
   ScaleHeight     =   7110
   ScaleWidth      =   12105
   Begin VB.CommandButton Command2 
      Caption         =   "next"
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox code 
      Height          =   3855
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "writecode.frx":0000
      Top             =   1800
      Width           =   9375
   End
   Begin VB.TextBox proname 
      Height          =   270
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label realcode 
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   5880
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "code:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "name:"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "writecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Integer
Const a As Byte = 20 '密钥
Const b As Byte = 40 '密钥

Private Function StrJiaMi(ByVal strSource As String, ByVal Key1 As Byte, _
ByVal Key2 As Integer) As String
Dim bLowData As Byte
Dim bHigData As Byte
Dim i As Integer
Dim strEncrypt As String
Dim strChar As String
For i = 1 To Len(strSource)
'从待加（解）密字符串中取出一个字符
strChar = Mid(strSource, i, 1)
'取字符的低字节和Key1进行异或运算
bLowData = AscB(MidB(strChar, 1, 1)) Xor Key1
'取字符的高字节和K2进行异或运算
bHigData = AscB(MidB(strChar, 2, 1)) Xor Key2
'将运算后的数据合成新的字符
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

Private Sub Command1_Click()
Form1.incommand.Text = "script-creat " + proname.Text + "," + realcode.Caption

End Sub

Private Sub Command2_Click()
code.Text = Replace(code.Text, vbCrLf, "")
num = num + 1
realcode.Caption = realcode.Caption + "[" & num & "]" + StrJiaMi(code.Text, a, b) + "[/" & num & "]"
code.Text = ""
End Sub

