VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form showads 
   Caption         =   "网页展示"
   ClientHeight    =   5280
   ClientLeft      =   9240
   ClientTop       =   4500
   ClientWidth     =   9855
   LinkTopic       =   "Form2"
   ScaleHeight     =   5280
   ScaleWidth      =   9855
   Begin VB.CommandButton Command1 
      Caption         =   "结束运行"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   4320
      Top             =   2400
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9375
      ExtentX         =   16536
      ExtentY         =   8070
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
      Location        =   ""
   End
End
Attribute VB_Name = "showads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormOldWidth As Long
'保存窗体的原始宽度
Private FormOldHeight As Long
'保存窗体的原始高度

Public Sub ResizeInit(FormName As Form)
    Dim Obj As Control
    FormOldWidth = FormName.ScaleWidth
    FormOldHeight = FormName.ScaleHeight
    On Error Resume Next
    For Each Obj In FormName
        Obj.Tag = Obj.Left & " " & Obj.Top & " " & Obj.Width & " " & Obj.Height & " "
    Next Obj
    On Error GoTo 0
End Sub

Public Sub ResizeForm(FormName As Form)
    Dim Pos(4) As Double
    Dim i      As Long, TempPos As Long, StartPos As Long
    Dim Obj    As Control
    Dim ScaleX As Double, ScaleY As Double

    ScaleX = FormName.ScaleWidth / FormOldWidth
    ScaleY = FormName.ScaleHeight / FormOldHeight
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        For i = 0 To 4
            TempPos = InStr(StartPos, Obj.Tag, " ", vbTextCompare)
            If TempPos > 0 Then
                Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(i) = 0
            End If

            Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
        Next i
    Next Obj
    On Error GoTo 0
End Sub

Private Sub Command1_Click()
MsgBox "网页浏览已超时！"
Form1.back.Caption = "time out"
Unload Me
End Sub

Private Sub Form_Resize()
    Call ResizeForm(Me) '确保窗体改变时控件随之改变
End Sub
Private Sub Form_Load()
    Call ResizeInit(Me) '在程序装入时加入
 '以下四句是运行使窗体最大化
Me.Top = 0
    Me.Left = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
End Sub

Private Sub Timer1_Timer()
MsgBox "网页浏览已超时！"
Form1.back.Caption = "time out"
Unload Me

End Sub

Private Sub web_DocumentComplete(ByVal pDisp As Object, url As Variant)
    If InStr(url, Me.Caption) Then Form1.back.Caption = url
End Sub
