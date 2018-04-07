VERSION 5.00
Begin VB.Form showaddress 
   Caption         =   "your address"
   ClientHeight    =   3030
   ClientLeft      =   16140
   ClientTop       =   5310
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox address 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "showaddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
