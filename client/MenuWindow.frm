VERSION 5.00
Begin VB.Form MenuWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "放置菜单"
   ClientHeight    =   3120
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Formater 
      Caption         =   "Label1"
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Menu msgMenu 
      Caption         =   "消息菜单"
      Begin VB.Menu copyMsg 
         Caption         =   "复制"
      End
      Begin VB.Menu banMsg 
         Caption         =   "禁言"
         Begin VB.Menu min1ban 
            Caption         =   "1分钟"
         End
         Begin VB.Menu min5ban 
            Caption         =   "5分钟"
         End
         Begin VB.Menu min10ban 
            Caption         =   "10分钟"
         End
         Begin VB.Menu customban 
            Caption         =   "自定义时长..."
         End
      End
      Begin VB.Menu undoBan 
         Caption         =   "解除禁言"
         Visible         =   0   'False
      End
      Begin VB.Menu kickGroup 
         Caption         =   "移出群聊"
      End
      Begin VB.Menu nonono 
         Caption         =   "  "
      End
   End
   Begin VB.Menu groupMenu 
      Caption         =   "组菜单"
      Begin VB.Menu quitGroup 
         Caption         =   "退出该组"
      End
   End
End
Attribute VB_Name = "MenuWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id As Integer, groupId As Integer

Private Sub copyMsg_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText selectMsg.Content
End Sub

Private Sub customban_Click()
    Dim w As String, du As Long
    On Error GoTo reinput
reinput:
    w = InputBox("您想要禁言对方多长时间？（秒数）", , 60)
    du = Val(w)
    If du <= 0 Then
        MsgBox "请输入正确的数字！", 48
        GoTo reinput
    End If
    Client.Winsock1.SendData "addban;" & groups(MainPage.selectIndex).id & ";" & du & ";" & id & vbCrLf
End Sub


Private Sub kickGroup_Click()
    Client.Winsock1.SendData "deletemember;" & groups(MainPage.selectIndex).id & ";" & id & vbCrLf
End Sub

Private Sub min10ban_Click()
    Client.Winsock1.SendData "addban;" & groups(MainPage.selectIndex).id & ";600;" & id & vbCrLf
End Sub

Private Sub min1ban_Click()
    Client.Winsock1.SendData "addban;" & groups(MainPage.selectIndex).id & ";60;" & id & vbCrLf
End Sub

Private Sub min5ban_Click()
    Client.Winsock1.SendData "addban;" & groups(MainPage.selectIndex).id & ";300;" & id & vbCrLf
End Sub

Private Sub quitGroup_Click()
    If ECore.SimpleMsg("您确定要执行此操作？此操作不可逆。", quitGroup.Caption & "组“" & groups(groupId).Name & "”", StrArray("确定", "取消"), UseBlur:=False) = 0 Then
        Client.Winsock1.SendData "quitgroup;" & groups(groupId).id & vbCrLf
    End If
End Sub

Private Sub undoBan_Click()
    Client.Winsock1.SendData "addban;" & groups(MainPage.selectIndex).id & ";0;" & id & vbCrLf
End Sub
