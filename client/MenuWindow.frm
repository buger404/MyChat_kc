VERSION 5.00
Begin VB.Form MenuWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "放置菜单"
   ClientHeight    =   3135
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
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
      Begin VB.Menu kickGroup 
         Caption         =   "移出群聊"
      End
   End
End
Attribute VB_Name = "MenuWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id As Integer

Private Sub copyMsg_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText selectMsg.Content
End Sub

Private Sub Form_Load()

End Sub
