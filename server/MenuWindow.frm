VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MenuWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "放置菜单"
   ClientHeight    =   3120
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog FileOpens 
      Left            =   2400
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "导入机器人"
      Filter          =   "机器人脚本|*.vbs|"
   End
   Begin VB.Label Formater 
      Caption         =   "Label1"
      Height          =   15
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   1095
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
      Begin VB.Menu kickGroup 
         Caption         =   "移出群聊"
      End
   End
   Begin VB.Menu groupMenu 
      Caption         =   "组菜单"
      Begin VB.Menu quitGroup 
         Caption         =   "解散"
      End
   End
   Begin VB.Menu robotMenu 
      Caption         =   "机器人菜单"
      Begin VB.Menu robotBtn 
         Caption         =   "导入机器人..."
         Index           =   0
      End
   End
End
Attribute VB_Name = "MenuWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id As Integer, groupid As Integer

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
    Server.ProcessBan groups(MainPage.selectIndex).id, id, du
End Sub

Private Sub min10ban_Click()
    Server.ProcessBan groups(MainPage.selectIndex).id, id, 600
End Sub

Private Sub min1ban_Click()
    Server.ProcessBan groups(MainPage.selectIndex).id, id, 60
End Sub

Private Sub min5ban_Click()
    Server.ProcessBan groups(MainPage.selectIndex).id, id, 300
End Sub

Private Sub quitGroup_Click()
    If ECore.SimpleMsg("您确定要执行此操作？此操作不可逆。", quitGroup.Caption & "组“" & groups(groupid).Name & "”", StrArray("确定", "取消"), UseBlur:=False) <> 0 Then Exit Sub
    DeleteGroup groups(groupid).id
    For Each w In Server.Winsock
        If w.State = 7 Then w.SendData "deletegroup;" & groups(groupid).id & vbCrLf
        DoEvents
    Next
End Sub

Private Sub robotBtn_Click(index As Integer)
    If index = 0 Then
        FileOpens.ShowOpen
        If FileOpens.filename <> "" Then
            Dim Name As String, t() As String
            t = Split(FileOpens.filename, "\")
            Name = t(UBound(t))
            FileCopy FileOpens.filename, App.path & "\robots\" & Name
            Robots.ImportRobot Name
            MsgBox "导入成功！", 64
        End If
    Else
        MsgBox "使用帮助：" & vbCrLf & vbCrLf & machine(index).Eval("Guidence"), 64, robotBtn(index).Caption
    End If
End Sub
