VERSION 5.00
Begin VB.Form MenuWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ò˵�"
   ClientHeight    =   3128
   ClientLeft      =   80
   ClientTop       =   672
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3128
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu msgMenu 
      Caption         =   "��Ϣ�˵�"
      Begin VB.Menu copyMsg 
         Caption         =   "����"
      End
      Begin VB.Menu banMsg 
         Caption         =   "����"
         Begin VB.Menu min1ban 
            Caption         =   "1����"
         End
         Begin VB.Menu min5ban 
            Caption         =   "5����"
         End
         Begin VB.Menu min10ban 
            Caption         =   "10����"
         End
         Begin VB.Menu customban 
            Caption         =   "�Զ���ʱ��..."
         End
      End
      Begin VB.Menu kickGroup 
         Caption         =   "�Ƴ�Ⱥ��"
      End
   End
   Begin VB.Menu groupMenu 
      Caption         =   "��˵�"
      Begin VB.Menu quitGroup 
         Caption         =   "�˳�����"
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
    w = InputBox("����Ҫ���ԶԷ��೤ʱ�䣿��������", , 60)
    du = Val(w)
    If du <= 0 Then
        MsgBox "��������ȷ�����֣�", 48
        GoTo reinput
    End If
    Client.Winsock1.SendData "addban;" & groups(MainPage.selectIndex).id & ";" & du & ";" & id & vbCrLf
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
    If ECore.SimpleMsg("��ȷ��Ҫִ�д˲������˲��������档", quitGroup.Caption & "�顰" & groups(groupId).Name & "��", StrArray("ȷ��", "ȡ��"), UseBlur:=False) = 0 Then
        Client.Winsock1.SendData "quitgroup;" & groups(groupId).id & vbCrLf
    End If
End Sub
