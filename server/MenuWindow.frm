VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MenuWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ò˵�"
   ClientHeight    =   3120
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog FileOpens 
      Left            =   2400
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "���������"
      Filter          =   "�����˽ű�|*.vbs|"
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
         Caption         =   "��ɢ"
      End
   End
   Begin VB.Menu robotMenu 
      Caption         =   "�����˲˵�"
      Begin VB.Menu robotBtn 
         Caption         =   "���������..."
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
    w = InputBox("����Ҫ���ԶԷ��೤ʱ�䣿��������", , 60)
    du = Val(w)
    If du <= 0 Then
        MsgBox "��������ȷ�����֣�", 48
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
    If ECore.SimpleMsg("��ȷ��Ҫִ�д˲������˲��������档", quitGroup.Caption & "�顰" & groups(groupid).Name & "��", StrArray("ȷ��", "ȡ��"), UseBlur:=False) <> 0 Then Exit Sub
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
            MsgBox "����ɹ���", 64
        End If
    Else
        MsgBox "ʹ�ð�����" & vbCrLf & vbCrLf & machine(index).Eval("Guidence"), 64, robotBtn(index).Caption
    End If
End Sub
