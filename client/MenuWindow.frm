VERSION 5.00
Begin VB.Form MenuWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ò˵�"
   ClientHeight    =   3135
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
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
