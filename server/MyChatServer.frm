VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Server 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "�����"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   Icon            =   "MyChatServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   503
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   3660
      Width           =   6300
   End
   Begin VB.CommandButton OCR 
      Caption         =   "ͼƬת����"
      Height          =   615
      Left            =   6480
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton Audio 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����ʶ��ճ��"
      Height          =   615
      Left            =   6480
      TabIndex        =   11
      ToolTipText     =   "������˵��"
      Top             =   1920
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      ToolTipText     =   "������˵��"
      Top             =   1440
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����¼"
      Height          =   375
      Left            =   6480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   9
      ToolTipText     =   "������D��"
      Top             =   960
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "MyChatServer.frx":1BCC2
      Top             =   720
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   270
      Left            =   5565
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3615
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSWinsockLib.Winsock lis 
      Left            =   6960
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   6480
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "����IP��Port�������С���ɣ�"
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ſͻ���������"
      Height          =   255
      Left            =   4230
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�Ͽ���"
      Height          =   270
      Left            =   3015
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim State As Boolean, pop As Single
Dim grpExistId As Integer
Dim pypid
Dim g As String, q As Single, m As Single
Dim MainPage As MainPage, IPPage As IPPage

Public Sub Command1_Click()
    Dim C As Single
    C = Val(Text2.Text)
    
    If C > Winsock.ubound Then
        MsgBox ("û�д��û�")
    Else
        If Winsock(C).State = 7 Then
            Winsock(C).close
            MsgBox ("�ѶϿ�")
        End If
    End If

    pop = pop - 1
    Me.Caption = lis.LocalIP & " - " & "������" & pop & "��"
    Text2.Text = ""
End Sub
'===============================================================================================================
'Emerald��ܲ���
Private Sub InitEmeraldFramework()
    '����Emerald
    StartEmerald Me.hwnd, 1100, 600, False
    'ScaleGame Screen.Width / Screen.TwipsPerPixelX / 1280, ScaleDefault
    '����������Ⱦ
    Set EF = New GFont
    EF.MakeFont "΢���ź�"
    'ʵ����ҳ�����������
    Set ECore = New GMan
    'ʵ����ҳ�������
    Set MainPage = New MainPage
    Set IPPage = New IPPage
    '��ʾ
    DrawTimer.Enabled = True
    ECore.ActivePage = "IPPage"
End Sub
Private Sub UnloadEmeraldFramework()
    DrawTimer.Enabled = False
    EndEmerald
End Sub
Private Sub DrawTimer_Timer()
    '���»���
    ECore.Display
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    '���������Ϣ
    UpdateMouse x, y, 1, button
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    '���������Ϣ
    If Mouse.State = 0 Then
        UpdateMouse x, y, 0, button
    Else
        Mouse.x = x: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    '���������Ϣ
    UpdateMouse x, y, 2, button
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub
'===============================================================================================================
Public Sub Command2_Click()

    If Text4.Text = "" Then VBA.Beep: Exit Sub
    
    Dim S As Single
    S = 1
    Do While (S <= Winsock.ubound)
        If Winsock(S).State = 7 Then
            Winsock(S).SendData "msg;" + "groupid;" + "����" + ";id;" + Base64EncodeString(Text4.Text) + ";"
            DoEvents
        End If
        S = S + 1
    Loop
    
    Text3.Text = "�ң�" & Text4.Text & vbCrLf & Text3.Text
    Text4.Text = ""

End Sub

Public Sub Command3_Click()
    Text3.Text = ""
    Text4.Text = ""
End Sub

Public Sub Command4_Click()
    Open App.path & "\" & "�������Ϣ��¼" & Str(q) & ".txt" For Output As #1
    Print #1, Text3.Text
    Close #1
    q = q + 1
End Sub

Public Sub Command5_Click()
    If State = False Then
        State = True
        Command5.Caption = "�������"
        Dim S As Single
        g = "�����������˽���"
        S = 1
        Do While (S <= Winsock.ubound)
            If Winsock(S).State = 7 Then
                Winsock(S).SendData g
                DoEvents
            End If
            S = S + 1
        Loop
        Text3.Text = g & vbCrLf & Text3.Text
    Else
        State = False
        Command5.Caption = "����"
        g = "�������ر��˽���"
        S = 1
        Do While (S <= Winsock.ubound)
            If Winsock(S).State = 7 Then
                Winsock(S).SendData g
                DoEvents
            End If
            S = S + 1
        Loop
        Text3.Text = g & vbCrLf & Text3.Text
    End If
End Sub

Public Sub Audio_Click()
    ShellEx "python """ & App.path & "\" & "server.py"" -y "
    If Dir("audio_text.txt") = "" Then Shell "python """ & App.path & "\" & "server.py"" -y "
    
    Dim strfile As String
    strfile = "audio_text.txt"
    Open strfile For Input As #1
        Text4.Text = StrConv(InputB(FileLen(strfile), 1), vbUnicode)
    Close #1
    Kill "audio.wav"
    Kill "audio_text.txt"
    Kill "audio.pcm"
End Sub



Public Sub OCR_Click()
    ShellEx "python """ & App.path & "\" & "server.py"" -t "
    If Dir("OCR_text.txt") = "" Then Shell "python """ & App.path & "\" & "server.py"" -t "
    
    Dim strfile As String
    strfile = "OCR_text.txt"
    Open strfile For Input As #1
        Text4.Text = StrConv(InputB(FileLen("OCR_text.txt"), 1), vbUnicode)
    Close #1
    Kill "OCR_text.txt"
End Sub

Private Sub Form_Load()
    ReDim groups(0)
    '������
    AddGroup 1, 1, True, "����������1"
    AddGroup 2, 1, False, "δ�������"
    AddGroup 3, 1, True, "����������2"
    AddGroup 4, 1, True, "testtest"
    AddGroup 5, 1, True, "hash"
    AddMessage 1, 1, "������Ա", "�ҷ�����һ����Ϣ��������������"
    AddMessage 1, 2, "������Ա", "�һ��ܻ���" & vbCrLf & "��������"
    AddMessage 1, -1, "ϵͳ��Ϣ", "�������ԣ��Ź֡�"
    AddMessage 1, -2, "��ʦ", "��Ҫ�ҷ���Ϣ"
    userId = -2
    
    Call InitEmeraldFramework
    Set Shadow = New aShadow
    With Shadow
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 8
            .Transparency = 20
        End If
    End With
    
    Text3.Visible = False: Text4.Visible = False
    Me.Show
    
    pypid = Shell("python """ & App.path & "\" & "server.py"" -o " & lis.LocalIP, 6)
    
<<<<<<< HEAD
    Call AddGroup(0, -2, True, "����")
    grpExistId = 0
    
    
    Text3.Visible = True: Text4.Visible = True
=======
    Text3.Visible = False: Text4.Visible = True
>>>>>>> 3b07fd90bb91919c2b4047f89cff85163e999376
    Command5.Enabled = False
    State = False
    m = 1
    
    lis.LocalPort = 2001
    lis.Listen
    
    '��׼��������
    Text3.Move 300 + 0, 60, Me.ScaleWidth - 300, Me.ScaleHeight - 60 - 120
    Text4.Move 300 + 50, Me.ScaleHeight - 80 + 25, Me.ScaleWidth - 245 - 300, 80 - 50
    
    Me.Caption = lis.LocalIP & " - " & "������" & pop & "��"
    
    ECore.NewTransform , , "MainPage"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Server.Command4_Click
    TerminateProcess OpenProcess(PROCESS_TERMINATE, 0, pypid), 0
    
    Set Shadow = Nothing
    Call UnloadEmeraldFramework
End Sub

Private Sub lis_ConnectionRequest(ByVal requestID As Long)
    Load Winsock(m)
    Command5.Enabled = True
    
    pop = Winsock.ubound
    
    If Winsock(m).State = sckClosed Then
        Winsock(m).Accept requestID
    End If
    
    Call SetJoinState(0, True)
    
    Me.Caption = lis.LocalIP & " - " & "������" & pop & "��"
    
    m = m + 1
End Sub

Private Sub Text4_Change()
    '�Զ������ı����С
    Dim Line As Long
    Line = UBound(Split(Text4.Text, vbCrLf)) + 1
    If Line <= 0 Then Line = 1
    Dim Border As Integer, Height As Long
    Border = IIf(Line > 1, 1, 0)
    Height = Line * 30
    If Height > Me.ScaleHeight - 120 Then Height = Me.ScaleHeight - 120 '��ֹ���������
    If Text4.BorderStyle <> Border Then Text4.BorderStyle = Border
    If Text4.Height <> Height Then
        Text4.Height = Height
        Text4.Top = Me.ScaleHeight - 80 + 25 - Height + 30
    End If
End Sub

Private Sub Winsock_Close(index As Integer)
    pop = pop - 1
    Me.Caption = lis.LocalIP & " - " & "������" & pop & "��"
End Sub

Private Sub Winsock_DataArrival(index As Integer, ByVal bytesTotal As Long)
    If State = True Then Exit Sub
    
    Dim strSplit
    Dim id As Integer
    Dim MsgType As String
    Dim grpId As String
    Dim name As String
    Dim MsgContent As String
    Dim strData As String
    Winsock(index).GetData strData
    
    Dim S As Single
    S = 1
    Do While (S <= Winsock.ubound)
        If Winsock(S).State = 7 Then
            Winsock(S).SendData strData
            DoEvents
        End If
        S = S + 1
    Loop
    
    strSplit = Split(strData, ";")
    id = index
    MsgType = strSplit(0)
<<<<<<< HEAD

    
    Select Case MsgType
    Case "msg"
    name = strSplit(2)
    grpId = strSplit(1)
    MsgContent = strSplit(4)
    MsgContent = Base64DecodeString(MsgContent)
    Text3.Text = name + ":" + MsgContent + "   #" + Str(id) + "#" + Str(grpId) + "#" + vbCrLf + Text3.Text
    Case "picmsg"
    Case "addgroup"
    Case "okgroup"
    Case "creategroup"
    
    Dim grpCreateName As String
    
    End Select
=======
    name = strSplit(2)
    MsgContent = strSplit(4)
    MsgContent = Base64DecodeString(MsgContent)
>>>>>>> 3b07fd90bb91919c2b4047f89cff85163e999376
    
End Sub
