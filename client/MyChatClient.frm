VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Client 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8760
   ForeColor       =   &H00FFFFFF&
   Icon            =   "MyChatClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   584
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton logBtn 
      Caption         =   "logBtn"
      Height          =   615
      Left            =   7080
      TabIndex        =   16
      Top             =   5040
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog ColorPad 
      Left            =   1248
      Top             =   936
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
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
      Height          =   480
      Left            =   2184
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   312
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   3
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1560
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   312
      Width           =   480
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
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
      Height          =   468
      Left            =   936
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   312
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   312
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   312
      Width           =   480
   End
   Begin VB.TextBox Text2 
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
      Height          =   588
      Left            =   312
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5304
      Width           =   636
   End
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   16
      Left            =   312
      Top             =   936
   End
   Begin VB.CommandButton OCR 
      Caption         =   "ͼƬת����"
      Height          =   495
      Left            =   7680
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Audio 
      Caption         =   "����ʶ��"
      Height          =   495
      Left            =   7680
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton save3 
      Caption         =   "����"
      Height          =   495
      Left            =   6552
      TabIndex        =   13
      ToolTipText     =   "������D��"
      Top             =   2340
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton save2 
      Caption         =   "����"
      Height          =   495
      Left            =   6552
      TabIndex        =   12
      ToolTipText     =   "������D��"
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DJ Mode"
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      ToolTipText     =   "�������HARDBASSʹ��Ŷ��"
      Top             =   2325
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����"
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      ToolTipText     =   "�ٺ٣���ˮ��ʼ��"
      Top             =   1950
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���±�"
      Height          =   375
      Left            =   7680
      TabIndex        =   7
      ToolTipText     =   "���ԼǱʼ��ޣ�"
      Top             =   1575
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����"
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      ToolTipText     =   "����ʲô�ɣ�"
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   375
      Left            =   6396
      TabIndex        =   5
      Top             =   4056
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���"
      Height          =   615
      Left            =   5304
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   615
      Left            =   5304
      TabIndex        =   1
      Top             =   2340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   780
      Top             =   936
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim grpId As Integer
Dim i As Single, p As Single, q As Single, dr As Single
Dim IPPage As IPPage
Public LBtnColor As Long, RBtnColor As Long
Public DotMode As Boolean
Dim DrawX As Single, DrawY As Single
Dim Shadow As aShadow

Public Sub Audio_Click()
    ShellEx "python """ & App.path & "\" & "client.py"" -a "
    If Dir("audio_text.txt") = "" Then ShellEx "python """ & App.path & "\" & "client.py"" -a "
    Dim strfile As String
    strfile = "audio_text.txt"
    Open strfile For Input As #1
        Text2.Text = StrConv(InputB(FileLen(strfile), 1), vbUnicode)
    Close #1
    Kill "audio_text.txt"
    Kill "audio.pcm"
    Kill "audio.wav"
End Sub

Private Sub Command2_Click()
    '����
    Call SendMsg
End Sub

Public Sub Command4_Click()
    Picture1.Cls
End Sub



Private Sub logIn()
    'If Dir("id_info.txt") <> "" Then Kill ("id_info.txt")
    'If Dir("id_info.txt") <> "" Then Kill "id_info.txt"
    'If Dir("face.png") <> "" Then Kill "face.png"
    
    ShellEx "python """ & App.path & "\" & "client.py"" -l " & Winsock1.RemoteHost
    If Dir("id_info.txt") = "" Then MsgBox "���޴��ˣ�����ע�᣿", 16, "��½ʧ��": End
    Open "id_info.txt" For Input As 1
    A = StrConv(InputB(FileLen("id_info.txt"), 1), vbUnicode)
    S = Split(A, ",")
    Close #1
    If S(0) = "404" Then MsgBox "ip��ַ��������ip��ַ", 16, "ip��ַ����": End
    Me.Caption = S(0)
    
    'If Dir("id_info.txt") <> "" Then Kill "id_info.txt"
    'If Dir("face.png") <> "" Then Kill "face.png"
    
    Winsock1.RemotePort = 2001
    If Winsock1.State = sckClosed Then
        Winsock1.Connect
    End If
    
    Text1.Text = "welcome," & Me.Caption & "!"
End Sub

Public Sub OCR_Click()
    ShellEx "python """ & App.path & "\" & "client.py"" -o "
    If Dir("OCR_text.txt") = "" Then ShellEx "python """ & App.path & "\" & "client.py"" -o "
    Dim strfile As String
    strfile = "OCR_text.txt"
    Open strfile For Input As #1
        Text2.Text = StrConv(InputB(FileLen(strfile), 1), vbUnicode)
    Close #1
    Kill "OCR_text.txt"
    Kill "ocr_img.png"
End Sub
'===============================================================================================================
'Emerald��ܲ���
Private Sub InitEmeraldFramework()
    '����Emerald
    StartEmerald Me.hwnd, 1100, 600, False
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

Private Sub Form_Load()
    ReDim groups(0): ReDim bans(0)
    '������
    AddGroup 1, 1, True, "����"
    '����ͷ���˲�һ���Ƕ�̬�ģ���Ҫ��ȡ��
    'userId = -2: userName = "��ʦ"
    
    Call InitEmeraldFramework
    Set Shadow = New aShadow
    With Shadow
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 8
            .Transparency = 20
        End If
    End With
    
    i = 1: p = 1: q = 1: dr = 1
    DrawX = -100

    save2.Visible = False
    
    Text2.Enabled = False
    Command2.Enabled = False

    Dim A As String
    Dim S

    Dim o As Object
    On Error Resume Next
    For Each o In Me.Controls
        If Not (o Is Me) Then o.Visible = False
    Next
    
    Me.Show
    Do
        ECore.Display
        DoEvents
    Loop Until Winsock1.RemoteHost <> ""
    

    If Winsock1.RemoteHost <> "" Then logIn
    grpId = 1
    
    '��׼��������
    Text5.Move 0, 60, Me.ScaleWidth, Me.ScaleHeight - 60 - 120
    Text1.Move Text5.Left, Text5.Top, Text5.Width, Text5.Height
    Picture1.Move Text5.Left, Text5.Top, Text5.Width, Text5.Height
    Picture2.Move Text5.Left, Text5.Top, Text5.Width, Text5.Height
    Text2.Move 300 + 50, Me.ScaleHeight - 80 + 25, Me.ScaleWidth - 245 - 300, 80 - 50
    
    Text1.Visible = False
    Text5.Visible = False
    Picture1.Visible = False
    Picture2.Visible = False
    
    Command4.Visible = False
    
    ECore.NewTransform , , "MainPage"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Client.save3_Click
    Set Shadow = Nothing
    Call UnloadEmeraldFramework
End Sub

Public Sub Option1_Click()
    save2.Visible = False
    save3.Visible = False
    
    Command4.Visible = True
    
    Picture1.Visible = True
    Text1.Visible = False
    Text5.Visible = False
    Picture2.Visible = False
End Sub
Public Sub Option2_Click()
    save2.Visible = True
    save3.Visible = False
    
    Command4.Visible = False
    
    Picture1.Visible = False
    Text1.Visible = False
    Text5.Visible = True
    Picture2.Visible = False
End Sub
Public Sub Option3_Click()
    save2.Visible = False
    save3.Visible = True
    
    Command4.Visible = False
    
    Picture1.Visible = False
    Text1.Visible = False
    Text5.Visible = False
    Picture2.Visible = False
End Sub

Public Sub Option4_Click()
    save2.Visible = False
    save3.Visible = False
    
    Command4.Visible = False
    
    Picture1.Visible = False
    Text1.Visible = False
    Text5.Visible = False
    Picture2.Visible = True
End Sub

Private Sub Picture1_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If button > 0 Then
        If DrawX = -100 Or DotMode Then
            Picture1.PSet (x, y), IIf(button = 1, LBtnColor, RBtnColor)
        Else
            Picture1.Line (DrawX, DrawY)-(x, y), IIf(button = 1, LBtnColor, RBtnColor)
        End If
    End If
    DrawX = x: DrawY = y
End Sub

Private Sub Picture1_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    DrawX = -100
End Sub

Public Sub save2_Click()
    Open App.path & "\" & Me.Caption & "�ıʼ�" & Str(p) & ".txt" For Output As #1
    Print #1, Text5.Text
    Close #1
    MsgBox "�ѳɹ�����ʼǣ�" & vbCrLf & App.path & "\" & Me.Caption & "�ıʼ�" & Str(p) & ".txt", 64, "����ɹ�"
    p = p + 1
End Sub
Public Sub save3_Click()
    Open App.path & "\" & Me.Caption & "����Ϣ��¼" & Str(q) & ".txt" For Output As #1
    Print #1, Text1.Text
    Close #1
    MsgBox "�ѳɹ�������Ϣ��¼��" & vbCrLf & App.path & "\" & Me.Caption & "����Ϣ��¼" & Str(q) & ".txt", 64, "����ɹ�"
    q = q + 1
End Sub
Public Sub saveDrawing()
    SavePicture Client.Picture1.Image, App.path & "\" & Me.Caption & "��Ϳѻ" & Str(dr) & ".bmp"
    MsgBox "�ѳɹ�����Ϳѻ��" & vbCrLf & App.path & "\" & Me.Caption & "��Ϳѻ" & Str(dr) & ".bmp", 64, "����ɹ�"
    dr = dr + 1
End Sub

Private Sub Text2_Change()
    If Picture2.Visible = True Then
        Picture2.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    End If

    If i >= 15 Then i = 1
    i = i + 1
    
    '�Զ������ı����С
    Dim Line As Long
    Line = UBound(Split(Text2.Text, vbCrLf)) + 1
    If Line <= 0 Then Line = 1
    Dim Border As Integer, Height As Long
    Border = IIf(Line > 1, 1, 0)
    Height = Line * 30
    If Height > Me.ScaleHeight - 120 Then Height = Me.ScaleHeight - 120 '��ֹ���������
    If Text2.BorderStyle <> Border Then Text2.BorderStyle = Border
    If Text2.Height <> Height Then
        Text2.Height = Height
        Text2.Top = Me.ScaleHeight - 80 + 25 - Height + 30
    End If
End Sub
Public Sub SendMsg()
    If Winsock1.State <> 7 Then Exit Sub
    If Text2.Text = "" Then
        VBA.Beep
    Else
        Winsock1.SendData "msg;" + Str(grpId) + ";" + Me.Caption + ";id;" + Base64EncodeString(Text2.Text) + ";"
        'Winsock1.SendData Text2.Text
        Text2.Text = ""
    End If
End Sub

Public Sub Command3_Click()
    Text1.Text = ""
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And IsKeyPress(VK_CONTROL) Then
        Call SendMsg
    End If
End Sub

Private Sub Text5_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    Text5.PasswordChar = none
    If button <> 1 Then Exit Sub
    Text5.PasswordChar = "*"
End Sub

Private Sub Winsock1_Close()
    MsgBox ("�ѶϿ�������������")
    Unload Me
End Sub

Private Sub Winsock1_Connect()
    Text2.Enabled = True
    Command2.Enabled = True
    Text1.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    '����
    Dim strdata As String
    Dim strSplit
    Dim MsgType As String
    Winsock1.GetData strdata
    
    strSplit = Split(strdata, ";")
    MsgType = strSplit(0)
    
    Select Case MsgType
    Case "msg"
        Dim grpId As Integer
        Dim id As Integer
        Dim Name As String
        Dim MsgContent As String
        grpId = Int(strSplit(1))
        id = Int(strSplit(3))
        Name = strSplit(2)
        MsgContent = strSplit(4)
        MsgContent = Base64DecodeString(MsgContent)
        'Text1.Text = Name + ":" + MsgContent + vbCrLf + Text1.Text
        Call AddMessage(grpId, id, Name, MsgContent)
    
    End Select
End Sub
