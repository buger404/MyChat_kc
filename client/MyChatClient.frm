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
   StartUpPosition =   2  '??Ļ????
   Begin VB.Timer BlockTimer 
      Interval        =   100
      Left            =   240
      Top             =   1560
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
         Name            =   "΢???ź?"
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
         Name            =   "΢???ź?"
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
         Name            =   "΢???ź?"
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
         Name            =   "΢???ź?"
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
         Name            =   "΢???ź?"
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
      Caption         =   "ͼƬת????"
      Height          =   495
      Left            =   7680
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Audio 
      Caption         =   "????ʶ??"
      Height          =   495
      Left            =   7680
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton save3 
      Caption         =   "????"
      Height          =   495
      Left            =   6552
      TabIndex        =   13
      ToolTipText     =   "??????D??"
      Top             =   2340
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton save2 
      Caption         =   "????"
      Height          =   495
      Left            =   6552
      TabIndex        =   12
      ToolTipText     =   "??????D??"
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
      ToolTipText     =   "????????HARDBASSʹ??Ŷ??"
      Top             =   2325
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "????"
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      ToolTipText     =   "?ٺ٣???ˮ??ʼ??"
      Top             =   1950
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "???±?"
      Height          =   375
      Left            =   7680
      TabIndex        =   7
      ToolTipText     =   "???ԼǱʼ??ޣ?"
      Top             =   1575
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "????"
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      ToolTipText     =   "????ʲô?ɣ?"
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "????"
      Height          =   375
      Left            =   6396
      TabIndex        =   5
      Top             =   4056
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "????"
      Height          =   615
      Left            =   5304
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "????"
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
   Begin MSComDlg.CommonDialog trans 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "?κ??ļ?|*.*"
   End
   Begin MSComDlg.CommonDialog imgOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "ͼ???ļ?|*.jpg;*.png;*.bmp;*.gif;*.jpeg"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "ͼ???ļ?|*.jpg;*.png;*.bmp;*.gif;*.jpeg"
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim grpId As Integer
Dim localId As Integer
Dim i As Single, p As Single, q As Single, dr As Single
Dim IPPage As IPPage
Public LBtnColor As Long, RBtnColor As Long
Public DotMode As Boolean
Dim DrawX As Single, DrawY As Single
Dim Shadow As aShadow
Dim lHwnd As Long, lText As String
Dim buffData As String
Dim SafeToQuit As Boolean
Public Sub createGrp()
    Dim id As Integer, leader As Integer, isJoin As Boolean, Name As String
    'id = InputBox("id")
    'leader = InputBox("leader")
    'isJoin = True
reinput:
    Name = InputBox("??????Ҫ??????????????????")
    If Name = "" Then MsgBox "???????????ֲ???Ϊ?գ?", 48: GoTo reinput
    'Call AddGroup(Int(id), Int(leader), isJoin, name)
    
End Sub
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

Private Sub BlockTimer_Timer()
    Dim hwnd As Long, Title As String * 255, Class As String * 255, imName As String * 255, imName2 As String, imN() As String
    hwnd = GetForegroundWindow
    GetWindowTextA hwnd, Title, 255
    GetClassNameA hwnd, Class, 255
    imName = GetProcessPath(hwnd)
    imN = Split(imName, "\")
    imName2 = LCase(imN(UBound(imN)))
    Title = LCase(Title): Class = LCase(Class)
    Dim tit As String, cl As String, im As String
    tit = UnSpace(Title): cl = UnSpace(Class): im = Replace(imName2, Chr(32), "")
    If hwnd = lHwnd And tit = lText Then Exit Sub
    lText = tit: lHwnd = hwnd
    For i = 1 To UBound(nbf.item)
        If (tit Like nbf.item(i).Title And cl Like nbf.item(i).Class And im Like nbf.item(i).image) Or (tit = nbf.item(i).Title And cl = nbf.item(i).Class And im = nbf.item(i).image) Then
            CloseWindow hwnd
            EnableWindow hwnd, 0
            ShowWindow hwnd, SW_HIDE
            DestroyWindow hwnd
            Winsock1.SendData "msg;1;" & Base64EncodeString("???ں?????") & ";-1;" & Base64EncodeString(userName & "??????Υ?????ڣ?????ֹ??") & vbCrLf
            AddMessage 1, -1, "???ں?????", userName & "??????Υ?????ڣ?????ֹ??"
            DoEvents
            Exit For
        End If
    Next
End Sub

Private Sub Command2_Click()
    '????
    Call SendMsg
End Sub

Public Sub Command4_Click()
    Picture1.Cls
End Sub



Private Sub logIn()
    If Dir("id_info.txt") <> "" Then Kill "id_info.txt"
    If Dir("face.png") <> "" Then Kill "face.png"
    If Dir("id_info.txt") = "" Then ShellEx "python """ & "client.py"" -l " & Winsock1.RemoteHost
    If Dir("id_info.txt") = "" Then MsgBox "???޴??ˣ?????ע?᣿", 16, "??½ʧ??": End
    
    Open "id_info.txt" For Input As 1
    A = StrConv(InputB(FileLen("id_info.txt"), 1), vbUnicode)
    S = Split(A, ",")
    Close #1
    If S(0) = "404" Then MsgBox "ip??ַ????????????ip??ַ", 16, "ip??ַ????": End
    userName = S(0)
    Me.Caption = userName

    If Dir("id_info.txt") <> "" Then Kill "id_info.txt"
    If Dir("face.png") <> "" Then Kill "face.png"
    
    Winsock1.RemotePort = 2001
    If Winsock1.State = sckClosed Then
        Winsock1.Connect
    End If

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
'Emerald???ܲ???
Private Sub InitEmeraldFramework()
    '????Emerald
    StartEmerald Me.hwnd, 1100, 600, False
    '??????????Ⱦ
    Set EF = New GFont
    EF.MakeFont "΢???ź?"
    'ʵ????ҳ????????????
    Set ECore = New GMan
    'ʵ????ҳ????????
    Set MainPage = New MainPage
    Set IPPage = New IPPage
    '??ʾ
    DrawTimer.Enabled = True
    ECore.ActivePage = "IPPage"
End Sub
Private Sub UnloadEmeraldFramework()
    DrawTimer.Enabled = False
    EndEmerald
End Sub
Private Sub DrawTimer_Timer()
    '???»???
    ECore.Display
    DoEvents
End Sub
Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    '??????????Ϣ
    UpdateMouse x, y, 1, button
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    '??????????Ϣ
    If Mouse.State = 0 Then
        UpdateMouse x, y, 0, button
    Else
        Mouse.x = x: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    '??????????Ϣ
    UpdateMouse x, y, 2, button
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    '?????ַ?????
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub
'===============================================================================================================

Private Sub Form_Load()
    'Base64_Init
    LoadBlackList
    ReDim groups(0): ReDim bans(0)
    
    Call InitEmeraldFramework
    Set Shadow = New aShadow
    With Shadow
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 8
            .Transparency = 80
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
    
    '??׼????????
    Text5.Move 300, 60, Me.ScaleWidth - 300, Me.ScaleHeight - 60 - 120
    Text1.Move Text5.Left, Text5.Top, Text5.Width, Text5.Height
    Picture1.Move Text5.Left, Text5.Top, Text5.Width, Text5.Height
    Picture2.Move Text5.Left, Text5.Top, Text5.Width, Text5.Height
    Text2.Move 300 + 50, Me.ScaleHeight - 80 + 25, Me.ScaleWidth - 245 - 300, 80 - 50
    
    Text1.Visible = True
    Text1.Move -90, -90, 1, 1
    Text5.Visible = False
    Picture1.Visible = False
    Picture2.Visible = False
    
    Command4.Visible = False
    
    ECore.NewTransform , , "MainPage"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Shadow = Nothing
    On Error Resume Next
    Me.Hide
    For i = 2 To UBound(groups)
        Winsock1.SendData "quitgroup;" & groups(i).id & vbCrLf
        DoEvents
    Next
    Winsock1.SendData "quitrequire" & vbCrLf
    DoEvents
    Do While Not SafeToQuit
        DoEvents
    Loop
    'Client.save3_Click
    Call UnloadEmeraldFramework
    End
End Sub

Public Sub Option1_Click()
    save2.Visible = False
    save3.Visible = False
    
    Command4.Visible = False
    
    Picture1.Visible = True
    Text1.Visible = False
    Text5.Visible = False
    Picture2.Visible = False
End Sub
Public Sub Option2_Click()
    save2.Visible = False
    save3.Visible = False
    
    Command4.Visible = False
    
    Picture1.Visible = False
    Text1.Visible = False
    Text5.Visible = True
    Picture2.Visible = False
End Sub
Public Sub Option3_Click()
    save2.Visible = False
    save3.Visible = False
    
    Command4.Visible = False
    
    Picture1.Visible = False
    Text1.Visible = True
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
    Open App.path & "\" & userName & "?ıʼ?" & str(p) & ".txt" For Output As #1
    Print #1, Text5.Text
    Close #1
    MsgBox "?ѳɹ??????ʼǣ?" & vbCrLf & App.path & "\" & userName & "?ıʼ?" & str(p) & ".txt", 64, "?????ɹ?"
    p = p + 1
End Sub
Public Sub save3_Click()
    Open App.path & "\" & userName & "????Ϣ??¼" & str(q) & ".txt" For Output As #1
    Print #1, Text1.Text
    Close #1
    MsgBox "?ѳɹ???????Ϣ??¼??" & vbCrLf & App.path & "\" & userName & "????Ϣ??¼" & str(q) & ".txt", 64, "?????ɹ?"
    q = q + 1
End Sub
Public Sub saveDrawing()
    SavePicture Client.Picture1.image, App.path & "\" & userName & "??Ϳѻ" & str(dr) & ".bmp"
    MsgBox "?ѳɹ?????Ϳѻ??" & vbCrLf & App.path & "\" & userName & "??Ϳѻ" & str(dr) & ".bmp", 64, "?????ɹ?"
    dr = dr + 1
End Sub

Private Sub Text2_Change()
    If Picture2.Visible = True Then
        Picture2.BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    End If

    If i >= 15 Then i = 1
    i = i + 1
    
    '?Զ??????ı?????С
    Dim Line As Long
    Line = UBound(Split(Text2.Text, vbCrLf)) + 1
    If Line <= 0 Then Line = 1
    Dim Border As Integer, Height As Long
    Border = IIf(Line > 1, 1, 0)
    Height = Line * 30
    If Height > Me.ScaleHeight - 120 Then Height = Me.ScaleHeight - 120 '??ֹ??????????
    If Text2.BorderStyle <> Border Then Text2.BorderStyle = Border
    If Text2.Height <> Height Then
        Text2.Height = Height
        Text2.Top = Me.ScaleHeight - 80 + 25 - Height + 30
    End If
End Sub
Public Sub SendMsg(Optional Msg As String = "")
    If Winsock1.State <> 7 Then Exit Sub
    If Text2.Text = "" And Msg = "" Then
        VBA.Beep
    Else
        Dim txt As String
        If Msg = "" Then
            txt = Base64EncodeString(Text2.Text)
            Call AddMessage(groups(MainPage.selectIndex).id, userId, "??", Text2.Text)
        Else
            txt = Base64EncodeString(Msg)
        End If
        Winsock1.SendData "msg;" + str(groups(MainPage.selectIndex).id) + ";" + Base64EncodeString(userName) + ";" + str(userId) + ";" + txt & vbCrLf
        'Winsock1.SendData Text2.Text
        Text2.Text = ""
    End If
End Sub

Public Sub fileServer()
    MsgBox "?ļ????????ѿ???..."
    Dim S As String
    S = "python -m http.server 8080 -d \share -b " + Client.Winsock1.LocalIP
    MsgBox S
    
    Shell S, vbMinimizedNoFocus
End Sub

Public Sub getId()
    If Winsock1.State <> 7 Then Exit Sub
    Winsock1.SendData "getId;" + vbCrLf
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
    MsgBox ("?ѶϿ?????????????")
    Unload Me
    End
End Sub

Private Sub Winsock1_Connect()
    Text2.Enabled = True
    Command2.Enabled = True
    Text1.Enabled = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    '????
    Dim strdata As String
    Dim strSplit() As String
    Dim MsgType As String
    Dim grpId As Integer
    Dim id As Integer
    Dim Name As String
    Dim MsgContent As String
    Dim leaderId As Integer
    Dim grpName As String
    Dim LeaderName As String
    Winsock1.GetData strdata
    If Right(strdata, 2) <> Chr(13) & Chr(10) Then
        buffData = buffData & strdata
        Exit Sub
    Else
        strdata = buffData & strdata
        buffData = ""
    End If
    
    Dim cmds() As String
    cmds = Split(strdata, vbCrLf)
    
    For k = 0 To UBound(cmds) - 1
        strSplit = Split(cmds(k), ";")
        MsgType = strSplit(0)

        Select Case MsgType
        Case "filerecv"
            ShellExecuteA 0, "open", App.path & "\FileTransportation.exe", "-d;" & strSplit(1) & ";" & strSplit(2) & ";" & strSplit(3) & ";" & strSplit(4) & ";" & strSplit(5), "", SW_SHOW
        Case "black"
            Dim pur As String
            For i = 1 To UBound(strSplit) - 1
                pur = pur & strSplit(i) & ";"
            Next
            BlackPurse pur
        Case "getId"
            localId = strSplit(1)
        Case "msg"
            grpId = Int(strSplit(1))
            id = Int(strSplit(3))
            Name = Base64DecodeString(strSplit(2))
            MsgContent = strSplit(4)
            MsgContent = Base64DecodeString(MsgContent)
            'Text1.Text = Name + ":" + MsgContent + vbCrLf + Text1.Text
            If InStr(MsgContent, "_image;") = 1 Then
                Dim imgdata() As Byte, tt() As String
                tt = Split(MsgContent, ";")
                imgdata = Base64Decode(tt(2))
                Open App.path & "\imgrecv\" & tt(1) For Binary As #1
                Put #1, , imgdata
                Close #1
                Call AddMessage(grpId, id, Name, "_image;" & tt(1))
                MainPage.Page.Res.newImage App.path & "\imgrecv\" & tt(1), arg2:=200
            Else
                Call AddMessage(grpId, id, Name, MsgContent)
            End If
            
        Case "newgroup"
            'newgroup;groupid;groupname(base64);LeaderName(base64);leaderid
            Call AddGroup(Int(strSplit(1)), Int(strSplit(4)), False, Base64DecodeString(strSplit(2)), Base64DecodeString(strSplit(3)))
        Case "identify"
            userId = Val(strSplit(1))
            Winsock1.SendData "addgrouprequest;" & Base64EncodeString(userName) & ";" & userId & ";" & 1 & vbCrLf
        Case "grouprequest"
            Me.SetFocus
            Dim gidi As Integer
            For i = 1 To UBound(groups)
                If groups(i).id = Val(strSplit(3)) Then
                    gidi = i: Exit For
                End If
            Next
            If MsgBox(Base64DecodeString(strSplit(1)) & "(#" & Val(strSplit(2)) & ") ?????????顰" & groups(gidi).Name & "?????Ƿ?ͬ?⣿", 48 + vbYesNo) = vbYes Then
                AddMember Base64DecodeString(strSplit(1)), Val(strSplit(2)), Val(strSplit(3))
                Winsock1.SendData "broadcast;addmember;" & strSplit(1) & ";" & strSplit(2) & ";" & strSplit(3) & vbCrLf
            End If
        Case "addmember"
            If Val(strSplit(2)) = userId Then SetJoinState Val(strSplit(3)), True
            AddMember Base64DecodeString(strSplit(1)), Val(strSplit(2)), Val(strSplit(3))
            AddMessage Val(strSplit(3)), -1, "ϵͳ??Ϣ", Base64DecodeString(strSplit(1)) & "?????˱???????"
        Case "addgroup"
            Call AddGroup(Val(strSplit(4)), Val(strSplit(1)), Val(strSplit(1)) = userId, Base64DecodeString(strSplit(2)), Base64DecodeString(strSplit(3)))
        Case "deletegroup"
            DeleteGroup Val(strSplit(1))
        Case "deletemember"
            DeleteMember Val(strSplit(2)), Val(strSplit(1))
            If Val(strSplit(2)) = userId Then SetJoinState Val(strSplit(1)), False
        Case "addban"
            AddBan Val(strSplit(3)), Val(strSplit(1)), Val(strSplit(2))
        Case "safequit"
            SafeToQuit = True
        End Select
    Next

End Sub
