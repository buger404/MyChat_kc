VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IPPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   该类模块是由Emerald创建的 页面控制器 模板
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim IP As String
    Dim mode As Integer
'======================================

Public Sub Update()
    '应用界面每一帧刷新会调用的方法（页面绘制i代码请放在此处）
    
    Page.Clear argb(255, 0, 0, 0) '清空画布

    Page.DrawImage "back.jpg", 0, 0, Alpha:=0.3

    If mode = 1 Then
        Page.ShowLoading GW / 2 - 150, GH / 2 - 150, 300, 300, 5, argb(255, 0, 176, 240), argb(255, 255, 51, 102), argb(255, 255, 214, 100)
        'ECore.NewTransform transToLeft, 500, "MainPage"
        Exit Sub
    End If

    Dim p(199 + 4)
    For i = 0 To 99
        p(2 * i) = i / 99 * (GW + 40)
        p(2 * i + 1) = 250 + Sin(GetTickCount / 1000 + i / 50) * 40
    Next
    p(199 + 1) = GW: p(199 + 2) = GH: p(199 + 3) = 0: p(199 + 4) = GH
    Page.PaintPolygon2 argb(120, 255, 255, 255), 0, p, True
    For i = 0 To 99
        p(2 * i) = i / 99 * (GW + 40)
        p(2 * i + 1) = 270 + Sin(GetTickCount / 600 + i / 40) * 40
    Next
    p(199 + 1) = GW: p(199 + 2) = GH: p(199 + 3) = 0: p(199 + 4) = GH
    Page.PaintPolygon2 argb(255, 255, 255, 255), 0, p, True
    
    m = CheckMouse(GW - 45, 15, 30, 30)
    Page.Paint 1, GW - 45, 15, 30, 30, argb(IIf(m, 255, 100), 180, 180, 180)
    Page.Writes "×", GW - 41, 16, 18, argb(255, 0, 0, 0), style:=FontStyleBold
    If m = mMouseUp Then Unload Client: End
    
    Page.Writes "MyChat · 分组课堂讨论", GW / 2, GH / 2 - 170, 36, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    Page.Writes "输入IP地址然后开始吧", GW / 2, GH / 2 + 70, 20, argb(255, 0, 0, 0), align:=StringAlignmentCenter
    
    Page.ShowEdit IP, 0, GW / 2 - 150, GH / 2 + 130, 300, 40, argb(255, 0, 0, 0), argb(255, 232, 232, 232), argb(255, 222, 222, 222), argb(255, 0, 176, 240), size:=20

    If UBound(Split(IP, ".")) = 3 Then
        If Page.ShowColorButton(0, GW / 2 + 80 - 70, GH - 100, 140, 40, "登陆", argb(255, 255, 255, 255), argb(220, 84, 223, 143), argb(255, 84, 223, 143), size:=20) = mMouseUp Then
            Client.Winsock1.RemoteHost = IP
            ECore.NewTransform
            mode = 1
        End If
        If Page.ShowColorButton(0, GW / 2 - 80 - 70, GH - 100, 140, 40, "注册", argb(255, 255, 255, 255), argb(220, 84, 223, 143), argb(255, 84, 223, 143), size:=20) = mMouseUp Then
            mode = 1
            ShellEx "python """ & App.path & "\" & "client.py"" -s " & IP
            Kill "sign_img.png"
            mode = 0
        End If
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    Set Page = New GPage  '创建页面绘制器
    Page.Create Me  '将绘制器与控制器绑定
    Page.Res.newImage App.path & "\assets\back.jpg", GW, GH
    Page.Res.ApplyBlurEffect "back.jpg", 20, 0
    Page.Res.NewImages App.path & "\assets"  '导入资源
    ECore.Add Page, "IPPage"  '创建页面
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
