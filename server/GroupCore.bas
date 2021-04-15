Attribute VB_Name = "GroupCore"
Public Type Messages
    id As Integer
    Name As String
    Content As String
    time As Date
End Type
Public Type Member
    id As Integer
    Name As String
End Type
Public Type group
    id As Integer
    leader As Integer
    isJoin As Boolean
    Name As String
    msg() As Messages
    unreadTick As Integer
    members() As Member
    LeaderName As String
End Type
Public Type MsgBan
    id As Integer
    groupid As Integer
    StartTime As Date
    Duration As Long
End Type
Public Type Dump
    groups() As group
    bans() As MsgBan
End Type
Public userId As Integer, userName As String, realSize As Long
Public MainPage As MainPage, selectMsg As Messages
Public groups() As group, bans() As MsgBan
Public Sub DumpFile()
    Dim Dump As Dump

End Sub
Public Sub AddGroup(id As Integer, leader As Integer, isJoin As Boolean, Name As String, LeaderName As String)
    ReDim Preserve groups(UBound(groups) + 1)
    With groups(UBound(groups))
        .id = id
        .isJoin = isJoin
        .Name = Name
        .leader = leader
        .LeaderName = LeaderName
        ReDim .msg(0)
        ReDim .members(0)
    End With
    Call DumpFile
End Sub
Public Sub DeleteGroup(id As Integer)
    For i = 1 To UBound(groups)
        If groups(i).id = id Then
            For j = i To UBound(groups) - 1
                groups(j) = groups(j + 1)
            Next
            Exit For
        End If
    Next
    On Error GoTo sth
    ReDim Preserve groups(UBound(groups) - 1)
sth:
    If Err.Number = 0 Then Exit Sub
    realSize = UBound(groups) - 1
    Call DumpFile
End Sub
Public Sub AddMessage(id As Integer, memberid As Integer, Name As String, ByVal Content As String)
    For i = 1 To UBound(groups)
        If groups(i).id = id Then
            ReDim Preserve groups(i).msg(UBound(groups(i).msg) + 1)
            With groups(i).msg(UBound(groups(i).msg))
                .Content = Content
                .id = memberid
                .Name = Name
                .time = Now
            End With
            groups(i).unreadTick = groups(i).unreadTick + 1
            If groups(i).unreadTick > 100 Then groups(i).unreadTick = 100
            Exit For
        End If
    Next
    
    If memberid <> -4 And (Not Robots Is Nothing) Then
        For i = 1 To UBound(machine)
            Robots.currentRobot = MenuWindow.robotBtn(i).Caption
            On Error GoTo sth
            machine(i).Run "Process", Content, id, memberid
sth:
            If Err.Number <> 0 Then
                Robots.SendMessage id, "机器人“" & Robots.currentRobot & "”发生问题，请联系机器人制作者：" & vbCrLf & "第" & machine(i).Error.Line & "行：" & machine(i).Error.Description & vbCrLf & "错误码：" & machine(i).Error.Number
                Err.Clear
            End If
        Next
    End If
End Sub
Public Sub SetJoinState(id As Integer, isJoin As Boolean)
    '仅客户端使用！！！服务端忽视加入状态！
    For i = 1 To UBound(groups)
        If groups(i).id = id Then
            groups(i).isJoin = isJoin
            Exit For
        End If
    Next
End Sub
Public Sub AddMember(Name As String, id As Integer, group As Integer)
    For i = 1 To UBound(groups)
        If groups(i).id = group Then
            ReDim Preserve groups(i).members(UBound(groups(i).members) + 1)
            With groups(i).members(UBound(groups(i).members))
                .Name = Name
                .id = id
            End With
            Exit For
        End If
    Next
    Call DumpFile
End Sub
Public Sub DeleteMember(id As Integer, group As Integer)
    For i = 1 To UBound(groups)
        If groups(i).id = group Then
            For j = 1 To UBound(groups(i).members)
                If groups(i).members(j).id = id Then
                    groups(i).members(j) = groups(i).members(UBound(groups(i).members))
                    ReDim Preserve groups(i).members(UBound(groups(i).members) - 1)
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    Call DumpFile
End Sub
Public Sub AddBan(id As Integer, group As Integer, Duration As Long)
    If Duration = 0 Then
        For i = 1 To UBound(bans)
            If bans(i).id = id And bans(i).groupid = group Then
                bans(i) = bans(UBound(bans))
                ReDim Preserve bans(UBound(bans) - 1)
                Exit For
            End If
        Next
    Else
        ReDim Preserve bans(UBound(bans) + 1)
        With bans(UBound(bans))
            .id = id
            .groupid = group
            .StartTime = Now
            .Duration = Duration
        End With
        Call DumpFile
    End If
End Sub
Public Function IsBan(id As Integer, group As Integer)
    IsBan = False
    For i = 1 To UBound(bans)
        If bans(i).id = id And bans(i).groupid = group Then
            IsBan = True
            Exit For
        End If
    Next
End Function
