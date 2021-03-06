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
    Msg() As Messages
    unreadTick As Integer
    members() As Member
    LeaderName As String
End Type
Public Type MsgBan
    id As Integer
    groupId As Integer
    StartTime As Date
    Duration As Long
End Type
Public Type Dump
    groups() As group
    bans() As MsgBan
End Type
Public userId As Integer, userName As String
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
        ReDim .Msg(0)
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
    ReDim Preserve groups(UBound(groups) - 1)
    Call DumpFile
End Sub
Public Sub AddMessage(id As Integer, memberid As Integer, Name As String, Content As String)
    For i = 1 To UBound(groups)
        If groups(i).id = id Then
            ReDim Preserve groups(i).Msg(UBound(groups(i).Msg) + 1)
            With groups(i).Msg(UBound(groups(i).Msg))
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
End Sub
Public Sub SetJoinState(id As Integer, isJoin As Boolean)
    '??????????????????????????????????????
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
            If bans(i).id = id And bans(i).groupId = group Then
                bans(i) = bans(UBound(bans))
                ReDim Preserve bans(UBound(bans) - 1)
                Exit For
            End If
        Next
    Else
        ReDim Preserve bans(UBound(bans) + 1)
        With bans(UBound(bans))
            .id = id
            .groupId = group
            .StartTime = Now
            .Duration = Duration
        End With
        Call DumpFile
    End If
End Sub
Public Function IsBan(id As Integer, group As Integer)
    IsBan = False
    For i = 1 To UBound(bans)
        If bans(i).id = id And bans(i).groupId = group Then
            IsBan = True
            Exit For
        End If
    Next
End Function

