Attribute VB_Name = "GroupCore"
Public Type Messages
    id As Integer
    name As String
    Content As String
    time As Date
End Type
Public Type Group
    id As Integer
    leader As Integer
    isJoin As Boolean
    name As String
<<<<<<< HEAD
    msg() As Messages
=======
    Msg() As Messages
>>>>>>> 3b07fd90bb91919c2b4047f89cff85163e999376
End Type
Public userId As Integer
Public groups() As Group

Public Sub AddGroup(id As Integer, leader As Integer, isJoin As Boolean, name As String)
    ReDim Preserve groups(UBound(groups) + 1)
    With groups(UBound(groups))
        .id = id
        .isJoin = isJoin
        .name = name
        .leader = leader
        ReDim .Msg(0)
    End With
End Sub
Public Sub DeleteGroup(id As Integer)
    For i = 1 To UBound(groups)
        If groups(i).id = id Then
            For j = i To UBound(groups) - 1
                groups(j) = groups(j + 1)
            Next
            ReDim Preserve groups(UBound(groups) - 1)
            Exit For
        End If
    Next
End Sub
Public Sub AddMessage(id As Integer, memberid As Integer, name As String, Content As String)
    For i = 1 To UBound(groups)
        If groups(i).id = id Then
            ReDim Preserve groups(i).Msg(UBound(groups(i).Msg) + 1)
            With groups(i).Msg(UBound(groups(i).Msg))
                .Content = Content
                .id = id
                .name = name
                .time = Now
            End With
            Exit For
        End If
    Next
End Sub
Public Sub SetJoinState(id As Integer, isJoin As Boolean)
    For i = 1 To UBound(groups)
        If groups(i).id = id Then
            groups(i).isJoin = isJoin
            Exit For
        End If
    Next
End Sub

