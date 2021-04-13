Attribute VB_Name = "BlackList"
Public Type BlackListItem
    Title As String
    Class As String
    image As String
End Type
Public Type BlackFile
    item() As BlackListItem
End Type
Public bf As BlackFile, nbf As BlackFile
Public uwpSwitch As Long, uwpRet As Long
Public Function uwpChild(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    If uwpRet <> 0 Then GoTo last

    Dim Class As String * 255
    GetClassNameA hwnd, Class, 255
    
    If UnSpace(Class) = "Windows.UI.Core.CoreWindow" Then uwpRet = hwnd
    
last:
    uwpChild = True
End Function
Public Function uwpFind(ByVal hwnd As Long) As Long

    uwpRet = 0
    EnumChildWindows hwnd, AddressOf uwpChild, 0&
    uwpFind = uwpRet
    
End Function
Public Function GetProcessPath(hwnd As Long) As String
    On Error GoTo z
    
recheck:
    
    Dim PID As Long, Class As String * 255
    Dim cbNeeded As Long, szBuf(1 To 250) As Long, ret As Long, szPathName As String, nSize As Long, hProcess As Long
    
    Class = "": PID = 0
    
    GetWindowThreadProcessId hwnd, PID
    GetClassNameA hwnd, Class, 255
    
    If UnSpace(Class) = "ApplicationFrameWindow" And hwnd <> 0 Then 'UWP
        hwnd = uwpFind(hwnd)
        GoTo recheck
    End If
    
    hProcess = OpenProcess(&H400 Or &H10, 0, PID)
    If hProcess <> 0 Then
        szPathName = Space(260): nSize = 500
        ret = GetModuleFileNameExA(hProcess, szBuf(1), szPathName, nSize)
        GetProcessPath = Left(szPathName, ret)
    End If
    
    ret = CloseHandle(hProcess)
    If GetProcessPath = "" Then
        GetProcessPath = "System"
    End If
    
    Exit Function
z:
End Function
Public Sub LoadBlackList()
    ReDim bf.item(0)
    ReDim nbf.item(0)
    If Dir(App.path & "\blacklist.bin") <> "" Then
        Open App.path & "\blacklist.bin" For Binary As #1
        Get #1, , nbf
        Close #1
        bf = nbf
    End If
End Sub
Public Sub Dump()
    bf = nbf
    Open App.path & "\blacklist.bin" For Binary As #1
    Put #1, , bf
    Close #1
End Sub
Public Function GetBlackString() As String
    Dim ret As String
    For i = 1 To UBound(nbf.item)
        With nbf.item(i)
            ret = ret & Base64EncodeString(.Title) & ";" & Base64EncodeString(.Class) & ";" & Base64EncodeString(.image) & ";"
        End With
    Next
    GetBlackString = ret
End Function
Public Function UnSpace(str As String) As String
    MenuWindow.Formater.Caption = str
    UnSpace = MenuWindow.Formater.Caption
End Function
Public Sub AddBlack(Title As String, Class As String, image As String)
    ReDim Preserve nbf.item(UBound(nbf.item) + 1)
    With nbf.item(UBound(nbf.item))
        .Title = UnSpace(Title)
        .Class = UnSpace(Class)
        .image = UnSpace(image)
    End With
End Sub
Public Function BlackPurse(pur As String)
    Dim t() As String
    ReDim bf.item(0)
    t = Split(pur, ";")
    For i = 0 To UBound(t) Step 3
        If t(i) = "" Then Exit For
        AddBlack Base64DecodeString(t(i)), Base64DecodeString(t(i + 1)), Base64DecodeString(t(i + 2))
    Next
End Function
