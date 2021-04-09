Attribute VB_Name = "Core"
Public MusicList As GMusicList
Private Declare Function GetProcessVersion Lib "kernel32" (ByVal ProcessId As Long) As Long
Public Function ShellEx(ByVal FileName As String)
    Dim I As Long, J As Long
    I = Shell(FileName, vbNormalFocus)
    Do
        If GetProcessVersion(I) = 0 Then Exit Do
        ECore.Display: DoEvents
    Loop
End Function

