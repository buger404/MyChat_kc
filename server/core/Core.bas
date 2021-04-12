Attribute VB_Name = "Core"
Public MusicList As GMusicList, Robots As RobotCore
Public machine() As ScriptControl
Private Declare Function GetProcessVersion Lib "kernel32" (ByVal ProcessId As Long) As Long
Public Function ShellEx(ByVal FileName As String)
    Dim i As Long, j As Long
    i = Shell(FileName, vbNormalFocus)
    Do
        If GetProcessVersion(i) = 0 Then Exit Do
        ECore.Display: DoEvents
    Loop
End Function

