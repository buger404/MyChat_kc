Attribute VB_Name = "Base64"
Function Base64EncodeString(ByVal str As String) As String
    Dim buf() As Byte
    buf = str
    Base64EncodeString = Base64Encode(buf)
End Function
Function Base64DecodeString(ByVal str As String) As String
    Dim buf() As Byte
    buf = Base64Decode(str)
    Base64DecodeString = buf
End Function
Function Base64Encode(str() As Byte) As String                                  'Base64 ����
    On Error GoTo over                                                          '�Ŵ�
    Dim buf() As Byte, Length As Long, mods As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    mods = (UBound(str) + 1) Mod 3   '����3������
    Length = UBound(str) + 1 - mods
    Dim dict() As Byte
    dict = StrConv(B64_CHAR_DICT, vbFromUnicode)
    ReDim buf(Length / 3 * 4 + IIf(mods <> 0, 4, 0) - 1)
    Dim i As Long
    For i = 0 To Length - 1 Step 3
        buf(i / 3 * 4) = (str(i) And &HFC) / &H4
        buf(i / 3 * 4 + 1) = (str(i) And &H3) * &H10 + (str(i + 1) And &HF0) / &H10
        buf(i / 3 * 4 + 2) = (str(i + 1) And &HF) * &H4 + (str(i + 2) And &HC0) / &H40
        buf(i / 3 * 4 + 3) = str(i + 2) And &H3F
    Next
    If mods = 1 Then
        buf(Length / 3 * 4) = (str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (str(Length) And &H3) * &H10
        buf(Length / 3 * 4 + 2) = 64
        buf(Length / 3 * 4 + 3) = 64
    ElseIf mods = 2 Then
        buf(Length / 3 * 4) = (str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (str(Length) And &H3) * &H10 + (str(Length + 1) And &HF0) / &H10
        buf(Length / 3 * 4 + 2) = (str(Length + 1) And &HF) * &H4
        buf(Length / 3 * 4 + 3) = 64
    End If
    For i = 0 To UBound(buf)
        buf(i) = dict(buf(i))
        'Base64Encode = Base64Encode + Mid(B64_CHAR_DICT, buf(i) + 1, 1)
    Next
    Base64Encode = StrConv(buf, vbUnicode)
over:
End Function
Function Base64Decode(B64 As String) As Byte()                                  'Base64 ����
    On Error GoTo over                                                          '�Ŵ�
    Dim OutStr() As Byte, i As Long, j As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    If InStr(1, B64, "=") <> 0 Then B64 = Left(B64, InStr(1, B64, "=") - 1)     '�ж�Base64��ʵ����,��ȥ��λ
    Dim Length As Long, mods As Long
    mods = Len(B64) Mod 4
    Length = Len(B64) - mods
    Dim dict() As Byte
    dict = StrConv(B64_CHAR_DICT, vbFromUnicode)
    Dim wo() As Byte
    wo = StrConv(B64, vbFromUnicode)
    ReDim OutStr(Length / 4 * 3 - 1 + Switch(mods = 0, 0, mods = 2, 1, mods = 3, 2))
    For i = 1 To Length Step 4
        Dim buf(3) As Byte
        For j = 0 To 3
            buf(j) = InStr(1, B64_CHAR_DICT, Chr(wo(i + j - 1))) - 1
            'buf(j) = InStr(1, B64_CHAR_DICT, Mid(B64, i + j, 1)) - 1            '�����ַ���λ��ȡ������ֵ
        Next
        OutStr((i - 1) / 4 * 3) = buf(0) * &H4 + (buf(1) And &H30) / &H10
        OutStr((i - 1) / 4 * 3 + 1) = (buf(1) And &HF) * &H10 + (buf(2) And &H3C) / &H4
        OutStr((i - 1) / 4 * 3 + 2) = (buf(2) And &H3) * &H40 + buf(3)
    Next
    If mods = 2 Then
        OutStr(Length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Chr(wo(Length))) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Chr(wo(Length + 1))) - 1) And &H30) / 16
    ElseIf mods = 3 Then
        OutStr(Length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Chr(wo(Length))) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Chr(wo(Length + 1))) - 1) And &H30) / 16
        OutStr(Length / 4 * 3 + 1) = ((InStr(1, B64_CHAR_DICT, Chr(wo(Length + 1))) - 1) And &HF) * &H10 + ((InStr(1, B64_CHAR_DICT, Chr(wo(Length + 2))) - 1) And &H3C) / &H4
    End If
    Base64Decode = OutStr                                                       '��ȡ������
over:
End Function
