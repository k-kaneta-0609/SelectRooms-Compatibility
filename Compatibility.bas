Attribute VB_Name = "compatibility"
Option Compare Database
Option Explicit

Function Compatibility_Main() As Variant

    Dim name1 As String
    name1 = InputBox("1�l�ڂ̖��O�́H", "�p���̔��p�������œ��͂��ĂˁB"): If name1 = "" Then Exit Function
    
    Dim name2 As String
    name2 = InputBox("2�l�ڂ̖��O�́H", "�p���̔��p�������œ��͂��ĂˁB"): If name1 = "" Then Exit Function
    
    ' 1�g�ڂ̕�����𐔒l��
    Dim nameNum1() As Integer
    nameNum1 = nameNumbers(name1 & name2)
    
    ' 1�g�ڂ̑���
    Dim compatibility1 As Integer
    compatibility1 = Compatibility(nameNum1)
    
    ' 2�g�ڂ̕�����𐔒l��
    Dim nameNum2() As Integer
    nameNum2 = nameNumbers(name2 & name1)
    
    ' 2�g�ڂ̑���
    Dim compatibility2 As Integer
    compatibility2 = Compatibility(nameNum2)
    
    ' �傫�����̑�����\��
    Dim msg As String
    If compatibility1 > compatibility2 Then
        msg = CStr(compatibility1)
    Else
        msg = CStr(compatibility2)
    End If
    MsgBox msg, vbOKOnly + vbInformation, "���Ȃ��B�̑����ł��B"

End Function

Function nameNumbers(name As String) As Integer()

    Dim i As Integer
    Dim length As Integer
    ReDim nameNums(Len(name) - 1) As Integer
    For i = 0 To Len(name) - 1 Step 1
        nameNums(i) = Asc(Mid(name, i + 1, 1)) - 96
    Next i
    nameNumbers = nameNums

End Function

Function Compatibility(nameNums() As Integer) As Integer

    If UBound(nameNums) = 0 Then
        Compatibility = nameNums(0)
    Else
        Dim i As Integer
        ReDim newNameNums(UBound(nameNums) - 1) As Integer
        For i = 0 To UBound(newNameNums) Step 1
            newNameNums(i) = nameNums(i) + nameNums(i + 1)
            If 101 < newNameNums(i) Then
                newNameNums(i) = newNameNums(i) - 101
            End If
        Next i
        Compatibility = Compatibility(newNameNums)
    End If

End Function

