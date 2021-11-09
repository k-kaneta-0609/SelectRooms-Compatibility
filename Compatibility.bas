Attribute VB_Name = "compatibility"
Option Compare Database
Option Explicit

Function Compatibility_Main() As Variant

    Dim name1 As String
    name1 = InputBox("1人目の名前は？", "英字の半角小文字で入力してね。"): If name1 = "" Then Exit Function
    
    Dim name2 As String
    name2 = InputBox("2人目の名前は？", "英字の半角小文字で入力してね。"): If name1 = "" Then Exit Function
    
    ' 1組目の文字列を数値化
    Dim nameNum1() As Integer
    nameNum1 = nameNumbers(name1 & name2)
    
    ' 1組目の相性
    Dim compatibility1 As Integer
    compatibility1 = Compatibility(nameNum1)
    
    ' 2組目の文字列を数値化
    Dim nameNum2() As Integer
    nameNum2 = nameNumbers(name2 & name1)
    
    ' 2組目の相性
    Dim compatibility2 As Integer
    compatibility2 = Compatibility(nameNum2)
    
    ' 大きい方の相性を表示
    Dim msg As String
    If compatibility1 > compatibility2 Then
        msg = CStr(compatibility1)
    Else
        msg = CStr(compatibility2)
    End If
    MsgBox msg, vbOKOnly + vbInformation, "あなた達の相性です。"

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

