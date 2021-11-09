VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SelectRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Main_Click()

    Dim i As Integer
    
    ' 画面クリア
    inList.Value = ""
    outList.Value = ""
    
    ' 1行目入力
    Dim xNum As String
    xNum = InputBox("嫌いな番号は？", "質問①"): If xNum = "" Then Exit Sub
    Call InListWriteLine(xNum)
    
    ' 2行目入力
    Dim count As String
    count = InputBox("何部屋あるの？", "質問②"): If count = "" Then Exit Sub
    Dim roomCount As Integer
    roomCount = CInt(count)
    Call InListWriteLine(CStr(roomCount))
    
    ' 3行目以降入力
    Dim roomNums() As String
    For i = 1 To roomCount Step 1
        ReDim Preserve roomNums(i - 1)
        roomNums(i - 1) = InputBox("部屋番号は？", i & "部屋目"): If roomNums(i - 1) = "" Then Exit Sub
        Call InListWriteLine(roomNums(i - 1))
    Next i
    
    ' 嫌いな番号が含まれていない部屋番号を出力
    Dim exists As Boolean
    exists = False
    For i = LBound(roomNums) To UBound(roomNums) Step 1
        If 0 = InStr(1, roomNums(i), xNum, vbTextCompare) Then
            exists = True
            Call OutListWriteLine(roomNums(i))
        End If
    Next i
    If exists = False Then
        Call OutListWriteLine("none")
    End If

End Sub

Private Sub InListWriteLine(str)

    inList.Value = inList.Value & str & vbNewLine

End Sub

Private Sub OutListWriteLine(str)

    outList.Value = outList.Value & str & vbNewLine

End Sub
