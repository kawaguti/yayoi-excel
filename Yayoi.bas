Attribute VB_Name = "Yayoi"
Sub output2yayoi()

  Dim r As Range, c As Range
  Dim s As String
  Open ThisWorkbook.Path & "\" & "yayoi_import.txt" For Output As #1

  Set r = Range("A3")
  Do While Not IsEmpty(r)
    s = lineout(r)
    'Debug.Print s
    Print #1, s
    Set r = r.Offset(1, 0)
  Loop
    
  Close #1


End Sub

'各行の処理: セルから文字列に切り出す
Function lineout(ByRef r As Range) As String

  Dim s As String
  Dim c As Range
  Dim field_name As String, field_type As String
  Dim record_value As String
    
  Set c = r
  Do While Not IsEmpty(Range("A1").Offset(0, c.Column - 1))
    
    'セルからデータきりだし
    field_name = Range("A1").Offset(0, c.Column - 1)
    field_type = Range("A2").Offset(0, c.Column - 1)
    record_value = c.Value
    
    'フィールドタイプごとの補正
    If field_type = "数字" Then
      '数字: そのまま出力
      v = record_value
    ElseIf field_type = "金額" Then
      '金額: 空欄の時は0出力
      If record_value = "" Then
        v = 0
      Else
        v = record_value
      End If
    Else
      'その他: ダブルクォーテーション
      v = """" & record_value & """"
    End If
    
    Debug.Print r.Address, c.Address, field_name, field_type, v
    
    '出力文字列に入れる
    If s = "" Then
        '1項目め
        s = v
    Else
        '2項目め以降は区切り文字をつける
        s = s & "," & v
    End If
    
    Set c = c.Offset(0, 1)
  Loop
    
  lineout = s

End Function



