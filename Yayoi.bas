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

'�e�s�̏���: �Z�����當����ɐ؂�o��
Function lineout(ByRef r As Range) As String

  Dim s As String
  Dim c As Range
  Dim field_name As String, field_type As String
  Dim record_value As String
    
  Set c = r
  Do While Not IsEmpty(Range("A1").Offset(0, c.Column - 1))
    
    '�Z������f�[�^���肾��
    field_name = Range("A1").Offset(0, c.Column - 1)
    field_type = Range("A2").Offset(0, c.Column - 1)
    record_value = c.Value
    
    '�t�B�[���h�^�C�v���Ƃ̕␳
    If field_type = "����" Then
      '����: ���̂܂܏o��
      v = record_value
    ElseIf field_type = "���z" Then
      '���z: �󗓂̎���0�o��
      If record_value = "" Then
        v = 0
      Else
        v = record_value
      End If
    Else
      '���̑�: �_�u���N�H�[�e�[�V����
      v = """" & record_value & """"
    End If
    
    Debug.Print r.Address, c.Address, field_name, field_type, v
    
    '�o�͕�����ɓ����
    If s = "" Then
        '1���ڂ�
        s = v
    Else
        '2���ڂ߈ȍ~�͋�؂蕶��������
        s = s & "," & v
    End If
    
    Set c = c.Offset(0, 1)
  Loop
    
  lineout = s

End Function



