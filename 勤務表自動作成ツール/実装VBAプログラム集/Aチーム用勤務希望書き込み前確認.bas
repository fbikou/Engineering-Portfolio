Attribute VB_Name = "Module1"
Option Explicit
Sub A�`�[���p�Ζ���]�������ݑO�m�F()
  Dim ws As Worksheet
  Dim number_of_people As Integer
  Dim check_people_num As Integer 'number_of_people�̃��[�v�p�̕ϐ�
  Dim check_cell_num As Integer
  Dim check_sub_title(1) As String
  check_sub_title(0) = "�l���ݒ�"
  check_sub_title(1) = "�A���񐔐ݒ�"
  Dim check_sub_title_num As Integer
  Dim check_title_for_three_change(1) As String
  Dim check_title_num_for_three_change As Integer
  
  On Error GoTo ErrLabel
  
'�p�����[�^�ݒ�̊����m�F

  '�ǂݍ��ݗp�A�N�e�B�u�V�[�g�w��
  Set ws = ThisWorkbook.Sheets("A�`�[���p�Ζ���]�\")
  'name,number_of_people_reading
  For check_people_num = 19 To 19 + 2 * (30 - 1) Step 2 'For i=19 To 19+2*(30(�l)-1) Step 2
    If ws.Cells(check_people_num, 5) <> "" Then
      number_of_people = number_of_people + 1
    End If
  Next
  
  '�Ζ����x�̐ݒ�̊����m�F
  If ws.Cells(1, 4) <> "���㐧" And ws.Cells(1, 4) <> "�O��㐧" Then '���㐧, �O��㐧�̂ǂ�������͂���Ă��Ȃ��ꍇ
    MsgBox "�u�Ζ����x�v��, �u���㐧�v,����,�u�O��㐧�v�Ɠ��͂��ĉ����� �I", Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
    Exit Sub
  End If
  
  '���΂̐l���ݒ�̊����m�F
  If ws.Cells(2, 4) = "" Then
    MsgBox "���΂̐l���ݒ�����ĉ������I", _
    Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
    Exit Sub
  ElseIf VarType(ws.Cells(2, 4)) <> 5 Then
    MsgBox "���΂̐l���ݒ�̗��ɂ͐�������͂��ĉ������I" _
    & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
    Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
    Exit Sub
  '�Ζ��\�쐬�N���m�F
  ElseIf Not ws.Cells(9, 4) Like "20*/*/*" Then
    MsgBox "���N�����̋Ζ��\���쐬���邩��D4�Z���ɓ��͂��ĉ������I" _
    & vbCrLf & "��, �`���͈ȉ��̂悤�ɏ����ĉ������I" _
    & vbCrLf & "(��)2023(�N:�����ĉ������I)/1(��:�����ĉ������I)/1(��:1�̂܂ܕύX���Ȃ��ŉ������I)", _
    Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
    Exit Sub
  End If
  
  '���㐧�̃p�����[�^�ݒ�̊����m�F
  If ws.Cells(1, 4) = "���㐧" Then
    '1��1�`�[��������̓����̐l���ݒ�, �A�������ݒ�̊����m�F
    For check_cell_num = 3 To 4
      If ws.Cells(check_cell_num, 4) = "" Then
        MsgBox "������" & check_sub_title(check_sub_title_num) & "�����ĉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Sub
      ElseIf VarType(ws.Cells(check_cell_num, 4)) <> 5 Then
        MsgBox "������" & check_sub_title(check_sub_title_num) & "�̗��ɂ͐�������͂��ĉ������I" _
        & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Sub
      End If
      check_sub_title_num = check_sub_title_num + 1
    Next
    '1�J���̓����̍Œ�񐔂̐ݒ芮���m�F
    For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
      If ws.Cells(check_people_num, 6) = "" Then
        MsgBox ws.Cells(check_people_num, 5) & "�����1�J��������̓����̍Œ�񐔂�ݒ肵�ĉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Sub
      ElseIf VarType(ws.Cells(check_people_num, 6)) <> 5 Then
        MsgBox ws.Cells(check_people_num, 5) & "�����1�J��������̓����̍Œ�񐔂̗��ɂ͐�������͂��ĉ������I" _
        & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Sub
      End If
    Next
    
  '�O��㐧�̃p�����[�^�ݒ�̊����m�F
  ElseIf ws.Cells(1, 4) = "�O��㐧" Then
    check_title_for_three_change(0) = "�����"
    check_title_for_three_change(1) = "�[���"
    '1��1�`�[��������̏�(�[)��΂̐l���ݒ�, �A�������ݒ�̊����m�F
    For check_cell_num = 5 To 8
      If ws.Cells(check_cell_num, 4) = "" Then
        MsgBox check_title_for_three_change(check_title_num_for_three_change) & "��" & check_sub_title(check_sub_title_num) & "�����ĉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Sub
      ElseIf VarType(ws.Cells(check_cell_num, 4)) <> 5 Then
        MsgBox check_title_for_three_change(check_title_num_for_three_change) & "��" & check_sub_title(check_sub_title_num) & "�̗��ɂ͐�������͂��ĉ������I" _
        & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Sub
      End If
      check_sub_title_num = check_sub_title_num + 1
      If check_cell_num = 6 Then
        check_sub_title_num = 0
        check_title_num_for_three_change = check_title_num_for_three_change + 1
      End If
    Next
    '1�J���̏�(�[)��΂̍Œ�񐔂̐ݒ芮���m�F
    check_title_num_for_three_change = 0
    For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
      For check_cell_num = 7 To 8
        If ws.Cells(check_people_num, check_cell_num) = "" Then
          MsgBox ws.Cells(check_people_num, 5) & "�����1�J���������" & check_title_for_three_change(check_title_num_for_three_change) & "�̍Œ�񐔂�ݒ肵�ĉ������I", _
          Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
          Exit Sub
        ElseIf VarType(ws.Cells(check_people_num, check_cell_num)) <> 5 Then
          MsgBox ws.Cells(check_people_num, 5) & "�����1�J���������" & check_title_for_three_change(check_title_num_for_three_change) & "�̍Œ�񐔂̗��ɂ͐�������͂��ĉ������I" _
          & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
          Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
          Exit Sub
        End If
        check_title_num_for_three_change = check_title_num_for_three_change + 1
      Next
      check_title_num_for_three_change = 0
    Next
  End If
 '�l��1�J��������̋x�݂̐ݒ芮���m�F
  
  For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
    If ws.Cells(check_people_num, 9) = "" Then
      MsgBox ws.Cells(check_people_num, 5) & "�����1�J��������̋x�݂̍Œ�񐔂�ݒ肵�ĉ������I", _
      Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
      Exit Sub
    ElseIf VarType(ws.Cells(check_people_num, 9)) <> 5 Then
      MsgBox ws.Cells(check_people_num, 5) & "�����1�J��������̋x�݂̍Œ�񐔂̗��ɂ͐�������͂��ĉ������I" _
      & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
      Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
      Exit Sub
    End If
  Next
  
  MsgBox "�����A�`�[���Ζ���]�\�̏��������ł��I" _
      & vbCrLf & "��]���������񂾌�, �Ζ��\�����쐬�����s���ĉ������I", _
      Buttons:=vbInformation, Title:="�ݒ�҂̕��ւ̃��b�Z�[�W"
  Exit Sub
  
ErrLabel:
  MsgBox "�uA�`�[���p�Ζ���]�������ݑO�m�F�vVBA���s���ɃG���[���������܂����I" _
  & vbCrLf & "�V�[�g�uA�`�[���p�Ζ���]�\�v�����m�F�̏�, �ēx, ���s���ĉ������I" _
  & vbCrLf & "�G���[�����A�v��: " & Err.Source _
  & vbCrLf & "�G���[�ԍ�: " & Err.Number _
  & vbCrLf & "�G���[���e: " & Err.Description, _
  Buttons:=vbCritical, Title:="���s�҂̕������̃G���[�\��"
End Sub
