Attribute VB_Name = "Module4"
Option Explicit
'Private�ϐ��̒�`(���ꃂ�W���[�����ł�Public�ϐ��Ɠ����l�Ɋ֐����܂����Ŏg����)
  '�V�[�g�w��p�ϐ�
Private ws As Worksheet
Private Names As New Collection
Private how_work As String
Private number_of_day As Integer
Private number_of_people As Integer
Private request_table() As String
Private request_aroud_results_day() As Integer
Private request_aroud_results_month() As Integer
Private hope_or_not() As Integer
Private work_name_before As String
Private work_name_after As String
Private max_person_num As Integer
Private max_finding As Integer
Private fit_object_exist As String
'���[�v�ϐ��̒�`
Private j As Integer 'number_of_day�̃��[�v�p��j
Function pro_unpro(ByVal pro_or_unpro As Boolean, ByVal pro_unpro_sheet_name As String, Optional ByVal passward As String = "tokubetunatoki")
  If pro_or_unpro Then
    If ThisWorkbook.Worksheets(pro_unpro_sheet_name).ProtectContents = False Then
      ThisWorkbook.Worksheets(pro_unpro_sheet_name).Protect Password:=passward
    End If
  Else
    If ThisWorkbook.Worksheets(pro_unpro_sheet_name).ProtectContents = True Then
      ThisWorkbook.Worksheets(pro_unpro_sheet_name).Unprotect Password:=passward
    End If
  End If
End Function
Function high_speeding(ByVal start_end As Boolean) 'VBA�J�n�̎���True, �I���̎���False��VBA���������J�n�E�I������
  If start_end = True Then
    ' [VBA������]
    '(�㏑���ۑ��̌x�����b�Z�[�W���܂�)�x�����b�Z�[�W�𖳎�����悤�ɐݒ肷��
    Application.DisplayAlerts = False
    '�`����~����
    Application.ScreenUpdating = False
    '�C�x���g��}������
    Application.EnableEvents = False
    '�X�e�[�^�X�o�[�𖳌��ɂ���
    Application.DisplayStatusBar = False
    '�����v�Z�̒�~
    Application.Calculation = xlCalculationManual
  Else
    ' [VBA������]
    '�x�����b�Z�[�W��\��������悤�ɒ���
    Application.DisplayAlerts = True
    '�`����ĊJ����
    Application.ScreenUpdating = True
     '�C�x���g�̗}�����������
    Application.EnableEvents = True
    '�X�e�[�^�X�o�[��L���ɂ���
    Application.DisplayStatusBar = True
    '�����v�Z�ĊJ
    Application.Calculation = xlCalculationAutomatic
  End If
End Function
Function Before_building_check() As String
  Before_building_check = "�ݒ�s��, �G���[, �����s���L��"
'Function Before_building_check���̕ϐ��̒�`
  Dim check_people_num As Integer 'number_of_people�̃��[�v�p�̕ϐ�
  Dim check_cell_num As Integer
  Dim check_sub_title(1) As String
  check_sub_title(0) = "�l���ݒ�"
  check_sub_title(1) = "�A���񐔐ݒ�"
  Dim check_sub_title_num As Integer
  Dim check_title_for_three_change(1) As String
  Dim check_title_num_for_three_change As Integer
'�p�����[�^�ݒ�̊����m�F

  '�Ζ����x�̐ݒ�̊����m�F
  If ws.Cells(1, 4) <> "���㐧" And ws.Cells(1, 4) <> "�O��㐧" Then '���㐧, �O��㐧�̂ǂ�������͂���Ă��Ȃ��ꍇ
    MsgBox "�u�Ζ����x�v��, �u���㐧�v,����,�u�O��㐧�v�Ɠ��͂��ĉ����� �I", Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
    Exit Function
  End If
  
  '���΂̐l���ݒ�̊����m�F
  If ws.Cells(2, 4) = "" Then
    MsgBox "���΂̐l���ݒ�����ĉ������I", _
    Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
    Exit Function
  ElseIf VarType(ws.Cells(2, 4)) <> 5 Then
    MsgBox "���΂̐l���ݒ�̗��ɂ͐�������͂��ĉ������I" _
    & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
    Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
    Exit Function
  '�Ζ��\�쐬�N���m�F
  ElseIf Not ws.Cells(9, 4) Like "20*/*/*" Then
    MsgBox "���N�����̋Ζ��\���쐬���邩��D4�Z���ɓ��͂��ĉ������I" _
    & vbCrLf & "��, �`���͈ȉ��̂悤�ɏ����ĉ������I" _
    & vbCrLf & "(��)2023(�N:�����ĉ������I)/1(��:�����ĉ������I)/1(��:1�̂܂ܕύX���Ȃ��ŉ������I)", _
    Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
    Exit Function
  End If
  
  '���㐧�̃p�����[�^�ݒ�̊����m�F
  If ws.Cells(1, 4) = "���㐧" Then
    '1��1�`�[��������̓����̐l���ݒ�, �A�������ݒ�̊����m�F
    For check_cell_num = 3 To 4
      If ws.Cells(check_cell_num, 4) = "" Then
        MsgBox "������" & check_sub_title(check_sub_title_num) & "�����ĉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Function
      ElseIf VarType(ws.Cells(check_cell_num, 4)) <> 5 Then
        MsgBox "������" & check_sub_title(check_sub_title_num) & "�̗��ɂ͐�������͂��ĉ������I" _
        & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Function
      End If
      check_sub_title_num = check_sub_title_num + 1
    Next
    '1�J���̓����̍Œ�񐔂̐ݒ芮���m�F
    For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
      If ws.Cells(check_people_num, 6) = "" Then
        MsgBox Names((check_people_num - 19) / 2 + 1) & "�����1�J��������̓����̍Œ�񐔂�ݒ肵�ĉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Function
      ElseIf VarType(ws.Cells(check_people_num, 6)) <> 5 Then
        MsgBox Names((check_people_num - 19) / 2 + 1) & "�����1�J��������̓����̍Œ�񐔂̗��ɂ͐�������͂��ĉ������I" _
        & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Function
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
        Exit Function
      ElseIf VarType(ws.Cells(check_cell_num, 4)) <> 5 Then
        MsgBox check_title_for_three_change(check_title_num_for_three_change) & "��" & check_sub_title(check_sub_title_num) & "�̗��ɂ͐�������͂��ĉ������I" _
        & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
        Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
        Exit Function
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
          MsgBox Names((check_people_num - 19) / 2 + 1) & "�����1�J���������" & check_title_for_three_change(check_title_num_for_three_change) & "�̍Œ�񐔂�ݒ肵�ĉ������I", _
          Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
          Exit Function
        ElseIf VarType(ws.Cells(check_people_num, check_cell_num)) <> 5 Then
          MsgBox Names((check_people_num - 19) / 2 + 1) & "�����1�J���������" & check_title_for_three_change(check_title_num_for_three_change) & "�̍Œ�񐔂̗��ɂ͐�������͂��ĉ������I" _
          & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
          Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
          Exit Function
        End If
        check_title_num_for_three_change = check_title_num_for_three_change + 1
      Next
      check_title_num_for_three_change = 0
    Next
  End If
 '�l��1�J��������̋x�݂̐ݒ芮���m�F
  
  For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
    If ws.Cells(check_people_num, 9) = "" Then
      MsgBox Names((check_people_num - 19) / 2 + 1) & "�����1�J��������̋x�݂̍Œ�񐔂�ݒ肵�ĉ������I", _
      Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
      Exit Function
    ElseIf VarType(ws.Cells(check_people_num, 9)) <> 5 Then
      MsgBox Names((check_people_num - 19) / 2 + 1) & "�����1�J��������̋x�݂̍Œ�񐔂̗��ɂ͐�������͂��ĉ������I" _
      & vbCrLf & "��, ������œ��͂��Ȃ��ŉ������I", _
      Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
      Exit Function
    End If
  Next
  
'�G���[�m�F
  For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
    If ws.Cells(check_people_num, 52) <> "" Then
      MsgBox Names((check_people_num - 19) / 2 + 1) & "����̊�]���ŃG���[���������Ă��܂��I", _
      Buttons:=vbCritical, Title:="��]���͎҂̕������̃G���[�\��"
      Exit Function
    End If
  Next
  
  Before_building_check = "�ݒ�s��, �G���[, �����s������"
End Function
Function excess_miss(ByVal day_num_excess_miss As String) As Integer '�u�s�����Ă���v����-�ŕԂ�, �u�]���Ă���v����+��, �u(�w�������)�������v����0�ŕԂ�
  If InStr(day_num_excess_miss, "�s�����Ă��܂�") > 0 Then
    excess_miss = -Val(Left(day_num_excess_miss, InStr(day_num_excess_miss, "��") - 1))
  ElseIf InStr(day_num_excess_miss, "�]���Ă��܂�") > 0 Then
    excess_miss = Val(Left(day_num_excess_miss, InStr(day_num_excess_miss, "��") - 1))
  ElseIf InStr(day_num_excess_miss, "������") > 0 Then
    excess_miss = 0
  End If
End Function
Function after_find_function(ByVal classify_num As String, ByVal change_need_value As Integer)
  max_person_num = classify_num
  max_finding = change_need_value
  fit_object_exist = "�L��"
End Function
Function Conditional_Branching(ByVal class_num As Integer, ByVal Day_Num As Integer, ByVal Work_Name As String, ByVal equal_or_not_equal As String, ByVal compare_num As Integer, ByVal big_or_small As String) As Boolean
  Dim size_out As Boolean
  Dim agreement As Boolean
  If big_or_small = "big" Then
    If Day_Num > compare_num Then
      size_out = True
    Else
      size_out = False
    End If
  ElseIf big_or_small = "small" Then
    If Day_Num < compare_num Then
      size_out = True
    Else
      size_out = False
    End If
  End If
  If size_out = False Then
   If equal_or_not_equal = "equal" Then
      If request_table(class_num, Day_Num) = Work_Name Then
        agreement = True
      Else
        agreement = False
      End If
    ElseIf equal_or_not_equal = "not_equal" Then
      If request_table(class_num, Day_Num) <> Work_Name Then
        agreement = True
      Else
        agreement = False
      End If
    End If
  ElseIf size_out = True Then
    agreement = True
  End If
  Conditional_Branching = size_out Or ((Not size_out) And agreement)
End Function
Function After_day_Conditional_Branching(ByVal person_num As Integer, ByVal Another_Day_Num As Integer, ByVal Another_compare_num As Integer) As Boolean
  '���̊֐��́u���㐧�v�̏ꍇ�݂̂Ɏg��, ����, �x�݂��u�������������x�݁v�ɕς��鎞�Ɂu�������x�݁v�̕����������ł��邩�m�F����֐��ł���
  '�w��������̌��Ɋ܂܂�Ȃ�(�܂�, ���������)��, ����, ���̌��Ɋ܂܂�邪, ���̌�, ����Ԃ̋Ζ��`���͌Œ��]�ł͂Ȃ�, ����, ���΂͐l�����]���Ă�����,True��Ԃ�
  Dim Another_size_out As Boolean
  Dim hope_or_not_judge As Boolean
  Dim Surplus_or_equal As Boolean
  If Another_Day_Num > Another_compare_num Then
    Another_size_out = True
  Else
    Another_size_out = False
  End If
  If Another_size_out = False Then
    '�Œ��]�����̊m�F
    If hope_or_not(person_num, Another_Day_Num) <> 1 Then
      hope_or_not_judge = False
    Else
      hope_or_not_judge = True
    End If
    '���̓��̋Ζ��`�������΂Ȃ�, ���̓��̓��΂̍��v�l�����]���Ă��邩�̊m�F
    '���̓��̋Ζ��`�����x�݂Ȃ�, ���̓��̓��΂̍��v�l�����s�����Ă��Ȃ����̊m�F
    If (request_table(person_num, Another_Day_Num) = "��" And request_aroud_results_day(0, Another_Day_Num) > 0) Or _
    (request_table(person_num, Another_Day_Num) = "�x" And request_aroud_results_day(0, Another_Day_Num) >= 0) Then
      Surplus_or_equal = True
    Else
      Surplus_or_equal = False
    End If
  ElseIf Another_size_out = True Then
    Surplus_or_equal = True
  End If
  After_day_Conditional_Branching = (Another_size_out Or ((Not Another_size_out) And Surplus_or_equal)) And (Not hope_or_not_judge)
End Function
Function change_need_value(ByVal month_sum_index As Integer, ByVal excess_miss_index As Integer, ByVal person_num As Integer) As Integer
  '�ύX��̋Ζ��`�������΂������ꍇ,excess_miss_index��0��ݒ肵, �]��s�����̕]���������Ȃ����Ƃɂ���
  If excess_miss_index = 0 + 1 Then
    change_need_value = request_aroud_results_month(month_sum_index, person_num)
  Else
    change_need_value = request_aroud_results_month(month_sum_index, person_num) - request_aroud_results_month(excess_miss_index, person_num) * 100
  End If
End Function
Function day_work_classifying(ByVal Work_Name As String) As Variant '�u����, ����, �����, �[���, �x�݂̕��̒��Łv�u�����v���]���Ă���, ����, �ő�̕�+���̏����v��T�����鎞�ɂ��̕��̗v�f�����擾����ׂɎg��
  Dim day_work_classify_num As Integer
  Dim classify_people_num As Integer
  max_person_num = 0
  max_finding = 0
  fit_object_exist = "����"
  Dim answer(1) As Variant
  Dim after_work_classify_num As Integer
  Dim need_value As Integer
  '����,�x�݂͓��㐧, �O��㐧�̗����̏ꍇ�ł���̂�, ������num����
  If Work_Name = "��" Then
    day_work_classify_num = 0
  End If
  
  If work_name_after = "��" Then
    after_work_classify_num = 0
  End If
  
  If Work_Name = "�x" Then
    day_work_classify_num = UBound(request_aroud_results_month, 1) - 1
  End If
  
  If work_name_after = "�x" Then
    after_work_classify_num = UBound(request_aroud_results_month, 1) - 1
  End If
  
  If how_work = "���㐧" Then
    If Work_Name = "��" Then
      day_work_classify_num = 1
    End If
    If work_name_after = "��" Then
      after_work_classify_num = 1
    End If
  
    For classify_people_num = 0 To number_of_people - 1
      '�Œ�Ζ���]�łȂ����
      If hope_or_not(classify_people_num, j) = 0 Then
        need_value = change_need_value(day_work_classify_num, after_work_classify_num + 1, classify_people_num)
        If Work_Name = "��" Then
          '���΂̒T���̏ꍇ, �]��, �s�������̕\���͂Ȃ���, ���v�l�̔�r�݂̂��s��
          '���΂̕��̒��Ō����v���ő�̕��̐l���ԍ���max_person_num�ɑ����, �����v�l��max_finding�ɑ������
          '��, ���΂𓖒��ɕς��鎞��3���O�ɓ������Ȃ����Ƃƌ��3���Ԃ͓������Ȃ�����,
          '��̓���Ԃ��������x�݂ɕς��Ă�������(�܂�, ��̓���Ԃ̓��΂�1���̗]�肪���邱��)���m�F����
          If request_table(classify_people_num, j) = "��" Then
            If max_finding = 0 Or need_value > max_finding Then
              If work_name_after = "��" Then '�����ɕς���ꍇ
                If Conditional_Branching(classify_people_num, j - 3, "��", "not_equal", 0, "small") And _
                Conditional_Branching(classify_people_num, j + 1, "��", "not_equal", number_of_day - 1, "big") And _
                Conditional_Branching(classify_people_num, j + 2, "��", "not_equal", number_of_day - 1, "big") And _
                Conditional_Branching(classify_people_num, j + 3, "��", "not_equal", number_of_day - 1, "big") And _
                After_day_Conditional_Branching(classify_people_num, j + 1, number_of_day - 1) And _
                After_day_Conditional_Branching(classify_people_num, j + 2, number_of_day - 1) Then
                  Call after_find_function(classify_people_num, need_value)
                End If
              Else '���΂��x�݂ɕς���ꍇ�́u�������������x�݁v�̕ύX�\�m�F�͂����ɕς���
                Call after_find_function(classify_people_num, need_value)
              End If
            End If
          End If
        ElseIf Work_Name = "��" Then
          If request_table(classify_people_num, j) = Work_Name And request_aroud_results_month(day_work_classify_num + 1, classify_people_num) > 0 And _
          (max_finding = 0 Or need_value > max_finding) Then
          '�����̒T���̏ꍇ, �]��, �s�������̕\���������, �]���Ă��邱�Ƃ��m�F�������, ���v�l�̔�r�݂̂��s��
          '�����̕��̒��Ō����v���ő�̕��̐l���ԍ���max_person_num�ɑ����, �����v�l��max_finding�ɑ������
            Call after_find_function(classify_people_num, need_value)
          End If
        ElseIf Work_Name = "�x" Then
          If request_table(classify_people_num, j) = Work_Name And request_aroud_results_month(day_work_classify_num + 1, classify_people_num) > 0 And _
            (max_finding = 0 Or need_value > max_finding) Then
          '�x�݂̒T���̏ꍇ, �]��, �s�������̕\���������, �]���Ă��邱�Ƃ��m�F�������, ���v�l�̔�r�݂̂��s��
          '�x�݂̕��̒��Ō����v���ő�̕��̐l���ԍ���max_person_num�ɑ����, �����v�l��max_finding�ɑ������
          
          '�x�݂���΂ɕς��鎞��
          '�O2���Ԃɓ������������Ƃ��m�F����
            If work_name_after = "��" Then
              If Conditional_Branching(classify_people_num, j - 2, "��", "not_equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "��", "not_equal", 0, "small") Then
                Call after_find_function(classify_people_num, need_value)
              End If
          '�x�݂𓖒��ɕς��鎞��
          '2,3���O�ɓ������Ȃ�, ����, ���3���Ԃɓ������Ȃ����Ƃ�
          '��̓���Ԃ��������x�݂ɕς��Ă�������(�܂�, ��̓���Ԃ̓��΂�1���̗]�肪���邱��)���m�F����
            ElseIf work_name_after = "��" Then
              If Conditional_Branching(classify_people_num, j - 2, "��", "not_equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 3, "��", "not_equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j + 1, "��", "not_equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "��", "not_equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 3, "��", "not_equal", number_of_day - 1, "big") And _
              After_day_Conditional_Branching(classify_people_num, j + 1, number_of_day - 1) And _
              After_day_Conditional_Branching(classify_people_num, j + 2, number_of_day - 1) Then
                Call after_find_function(classify_people_num, need_value)
              End If
            End If
          End If
        End If
      End If
    Next
  
  ElseIf how_work = "�O��㐧" Then
    If Work_Name = "��" Then
      day_work_classify_num = 1
    ElseIf Work_Name = "�[" Then
      day_work_classify_num = 3
    End If
    If work_name_after = "��" Then
      after_work_classify_num = 1
    ElseIf work_name_after = "�[" Then
      after_work_classify_num = 3
    End If
    
    For classify_people_num = 0 To number_of_people - 1
      '�Œ�Ζ���]�łȂ����
      If hope_or_not(classify_people_num, j) = 0 Then
        need_value = change_need_value(day_work_classify_num, after_work_classify_num + 1, classify_people_num)
        '���΂̒T���̏ꍇ, �]��, �s�������̕\���͂Ȃ���, ���v�l�̔�r�݂̂��s��
        '���΂̕��̒��Ō����v���ő�̕��̐l���ԍ���max_person_num�ɑ����, �����v�l��max_finding�ɑ������
        If Work_Name = "��" Then
          If request_table(classify_people_num, j) = "��" And _
          (max_finding = 0 Or need_value > max_finding) Then
          '��, ���΂�����΂ɕς��鎞��, �O�����x�݂����� ����, (�������x��, ��������, ������2���オ�u�����x�v�ł���)
          '���Ƃ��m�F����
            If work_name_before = "��" And work_name_after = "��" Then
              If (Conditional_Branching(classify_people_num, j - 1, "�x", "equal", 0, "small") Or _
              Conditional_Branching(classify_people_num, j - 1, "��", "equal", 0, "small")) And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") Or _
              (Conditional_Branching(classify_people_num, j + 1, "��", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "�x", "equal", number_of_day - 1, "big")) _
              ) Then
                Call after_find_function(classify_people_num, need_value)
              End If
            '��, ���΂�[��΂ɕς��鎞��, �O�����x�� ����, (�������x��, ��������, ������2���オ�u�����x�v�ł���)
            '���Ƃ��m�F����
            ElseIf work_name_before = "��" And work_name_after = "�[" Then
              If Conditional_Branching(classify_people_num, j - 1, "�x", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") Or _
              (Conditional_Branching(classify_people_num, j + 1, "��", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "�x", "equal", number_of_day - 1, "big")) _
              ) Then
                Call after_find_function(classify_people_num, need_value)
              End If
            Else '���΂��x�݂ɕς��鎞�͏����͂��̂܂�
                Call after_find_function(classify_people_num, need_value)
            End If
          End If
        '�����,�[���, �x��(���ΈȊO)�̒T���̏ꍇ, �]��, �s�������̕\���������, �]���Ă��邱�Ƃ��m�F�������, ���v�l�̔�r���s��
        '�����,�[���, �x��(���ΈȊO)�̕��̒��Ō����v���ő�̕��̐l���ԍ���max_person_num�ɑ����, �����v�l��max_finding�ɑ������
        ElseIf Work_Name = "��" Then
          If request_table(classify_people_num, j) = "��" And request_aroud_results_month(day_work_classify_num + 1, classify_people_num) > 0 And _
          (max_finding = 0 Or need_value > max_finding) Then
          '��, ����΂���΂ɕς��鎞��, �O���������, ����, �[��Ζ��ł͂Ȃ����Ƃ��m�F����
            If work_name_before = "��" And work_name_after = "��" Then
              If Conditional_Branching(classify_people_num, j - 1, "��", "not_equal", 0, "small") And _
                Conditional_Branching(classify_people_num, j - 1, "�[", "not_equal", 0, "small") Then
                  Call after_find_function(classify_people_num, need_value)
              End If
            '��, ����΂�[��΂ɕς��鎞��, (�O�����x�݂ł���,����, (�������x��, ����, ������2���オ�����x�݂ł���))
            '����, ((2���O��1���O���x���[), ����, 1���オ�x�ł���)���Ƃ��m�F����
            ElseIf work_name_before = "��" And work_name_after = "�[" Then
              If _
              ( _
              Conditional_Branching(classify_people_num, j - 2, "�x", "equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "�[", "equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") _
              ) Or _
              ( _
              Conditional_Branching(classify_people_num, j - 1, "�x", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") Or _
              (Conditional_Branching(classify_people_num, j + 1, "��", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "�x", "equal", number_of_day - 1, "big")) _
              ) _
              ) Then
                Call after_find_function(classify_people_num, need_value)
              End If
            Else '����΂��x�݂ɕς��鎞�͏����͂��̂܂�
              Call after_find_function(classify_people_num, need_value)
            End If
          End If
        
        ElseIf Work_Name = "�[" Then
          If request_table(classify_people_num, j) = "�[" And request_aroud_results_month(day_work_classify_num + 1, classify_people_num) > 0 And _
          (max_finding = 0 Or need_value > max_finding) Then
          '��, �[��΂���΂ɕς��鎞��, �O�����x��, ����, (�������x�݂�������2����, ����΁��x�݂ł���)���Ƃ��m�F����
            If work_name_before = "�[" And work_name_after = "��" Then
              If Conditional_Branching(classify_people_num, j - 1, "�x", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") Or _
              (Conditional_Branching(classify_people_num, j + 1, "��", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "�x", "equal", number_of_day - 1, "big")) _
              ) Then
                Call after_find_function(classify_people_num, need_value)
              End If
            '��, �[��΂�����΂ɕς��鎞��, (�O�����x��, ����, (�������x�݂�������2����, ����΁��x�݂ł���))��
            '(2���O���O�����x���[, ����, �������x��)�ł��邱�Ƃ��m�F����
            ElseIf work_name_before = "�[" And work_name_after = "��" Then
              If ( _
              Conditional_Branching(classify_people_num, j - 1, "�x", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") Or _
              (Conditional_Branching(classify_people_num, j + 1, "��", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "�x", "equal", number_of_day - 1, "big")) _
              ) _
              ) _
              Or _
              ( _
              (Conditional_Branching(classify_people_num, j - 2, "�x", "equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "�[", "equal", 0, "small")) And _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") _
              ) _
              Then
                Call after_find_function(classify_people_num, need_value)
              End If
            Else '�[��΂��x�݂ɕς��鎞�͏����͂��̂܂�
                Call after_find_function(classify_people_num, need_value)
            End If
          End If
          
        ElseIf Work_Name = "�x" Then
          If request_table(classify_people_num, j) = "�x" And request_aroud_results_month(day_work_classify_num + 1, classify_people_num) > 0 And _
          (max_finding = 0 Or need_value > max_finding) Then
          '��, �x�݂���΂ɕς��鎞��, (�O������, �[�ł͂Ȃ�,) ����, (�������[�ł͂Ȃ�)���Ƃ��m�F����
            If work_name_before = "�x" And work_name_after = "��" Then
              If ( _
              Conditional_Branching(classify_people_num, j - 1, "��", "not_equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "�[", "not_equal", 0, "small") _
              ) And _
              Conditional_Branching(classify_people_num, j + 1, "�[", "not_equal", number_of_day - 1, "big") Then
                Call after_find_function(classify_people_num, need_value)
              End If
            '��, �x�݂�����΂ɕς��鎞��, (�O�����x��, ����, (�������x�݂�������2����, ����΁��x�݂ł���))��
            '(2���O���O�����x���[, ����, �������x��)�ł��邱�Ƃ��m�F����
            ElseIf work_name_before = "�x" And work_name_after = "��" Then
              If ( _
              Conditional_Branching(classify_people_num, j - 1, "�x", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") Or _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "��", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "�x", "equal", number_of_day - 1, "big") _
              ) _
              ) _
              ) _
              Or _
              ( _
              (Conditional_Branching(classify_people_num, j - 2, "�x", "equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "�[", "equal", 0, "small")) And _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") _
              ) _
              Then
                Call after_find_function(classify_people_num, need_value)
              End If
           '��, �x�݂�[��΂ɕς��鎞��,
           '(�O�����x��, ����, (�������x�݂�������2����, (����΂��[��)���x�݂ł���))��
           '(2���O���O�����x���[, ����, �������x��)�ł��邱�Ƃ��m�F����
            ElseIf work_name_before = "�x" And work_name_after = "�[" Then
              If ( _
              Conditional_Branching(classify_people_num, j - 1, "�x", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") Or _
              ( _
              (Conditional_Branching(classify_people_num, j + 1, "��", "equal", number_of_day - 1, "big") Or Conditional_Branching(classify_people_num, j + 1, "�[", "equal", number_of_day - 1, "big")) And _
              Conditional_Branching(classify_people_num, j + 2, "�x", "equal", number_of_day - 1, "big") _
              ) _
              ) _
              ) _
              Or _
              ( _
              (Conditional_Branching(classify_people_num, j - 2, "�x", "equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "�[", "equal", 0, "small")) And _
              Conditional_Branching(classify_people_num, j + 1, "�x", "equal", number_of_day - 1, "big") _
              ) _
              Then
                Call after_find_function(classify_people_num, need_value)
              End If
            End If
          End If
        End If
      End If
    Next
  End If
  answer(0) = max_person_num
  answer(1) = fit_object_exist
  day_work_classifying = answer
End Function
Function work_name_decide(ByVal before_name As String, ByVal after_name As String)
  work_name_before = before_name
  work_name_after = after_name
End Function
Function change_write()
  Dim change_before_num As Integer
  Dim change_after_num As Integer
  Dim change_person_num As Integer
  Dim keep_num_for_change As Integer
  
  If day_work_classifying(work_name_before)(1) = "�L��" Then
    keep_num_for_change = day_work_classifying(work_name_before)(0)
    request_table(keep_num_for_change, j) = work_name_after
  Else
    GoTo Point_Change_End
  End If
  
  '���΂��x�݂������ꍇ, ���㐧, �O��㐧�̂ǂ�������΂Ƌx�݂͂���̂�,
  '�ŏ���change_before_num, change_after_num�����蓖�Ă�
  If work_name_before = "��" Then
      change_before_num = 0
  End If
  
  If work_name_after = "��" Then
      change_after_num = 0
  End If
  
  If work_name_before = "�x" Then
      change_before_num = UBound(request_aroud_results_month, 1) - 1
  End If
  
  If work_name_after = "�x" Then
      change_after_num = UBound(request_aroud_results_month, 1) - 1
  End If
  
  'chage_person_num���Z�o����
  
  change_person_num = keep_num_for_change
  
  If how_work = "���㐧" Then
  '���㐧�ɂ�����ύX�O�̋Ζ����ԑт��ύX����邱�Ƃɂ��1���̗]��E�s���l���\���ƌ����v�ƌ����v�̗]��E�s���l���\���̕ύX
    If work_name_before = "��" Then
       change_before_num = 1
    End If
    
    If work_name_before <> "�x" Then '�ύX�O������,����(�x�݈ȊO)�̏ꍇ, 1���̗]��E�s���l���\����-1����
      request_aroud_results_day(change_before_num, j) = request_aroud_results_day(change_before_num, j) - 1
    End If
    
    request_aroud_results_month(change_before_num, change_person_num) = _
    request_aroud_results_month(change_before_num, change_person_num) - 1
    
    If work_name_before <> "��" Then '�ύX�O������,�x��(���ΈȊO)�̒T���̏ꍇ, �����v�̗]��, �s�������̕\����-1����
      request_aroud_results_month(change_before_num + 1, change_person_num) = _
      request_aroud_results_month(change_before_num + 1, change_person_num) - 1
    End If
      
    '�ύX�O��������, ���������j����������, �ύX�O�̓�������΂��x�݂ɕς���Ɠ�����, �����̖�������΂��x�݂ɕς���
    If work_name_before = "��" And j < number_of_day - 1 Then
      request_table(change_person_num, j + 1) = work_name_after
      If work_name_after = "��" Then
        request_aroud_results_day(0, j + 1) = request_aroud_results_day(0, j + 1) + 1 '1���̓��΂̗]��E�s���l���\����+1����
        request_aroud_results_month(0, change_person_num) = _
        request_aroud_results_month(0, change_person_num) + 1 '���΂̌����v��+1����
      ElseIf work_name_after = "�x" Then
        request_aroud_results_month(3, change_person_num) = _
        request_aroud_results_month(3, change_person_num) + 1 '�x�݂̌����v��+1����
        request_aroud_results_month(4, change_person_num) = _
        request_aroud_results_month(4, change_person_num) + 1 '�x�݂̌����v�̗]��E�s���l���\����+1����
      End If
    End If
    
  '���㐧�ɂ�����ύX��̋Ζ����ԑт��ύX����邱�Ƃɂ��1���̗]��E�s���l���\���ƌ����v�ƌ����v�̗]��E�s���l���\���̕ύX
    If work_name_after = "��" Then
      change_after_num = 1
    End If
    
    If work_name_after <> "�x" Then '�ύX�オ����,����(�x�݈ȊO)�̏ꍇ, 1���̗]��E�s���l���\����+1����
      request_aroud_results_day(change_after_num, j) = request_aroud_results_day(change_after_num, j) + 1
    End If
    
    '�ύX��̋Ζ����ԑт̌����v��+1����
    request_aroud_results_month(change_after_num, change_person_num) = _
    request_aroud_results_month(change_after_num, change_person_num) + 1
    
    If work_name_after <> "��" Then '�ύX�オ����,�x��(���ΈȊO)�̒T���̏ꍇ, �����v�̗]��, �s�������̕\����+1����
      request_aroud_results_month(change_after_num + 1, change_person_num) = _
      request_aroud_results_month(change_after_num + 1, change_person_num) + 1
    End If
      
    '�ύX�オ������, ���������j��菬����,�܂�,�����܂ōŒ�ł�1���͋󂢂Ă��鎞,
    '���΂��x�݂𓖒�(�ύX��)�ɕς���Ɠ�����, �����̓��΂��x�݂��u���v�� �ɕς���
    If work_name_after = "��" And j < number_of_day - 1 Then
      If request_table(change_person_num, j + 1) = "��" Then
        request_aroud_results_day(0, j + 1) = request_aroud_results_day(0, j + 1) - 1 '1���̓��΂̗]��E�s���l���\����-1����
        request_aroud_results_month(0, change_person_num) = _
        request_aroud_results_month(0, change_person_num) - 1 '���΂̌����v��-1����
      ElseIf request_table(change_person_num, j + 1) = "�x" Then
        request_aroud_results_month(3, change_person_num) = _
        request_aroud_results_month(3, change_person_num) - 1 '�x�݂̌����v��-1����
        request_aroud_results_month(4, change_person_num) = _
        request_aroud_results_month(4, change_person_num) - 1 '�x�݂̌����v�̗]��E�s���l���\����+1����
      End If
      request_table(change_person_num, j + 1) = "��"
      '���������j-1��菬����,�܂�,�����܂ōŒ�ł�2���͋󂢂Ă��鎞,
      '���΂��x�݂𓖒�(�ύX��)�ɕς���Ɠ�����, 2����̓��΂��u�x�v�� �ɕς���
      If j < (number_of_day - 1) - 1 Then
        If request_table(change_person_num, j + 2) = "��" Then
          request_aroud_results_day(0, j + 2) = request_aroud_results_day(0, j + 2) - 1 '1���̓��΂̗]��E�s���l���\����-1����
          request_aroud_results_month(0, change_person_num) = _
          request_aroud_results_month(0, change_person_num) - 1 '���΂̌����v��-1����
        End If
        request_table(change_person_num, j + 2) = "�x"
      End If
    End If
  
  ElseIf how_work = "�O��㐧" Then
    If work_name_before = "��" Then
      change_before_num = 1
    ElseIf work_name_before = "�[" Then
      change_before_num = 3
    End If
    
    '�ύX�O������, �����, �[��΂̏ꍇ(1���̗]��E�s���\����-1����)
    If work_name_before = "��" Or work_name_before = "��" Then
      request_aroud_results_day(change_before_num, j) = request_aroud_results_day(change_before_num, j) - 1
    ElseIf work_name_before = "�[" Then
      request_aroud_results_day(2, j) = request_aroud_results_day(2, j) - 1
    End If
    
    request_aroud_results_month(change_before_num, change_person_num) = _
    request_aroud_results_month(change_before_num, change_person_num) - 1
    
    If work_name_before <> "��" Then '�ύX�O�������, �[���,�x��(���ΈȊO)�̒T���̏ꍇ, �����v�̗]��, �s�������̕\��������
      request_aroud_results_month(change_before_num + 1, change_person_num) = _
      request_aroud_results_month(change_before_num + 1, change_person_num) - 1
    End If

    If work_name_after = "��" Then
      change_after_num = 1
    ElseIf work_name_after = "�[" Then
      change_after_num = 3
    End If
    
    '�ύX�オ����, �����, �[��΂̏ꍇ(1���̗]��E�s���\����+1����)
    If work_name_after = "��" Or work_name_after = "��" Then
      request_aroud_results_day(change_after_num, j) = request_aroud_results_day(change_after_num, j) + 1
    ElseIf work_name_after = "�[" Then
      request_aroud_results_day(2, j) = request_aroud_results_day(2, j) + 1
    End If
    
    request_aroud_results_month(change_after_num, change_person_num) = _
    request_aroud_results_month(change_after_num, change_person_num) + 1
    
    If work_name_after <> "��" Then '�ύX�オ�����, �[���,�x��(���ΈȊO)�̒T���̏ꍇ, �����v�̗]��, �s�������̕\��������
      request_aroud_results_month(change_after_num + 1, change_person_num) = _
      request_aroud_results_month(change_after_num + 1, change_person_num) + 1
    End If
  End If
Point_Change_End:
End Function
Function inversion(ByVal inversion_num As Integer) As Integer
  If inversion_num = 0 Then
    inversion = 1
  ElseIf inversion_num = 1 Then
    inversion = 0
  End If
End Function
Function insufficient_compensation(ByVal insufficient_work_name As String)
  Dim insufficient_num As Integer
  Dim compensation_work(1) As String
  If how_work = "���㐧" Then
    If insufficient_work_name = "��" Then
     insufficient_num = 0
     compensation_work(0) = "��"
    ElseIf insufficient_work_name = "��" Then
      insufficient_num = 1
      compensation_work(0) = "��"
    End If
    Call work_name_decide(compensation_work(0), insufficient_work_name)
    If request_aroud_results_day(inversion(insufficient_num), j) > 0 Then '���̓��̕�U�Ζ����ԑѐl�����]���Ă�����,
      While request_aroud_results_day(insufficient_num, j) < 0 And _
      day_work_classifying(compensation_work(0))(1) = "�L��" And _
      request_aroud_results_day(inversion(insufficient_num), j) > 0
      '1���̂��̐l���s����0�ɂȂ邩��U�Ζ����ԑѐl���̌����v�ŗ]���Ă���������Ȃ��Ȃ邩���̓��̕�U�Ζ����ԑѐl�����]��Ȃ��Ȃ鎞�܂�
       Call change_write
       '��U�Ζ����ԑт̕��̒��Ō����v���]���Ă���, ����, �����v���ő�̕�U�Ζ����ԑт̕���s���Ζ��ɕς���
      Wend
      If request_aroud_results_day(insufficient_num, j) = 0 Then '���̓��̕s���Ζ����s�����Ă����Ȃ��Ă�����, �p�X
      ElseIf day_work_classifying(compensation_work(0))(1) = "����" Or _
      request_aroud_results_day(inversion(insufficient_num), j) <= 0 Then
      '���̓��̕s���Ζ����܂��s�����Ă���, ��U�Ζ����ԑт̌����v���]���Ă���l�����̓��̕�U�Ζ����ԑт̗]��l�����Ȃ�������,
        Call work_name_decide("�x", insufficient_work_name)
        While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("�x")(1) = "�L��"
        '1���̕s���Ζ��̕s���l����0�ɂȂ邩�x�݂̌����v�ŗ]���Ă���������Ȃ��Ȃ鎞�܂�
          Call change_write
        Wend
      End If
    Else '���̓��̕�U�Ζ����ԑѐl�����]���Ă��Ȃ�������,
      Call work_name_decide("�x", insufficient_work_name)
      While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("�x")(1) = "�L��"
        '1���̕s���Ζ��̕s���l����0�ɂȂ邩�x�݂̌����v�ŗ]���Ă���������Ȃ��Ȃ鎞�܂�
        Call change_write
      Wend
    End If
  ElseIf how_work = "�O��㐧" Then
    Dim compensation_num(1) As Integer
    Dim three_change_loop_num As Integer
    If insufficient_work_name = "��" Then
      insufficient_num = 0
      compensation_num(0) = 1
      compensation_num(1) = 2
      compensation_work(0) = "��"
      compensation_work(1) = "�["
    ElseIf insufficient_work_name = "��" Then
      insufficient_num = 1
      compensation_num(0) = 0
      compensation_num(1) = 2
      compensation_work(0) = "��"
      compensation_work(1) = "�["
    ElseIf insufficient_work_name = "�[" Then
      insufficient_num = 2
      compensation_num(0) = 0
      compensation_num(1) = 1
      compensation_work(0) = "��"
      compensation_work(1) = "��"
    End If
    
    For three_change_loop_num = 0 To 1
      If request_aroud_results_day(compensation_num(three_change_loop_num), j) > 0 And _
      request_aroud_results_day(compensation_num(inversion(three_change_loop_num)), j) <= 0 Then '�ŏ�(���)�̕��̕�U�Ζ����ԑѐl���݂̂��]���Ă�����,
        Call work_name_decide(compensation_work(three_change_loop_num), insufficient_work_name)
        While request_aroud_results_day(insufficient_num, j) < 0 And _
        day_work_classifying(compensation_work(three_change_loop_num))(1) = "�L��" And _
        request_aroud_results_day(compensation_num(three_change_loop_num), j) > 0
        '1���̂��̐l���s����0�ɂȂ邩�ŏ�(���)�̕��̕�U�Ζ����ԑѐl���̌����v�ŗ]���Ă���������Ȃ��Ȃ邩
        '���̓��̍ŏ�(���)�̕��̕�U�Ζ����ԑѐl�����]��Ȃ��Ȃ鎞�܂�
         Call change_write
         '��U�Ζ����ԑт̕��̒��Ō����v���]���Ă���, ����, �����v���ő�̕�U�Ζ����ԑт̕���s���Ζ��ɕς���
        Wend
        If request_aroud_results_day(insufficient_num, j) = 0 Then '���̓��̕s���Ζ����s�����Ă����Ȃ��Ă�����, �p�X
        ElseIf day_work_classifying(compensation_work(three_change_loop_num))(1) = "����" Or request_aroud_results_day(compensation_num(three_change_loop_num), j) <= 0 Then
        '���̓��̕s���Ζ����܂��s�����Ă���, ��U�Ζ����ԑт̌����v���]���Ă���l�����̓��̕�U�Ζ����ԑт̗]��l�����Ȃ�������,
          Call work_name_decide("�x", insufficient_work_name)
          While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("�x")(1) = "�L��"
          '1���̕s���Ζ��̕s���l����0�ɂȂ邩�x�݂̌����v�ŗ]���Ă���������Ȃ��Ȃ鎞�܂�
            Call change_write
          Wend
        End If
      End If
    Next
    
    If request_aroud_results_day(compensation_num(0), j) > 0 And _
    request_aroud_results_day(compensation_num(1), j) > 0 Then '���̓��̕�U�Ζ����ԑѐl�����ǂ�����]���Ă�����,
      If insufficient_work_name = "��" Then '��U�Ζ����ԑт�����΂Ɛ[��Ζ��ł���ꍇ, (�s���Ζ����ԑт����΂̏ꍇ)
        If request_aroud_results_day(compensation_num(0), j) > request_aroud_results_day(compensation_num(1), j) Then '1��������ŏ���΂��[��΂��]���Ă�����,
          While request_aroud_results_day(insufficient_num, j) < 0 And _
          day_work_classifying("��")(1) = "�L��" And _
          request_aroud_results_day(compensation_num(0), j) > request_aroud_results_day(compensation_num(1), j) '���΂��s����, �����v�̊ϓ_����ύX�\�ȏ���΂�����, 1�������菀��΂��[��΂��]���Ă�����,
            Call work_name_decide("��", "��")
            Call change_write  '�܂�, ����΂��U�ɏ[�Ă�
          Wend
          If request_aroud_results_day(insufficient_num, j) = 0 Then
            GoTo Point_End
          ElseIf day_work_classifying("��")(1) = "����" Then
            GoTo Midnight_Start
          ElseIf request_aroud_results_day(compensation_num(0), j) = request_aroud_results_day(compensation_num(1), j) Then
           GoTo Semi_Night_Start
          End If
        ElseIf request_aroud_results_day(compensation_num(0), j) < request_aroud_results_day(compensation_num(1), j) Then '1��������Ő[��΂�����΂��]���Ă�����,
          While request_aroud_results_day(insufficient_num, j) < 0 And _
          day_work_classifying("�[")(1) = "�L��" And _
          request_aroud_results_day(compensation_num(1), j) > request_aroud_results_day(compensation_num(0), j) '���΂��s����, �����v�̊ϓ_����ύX�\�Ȑ[��΂�����, 1��������[��΂�����΂��]���Ă�����,
            Call work_name_decide("�[", "��")
            Call change_write  '�[��΂��U�ɏ[�Ă�
          Wend
          If request_aroud_results_day(insufficient_num, j) = 0 Then
            GoTo Point_End
          Else
            GoTo Semi_Night_Start
          End If
        ElseIf request_aroud_results_day(compensation_num(0), j) = request_aroud_results_day(compensation_num(1), j) Then '1��������̗]��l���Ő[���=����΂Ȃ�
Semi_Night_Start:
          Call work_name_decide("��", "��")
          Call change_write  '�܂�, ����΂��U�ɏ[�Ă�
          ElseIf request_aroud_results_day(compensation_num(1), j) > request_aroud_results_day(compensation_num(0), j) Then '1��������[��΂�����΂��]���Ă�����,
            GoTo Midnight_Start
          ElseIf request_aroud_results_day(insufficient_num, j) = 0 Then
            GoTo Point_End
          ElseIf day_work_classifying("�[")(1) = "����" Then
            If day_work_classifying("��")(1) = "�L��" Then
              GoTo Semi_Night_Start
            Else
Holiday_start:
              Call work_name_decide("�x", "��")
              While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("�x")(1) = "�L��"
                Call change_write
              Wend
              GoTo Point_End
            End If
          Else
Midnight_Start:
            Call work_name_decide("�[", "��")
            Call change_write  '�܂�, �[��΂��U�ɏ[�Ă�
            If request_aroud_results_day(insufficient_num, j) = 0 Then
              GoTo Point_End
            ElseIf request_aroud_results_day(1, j) = 0 And request_aroud_results_day(2, j) = 0 Then '����΂Ɛ[��΂̗����Ƃ��P���̋Ζ��l���ɗ]�肪�Ȃ��Ȃ�, �ύX�s�ɂȂ�����
              GoTo Holiday_start
            ElseIf day_work_classifying("��")(1) = "����" Then
              If day_work_classifying("�[")(1) = "�L��" Then
                GoTo Midnight_Start
              Else
                GoTo Holiday_start
              End If
            End If
          End If
        End If
Point_End:
      Else '��U�Ζ����ԑт����΂Ə�(�[)��΂ł���ꍇ
        Call work_name_decide(compensation_work(1), insufficient_work_name) '�܂�, ��(�[)��΂��U�ɏ[�Ă�
        While request_aroud_results_day(insufficient_num, j) < 0 And _
        day_work_classifying(compensation_work(1))(1) = "�L��" And _
        request_aroud_results_day(compensation_num(1), j) > 0
        '1���̂��̐l���s����0�ɂȂ邩��(�[)��΂̕�U�Ζ����ԑѐl���̌����v�ŗ]���Ă���������Ȃ��Ȃ邩
        '���̓��̏�(�[)��΂̕�U�Ζ����ԑѐl�����]��Ȃ��Ȃ鎞�܂�
          Call change_write
         '��(�[)��΂̕��̒��Ō����v���]���Ă���, ����, �����v���ő�̏�(�[)��΂̕���s���Ζ��ɕς���
        Wend
        If request_aroud_results_day(insufficient_num, j) = 0 Then
        ElseIf day_work_classifying(compensation_work(1))(1) = "����" Or request_aroud_results_day(compensation_num(1), j) = 0 Then
          Call work_name_decide(compensation_work(0), insufficient_work_name) '����, ���΂��U�ɏ[�Ă�
          While request_aroud_results_day(insufficient_num, j) < 0 And _
          day_work_classifying(compensation_work(0))(1) = "�L��" And _
          request_aroud_results_day(compensation_num(0), j) > 0
          '1���̂��̐l���s����0�ɂȂ邩���΂̕�U�Ζ����ԑѐl���̌����v�ŗ]���Ă���������Ȃ��Ȃ邩
          '���̓��̓��΂̕��̕�U�Ζ����ԑѐl�����]��Ȃ��Ȃ鎞�܂�
            Call change_write
           '���΂̕��̒��Ō����v���]���Ă���, ����, �����v���ő�̓��΂̕���s���Ζ��ɕς���
          Wend
          If request_aroud_results_day(insufficient_num, j) = 0 Then
          ElseIf day_work_classifying(compensation_work(0))(1) = "�L��" And request_aroud_results_day(compensation_num(0), j) > 0 Then
            Call work_name_decide("�x", insufficient_work_name)
            While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("�x")(1) = "�L��"
            '1���̕s���Ζ��̕s���l����0�ɂȂ邩�x�݂̌����v�ŗ]���Ă���������Ȃ��Ȃ鎞�܂�
              Call change_write
            Wend
          End If
        End If
      End If
    
    If request_aroud_results_day(compensation_num(0), j) <= 0 And _
    request_aroud_results_day(compensation_num(1), j) <= 0 Then '���̓��̕�U�Ζ����ԑѐl�����ǂ�����]���Ă��Ȃ�������,
      Call work_name_decide("�x", insufficient_work_name)
      While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("�x")(1) = "�L��"
        '1���̕s���Ζ��̕s���l����0�ɂȂ邩�x�݂̌����v�ŗ]���Ă���������Ȃ��Ȃ鎞�܂�
          Call change_write
      Wend
    End If
  End If
End Function
Sub B�`�[���p�Ζ��\�����쐬���s()
'Private�ϐ��̍ď�����
  '�O���[�o���ϐ��̒�`
  '�ǂݍ��ݗp�V�[�g�w��
  Set ws = ThisWorkbook.Sheets("B�`�[���p�Ζ���]�\")
  Set Names = New Collection
  how_work = ""
  number_of_day = 0
  number_of_people = 0
  Erase request_table
  Erase request_aroud_results_day
  Erase request_aroud_results_month
  Erase hope_or_not
  work_name_before = ""
  work_name_after = ""
  max_person_num = 0
  max_finding = 0
  fit_object_exist = ""
  '���[�v�ϐ��̒�`
  j = 0 'number_of_day�̃��[�v�p��j
  
'�ǂݍ��ݎ��p��1��1�l�Ζ��`���i�[�ϐ�
  Dim work_content As String
'Naming
  Dim i As Integer 'number_of_people�̃��[�v�p��i
  '�ȉ�, �f�o�b�O�p�ϐ�
  'Dim dbug_j As Integer
  'Dim dbug_titles As Integer
'���x����p�̕]���l�ϐ�
  Dim loss As Long
  Dim weight As Integer
  Dim save_loss As Long
  Dim save_request_for_evaluate() As String
  Dim distinction_first_0 As Boolean 'True�ɂȂ�����, ��x�͕]���������Ƃ�����
  
'���l�ǂݍ��ݗp�ϐ�
  Dim remarks_table() As String
  
  On Error GoTo ErrLabel
  
'exectuing
  
  'name,number_of_people_reading
  For i = 19 To 19 + 2 * (30 - 1) Step 2 'For i=19 To 19+2*(30(�l)-1) Step 2
    If ws.Cells(i, 5) <> "" Then
      Names.Add ws.Cells(i, 5)
      number_of_people = number_of_people + 1
    End If
  Next
  '���O������������, �I������
  If number_of_people = 0 Then
    MsgBox "�Ζ��҂̎�������͂��ĉ������I", _
    Buttons:=vbCritical, Title:="���s�҂̕������̃G���[�\��"
    Exit Sub
  End If
  
  'number_of_day_reading
  For j = 11 To 41 'For j = 11(K) To 41(AO)
    If ws.Cells(18, j) <> "�~" Then
      number_of_day = number_of_day + 1
    End If
  Next
  
  '�J�n�O�̐ݒ�s��, �G���[, �����s���m�F
  If Before_building_check = "�ݒ�s��, �G���[, �����s���L��" Then
    Exit Sub
  End If
  
  'VBA�������J�n
  Call high_speeding(True)
  
  '�Ζ����x�c��
  how_work = ws.Cells(1, 4)
  
  'request_aroud_results
  If how_work = "���㐧" Then
    ReDim request_aroud_results_day(2 - 1, number_of_day - 1)
    ReDim request_aroud_results_month(5 - 1, number_of_people - 1)
    'ReDim request_aroud_results_day((�]���w�W�̐�) - 1, number_of_day(people) - 1)
    '���㐧��1�����̕]���w�W
    'D1.1���̓��΂̗]��, �s���l���\��
    'D2.1���̓����̗]��, �s���l���\��
    '���㐧��1�������̕]���w�W
    'M1.1�J���̓��΂̍��v
    'M2.1�J���̓����̍��v
    'M3.1�J���̓����̗]��, �s�������\��
    'M4.1�J���̋x�݂̍��v
    'M5.1�J���̋x�݂̗]��, �s�������\��
    For j = 11 To 11 + number_of_day - 1
      'D2.1���̓����̗]��, �s���l���\��
      request_aroud_results_day(1, j - 11) = ws.Cells(12, j)
    Next
    For i = 19 To 19 + 2 * (number_of_people - 1) Step 2
    'M2.1�J���̓����̍��v
    'M3.1�J���̓����̗]��, �s�������\��
      request_aroud_results_month(1, (i - 19) / 2) = ws.Cells(i, 43)
      request_aroud_results_month(2, (i - 19) / 2) = excess_miss(ws.Cells(i, 44))
    Next
  ElseIf how_work = "�O��㐧" Then
    ReDim request_aroud_results_day(3 - 1, number_of_day - 1)
    ReDim request_aroud_results_month(7 - 1, number_of_people - 1)
    'ReDim request_aroud_results_day((�]���w�W�̐�) - 1, number_of_day(people) - 1)
    '�O��㐧��1�����̕]���w�W
    'D1.1���̓��΂̗]��, �s���l���\��
    'D2.1���̏���΂̗]��, �s���l���\��
    'D3.1���̐[��΂̗]��, �s���l���\��
    '�O��㐧��1�������̕]���w�W
    'M1.1�J���̓��΂̍��v
    'M2.1�J���̏���΂̍��v
    'M3.1�J���̏���΂̗]��, �s�������\��
    'M4.1�J���̐[��΂̍��v
    'M5.1�J���̐[��΂̗]��, �s�������\��
    'M6.1�J���̋x�݂̍��v
    'M7.1�J���̋x�݂̗]��, �s�������\��
    For j = 11 To 11 + number_of_day - 1
      'D2.1���̏���΂̗]��, �s���l���\��
      request_aroud_results_day(1, j - 11) = ws.Cells(13, j)
      'D3.1���̐[��΂̗]��, �s���l���\��
      request_aroud_results_day(2, j - 11) = ws.Cells(14, j)
    Next
    For i = 19 To 19 + 2 * (number_of_people - 1) Step 2
    'M2.1�J���̏���΂̍��v
    'M3.1�J���̏���΂̗]��, �s�������\��
    'M4.1�J���̐[��΂̍��v
    'M5.1�J���̐[��΂̗]��, �s�������\��
      request_aroud_results_month(1, (i - 19) / 2) = ws.Cells(i, 45)
      request_aroud_results_month(2, (i - 19) / 2) = excess_miss(ws.Cells(i, 46))
      request_aroud_results_month(3, (i - 19) / 2) = ws.Cells(i, 47)
      request_aroud_results_month(4, (i - 19) / 2) = excess_miss(ws.Cells(i, 48))
    Next
  End If
  '�`�ȉ�, ���ʕ����̑���J�n�`
  'D1.1���̓��΂̗]��, �s���l���\��
  For j = 11 To 11 + number_of_day - 1
    request_aroud_results_day(0, j - 11) = ws.Cells(11, j)
  Next
  'M1.1�J���̓��΂̍��v
  'M4(6).1�J���̋x�݂̍��v
  'M5(7).1�J���̋x�݂̗]��, �s�������\��
  For i = 19 To 19 + 2 * (number_of_people - 1) Step 2
    request_aroud_results_month(0, (i - 19) / 2) = ws.Cells(i, 42)
    request_aroud_results_month(UBound(request_aroud_results_month, 1) - 1, (i - 19) / 2) = ws.Cells(i, 49)
    request_aroud_results_month(UBound(request_aroud_results_month, 1), (i - 19) / 2) = excess_miss(ws.Cells(i, 50))
  Next
  '�`�ȏ�, ���ʕ����̑���I���`
  
  'request_table_reading��hope_or_not�̊i�[
  ReDim request_table(number_of_people - 1, number_of_day - 1)
  ReDim hope_or_not(number_of_people - 1, number_of_day - 1)
  ReDim remarks_table(number_of_people - 1, number_of_day - 1)
  For i = 19 To 19 + 2 * (number_of_people - 1) Step 2
    For j = 11 To 11 + number_of_day - 1
      work_content = ws.Cells(i, j)
      If how_work = "���㐧" Then
        If work_content <> "" And work_content <> "��" And work_content <> "�x" And work_content <> "��" And work_content <> "��" Then
          'VBA���������I����
          Call high_speeding(False)
          '�G���[���b�Z�[�W���o�͂�
          MsgBox Names((i - 19) / 2 + 1) & "�����" & CStr(j - 11 + 1) & "���̋Ζ��`�����s�K�؂ł��I" & vbCrLf & "�u���v�u�x�v�u���v�u���v�̂ǂꂩ���L�ڂ��ĉ������I", _
          Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
          'Sub�v���W�[�W�����I������
          Exit Sub
        End If
      ElseIf how_work = "�O��㐧" Then
        If work_content <> "" And work_content <> "��" And work_content <> "�x" And work_content <> "��" And work_content <> "�[" Then
          'VBA���������I����
          Call high_speeding(False)
          '�G���[���b�Z�[�W���o�͂�
          MsgBox Names((i - 19) / 2 + 1) & "�����" & CStr(j - 11 + 1) & "���̋Ζ��`�����s�K�؂ł��I" & vbCrLf & "�u���v�u�x�v�u���v�u�[�v�̂ǂꂩ���L�ڂ��ĉ������I", _
          Buttons:=vbCritical, Title:="�ݒ�҂̕������̃G���[�\��"
          'Sub�v���W�[�W�����I������
          Exit Sub
        End If
      End If
      '��]��������Ă�����,
      If work_content <> "" Then
        request_table((i - 19) / 2, j - 11) = work_content
        hope_or_not((i - 19) / 2, j - 11) = 1
      Else
        request_table((i - 19) / 2, j - 11) = "�x"
        request_aroud_results_month(UBound(request_aroud_results_month, 1) - 1, (i - 19) / 2) = request_aroud_results_month(UBound(request_aroud_results_month, 1) - 1, (i - 19) / 2) + 1
        request_aroud_results_month(UBound(request_aroud_results_month, 1), (i - 19) / 2) = request_aroud_results_month(UBound(request_aroud_results_month, 1), (i - 19) / 2) + 1
        hope_or_not((i - 19) / 2, j - 11) = 0
      End If
      remarks_table((i - 19) / 2, j - 11) = ws.Cells(i + 1, j)
    Next
  Next
  
  'Debug.Print'
  Debug.Print vbCrLf
  
  Debug.Print "how_work:" & how_work
  
  Debug.Print "number_of_people:" & number_of_people
  
  Debug.Print "number_of_day:" & number_of_day
calculation_again_point:
  'making_work_table
  If how_work = "���㐧" Then
    For j = 0 To number_of_day - 1
      If request_aroud_results_day(0, j) < 0 Then '���ΐl�����s�����Ă�����,
        Call insufficient_compensation("��")
      End If
    Next
   
    For j = 0 To number_of_day - 1
      Call work_name_decide("��", "��")
      If request_aroud_results_day(1, j) > 0 Then
        While request_aroud_results_day(1, j) > 0 And day_work_classifying("��")(1) = "�L��"
         '1���̓������ݒ�l���𒴂��Ă���, ����, ����������΂ɕύX�\�ȕ�������ꍇ
          Call change_write
        Wend
      End If
    Next
    
    For j = 0 To number_of_day - 1
      If request_aroud_results_day(1, j) < 0 Then '�����l�����s�����Ă�����
        Call insufficient_compensation("��")
      End If
    Next
    
    For j = 0 To number_of_day - 1
      Call work_name_decide("��", "��")
      If request_aroud_results_day(1, j) > 0 Then
        While request_aroud_results_day(1, j) > 0 And day_work_classifying("��")(1) = "�L��"
         '1���̓������ݒ�l���𒴂��Ă���, ����, ����������΂ɕύX�\�ȕ�������ꍇ
          Call change_write
        Wend
      End If
    Next
    
    For j = 0 To number_of_day - 1
      Call work_name_decide("��", "�x")
      If request_aroud_results_day(0, j) > 0 Then
        While request_aroud_results_day(0, j) > 0 And day_work_classifying("��")(1) = "�L��"
         '1���̓��΂��ݒ�l���𒴂��Ă���, ����, ���΂���x�݂ɕύX�\�ȕ�������ꍇ
          Call change_write
        Wend
      End If
    Next
    
  ElseIf how_work = "�O��㐧" Then
    
    For j = 0 To number_of_day - 1
      If request_aroud_results_day(0, j) < 0 Then '1���̓��ΐl�����s�����Ă�����,
        Call insufficient_compensation("��")
      End If
    Next
    
    For j = 0 To number_of_day - 1
    
      Call work_name_decide("��", "��")
      While request_aroud_results_day(1, j) > 0 And day_work_classifying("��")(1) = "�L��"
      '1���̏���΂��ݒ�l���𒴂��Ă���, ����, ����΂�����΂ɕύX�\�ȕ�������ꍇ
        Call change_write
      Wend
      
      Call work_name_decide("��", "�x")
      If request_aroud_results_day(1, j) = 0 Then
      ElseIf day_work_classifying("��")(1) = "�L��" Then
        While request_aroud_results_day(1, j) > 0 And day_work_classifying("��")(1) = "�L��"
        '1���̏���΂��ݒ�l���𒴂��Ă���, ����, ����΂���x�݂ɕύX�\�ȕ�������ꍇ
          Call change_write
        Wend
      End If
      
      Call work_name_decide("�[", "��")
      While request_aroud_results_day(2, j) > 0 And day_work_classifying("�[")(1) = "�L��"
      '1���̐[��΂��ݒ�l���𒴂��Ă���, ����, �[��΂�����΂ɕύX�\�ȕ�������ꍇ
        Call change_write
      Wend
      
      Call work_name_decide("�[", "�x")
      If request_aroud_results_day(2, j) = 0 Then
      ElseIf day_work_classifying("�[")(1) = "�L��" Then
        While request_aroud_results_day(2, j) > 0 And day_work_classifying("�[")(1) = "�L��"
        '1���̏���΂��ݒ�l���𒴂��Ă���, ����, �[��΂���x�݂ɕύX�\�ȕ�������ꍇ
          Call change_write
        Wend
      End If
      
    Next
    
    For j = 0 To number_of_day - 1
      
      If request_aroud_results_day(1, j) < 0 Then '1���̏���ΐl�����s�����Ă�����,
        Call insufficient_compensation("��")
      End If
      If request_aroud_results_day(2, j) < 0 Then '1���̐[��ΐl�����s�����Ă�����,
        Call insufficient_compensation("�[")
      End If
    Next
    
    For j = 0 To number_of_day - 1
    
      Call work_name_decide("��", "��")
      While request_aroud_results_day(1, j) > 0 And day_work_classifying("��")(1) = "�L��"
      '1���̏���΂��ݒ�l���𒴂��Ă���, ����, ����΂�����΂ɕύX�\�ȕ�������ꍇ
        Call change_write
      Wend
      
      Call work_name_decide("��", "�x")
      If request_aroud_results_day(1, j) = 0 Then
      ElseIf day_work_classifying("��")(1) = "�L��" Then
        While request_aroud_results_day(1, j) > 0 And day_work_classifying("��")(1) = "�L��"
        '1���̏���΂��ݒ�l���𒴂��Ă���, ����, ����΂���x�݂ɕύX�\�ȕ�������ꍇ
          Call change_write
        Wend
      End If
      
      Call work_name_decide("�[", "��")
      While request_aroud_results_day(2, j) > 0 And day_work_classifying("�[")(1) = "�L��"
      '1���̐[��΂��ݒ�l���𒴂��Ă���, ����, �[��΂�����΂ɕύX�\�ȕ�������ꍇ
        Call change_write
      Wend
      
      Call work_name_decide("�[", "�x")
      If request_aroud_results_day(2, j) = 0 Then
      ElseIf day_work_classifying("�[")(1) = "�L��" Then
        While request_aroud_results_day(2, j) > 0 And day_work_classifying("�[")(1) = "�L��"
        '1���̏���΂��ݒ�l���𒴂��Ă���, ����, �[��΂���x�݂ɕύX�\�ȕ�������ꍇ
          Call change_write
        Wend
      End If
    Next
    
    For j = 0 To number_of_day - 1
      If request_aroud_results_day(1, j) < 0 Then '1���̏���ΐl�����s�����Ă�����,
        Call insufficient_compensation("��")
      End If
      If request_aroud_results_day(2, j) < 0 Then '1���̐[��ΐl�����s�����Ă�����,
        Call insufficient_compensation("�[")
      End If
    Next
    
    For j = 0 To number_of_day - 1
      Call work_name_decide("��", "�x")
      If request_aroud_results_day(0, j) > 0 Then
        While request_aroud_results_day(0, j) > 0 And day_work_classifying("��")(1) = "�L��"
         '1���̓��΂��ݒ�l���𒴂��Ă���, ����, ���΂���x�݂ɕύX�\�ȕ�������ꍇ
          Call change_write
        Wend
      End If
    Next
    
  End If
  
  '�Ζ��\�]��
    '1���̗]��E�s���l���ɑ΂���]��
  For i = 0 To UBound(request_aroud_results_day, 1)
    For j = 0 To number_of_day - 1
      If request_aroud_results_day(i, j) = 0 Then
        weight = 0
      Else
        If i = 0 Then
          If request_aroud_results_day(i, j) > 0 Then
            weight = 1
          ElseIf request_aroud_results_day(i, j) < 0 Then
            weight = 2
          End If
        ElseIf i = 1 Or i = 2 Then
          If request_aroud_results_day(i, j) > 0 Then
            weight = 3
          ElseIf request_aroud_results_day(i, j) < 0 Then
            weight = 4
          End If
        End If
      End If
      loss = loss + Abs(request_aroud_results_day(i, j)) * weight
    Next
  Next
  '���΂̌����v��loss�̕]���ɂ͊֌W�Ȃ��̂�, 1����n�߂�
  For i = 1 To UBound(request_aroud_results_month, 1)
    For j = 0 To number_of_people - 1
      If request_aroud_results_month(i, j) = 0 Then
        weight = 0
      Else
        If i = 2 Or (how_work = "�O��㐧" And i = 4) Or i = UBound(request_aroud_results_month, 1) Then
          If request_aroud_results_month(i, j) > 0 Then
            weight = 1
          ElseIf request_aroud_results_month(i, j) < 0 Then
            weight = 3
          End If
        End If
      End If
      loss = loss + Abs(request_aroud_results_month(i, j)) * weight
    Next
  Next
  Debug.Print "loss:" & loss
  Debug.Print "distinction_first_0:" & distinction_first_0
  If (Not distinction_first_0 And save_loss = 0) Or save_loss > loss Then
    save_loss = loss
    loss = 0
    save_request_for_evaluate = request_table
    If Not distinction_first_0 Then
      distinction_first_0 = True
    End If
    GoTo calculation_again_point
  Else
    request_table = save_request_for_evaluate
  End If
  
  '�������ݗp�V�[�g�w��
  Set ws = ThisWorkbook.Sheets("�`�[���Ԓ����O�����Ζ��\")
  '�V�[�g�ی�̉���
  Call pro_unpro(False, "�`�[���Ԓ����O�����Ζ��\")
  '�폜
  Range(ws.Cells(60 + 19, 11), ws.Cells(60 + 78, 41)).ClearContents
  '���������ʂׂ̍��ɂ���
  Range(ws.Cells(60 + 19, 11), ws.Cells(60 + 78, 41)).Font.Bold = False
  '�Ά����̂ɂ���
  Range(ws.Cells(60 + 19, 11), ws.Cells(60 + 78, 41)).Font.Italic = False
  '���̐F��ԁ����ɕς���
  Range(ws.Cells(60 + 19, 11), ws.Cells(60 + 78, 41)).Font.color = RGB(0, 0, 0)
  '��������
  For i = 60 + 19 To 60 + 19 + 2 * (number_of_people - 1) Step 2
    For j = 11 To 11 + number_of_day - 1
      '�Œ��]���ł͂Ȃ�, VBA��Ŏ����쐬�����Ζ��`���Ȃ��
      If hope_or_not((i - 60 - 19) / 2, j - 11) <> 1 Then
        '�����ɕς���
        ws.Cells(i, j).Font.Bold = True
        '�Α̂ɕς���
        ws.Cells(i, j).Font.Italic = True
        '���̐F��Ԃɕς���
        ws.Cells(i, j).Font.color = RGB(255, 0, 0)
      End If
      ws.Cells(i, j) = request_table((i - 60 - 19) / 2, j - 11)
      ws.Cells(i + 1, j) = remarks_table((i - 60 - 19) / 2, j - 11)
    Next
  Next
  
  '�Ζ��\�����쐬�ŏI�����X�V
  ws.Cells(21, 1) = "B�`�[��:" & Format(Now(), "YYYY�NMM��DD��HH��MM��")
  '�V�[�g�ی�̍ėL����
  Call pro_unpro(True, "�`�[���Ԓ����O�����Ζ��\")
  'VBA���������I����
  Call high_speeding(False)
  
  '�������b�Z�[�W��\������
  MsgBox "B�`�[���P�̂ł̋Ζ��\�̍쐬���������܂����I" _
      & vbCrLf & "�V�[�g�u�`�[���Ԓ����O�����Ζ��\�v�����m�F�������I", _
      Buttons:=vbInformation, Title:="���s�҂̕��ւ̃��b�Z�[�W"
  Exit Sub
   
ErrLabel:
  '�V�[�g�̕ی�̍ėL����
  Call pro_unpro(True, "�`�[���Ԓ����O�����Ζ��\")
  'VBA���������I������
  Call high_speeding(False)
  MsgBox "�uB�`�[���p�Ζ��\�����쐬���s�vVBA���s���ɃG���[���������܂����I" _
  & vbCrLf & "�V�[�g�uB�`�[���p�Ζ���]�\�v�����m�F�̏�, �ēx, ���s���ĉ������I" _
  & vbCrLf & "�G���[�����A�v��: " & Err.Source _
  & vbCrLf & "�G���[�ԍ�: " & Err.Number _
  & vbCrLf & "�G���[���e: " & Err.Description, _
  Buttons:=vbCritical, Title:="���s�҂̕������̃G���[�\��"
End Sub



