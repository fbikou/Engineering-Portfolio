Attribute VB_Name = "Module6"
Option Explicit
Private offset As Integer
Private new_sheet_name As String
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
Function read_A_B_change(ByVal A_or_B As String) As Boolean
  If A_or_B = "A" Then
    offset = 19
  Else
    If offset = 19 Then
      offset = 79
      read_A_B_change = True
    End If
  End If
End Function
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
Function book_pro_unpro(ByVal pro_or_unpro As Boolean, Optional ByVal passward As String = "tokubetunatoki")
  If pro_or_unpro Then
    If ThisWorkbook.ProtectStructure = False Then
      ThisWorkbook.Protect Password:=passward
    End If
  Else
    If ThisWorkbook.ProtectStructure = True Then
      ThisWorkbook.Unprotect Password:=passward
    End If
  End If
End Function
Function exist_sht(ByVal sht_name As String)
    Dim ws As Variant
    For Each ws In Sheets
        If ws.Name = sht_name Then
            exist_sht = True ' ���݂���
            Exit Function
        End If
    Next
    ' ���݂��Ȃ�
    exist_sht = False
End Function
Function rng_new_sht_pst(ByVal new_sheet_name As String, ByVal copy_rng As Range, ByVal pst_rng_left_top_row As Integer, ByVal pst_rng_left_top_col As Integer)
  Dim new_sht_name As String: new_sht_name = new_sheet_name
  Dim ws_write As Worksheet
  Dim pst_row_bottom As Integer
  Dim pst_col_right As Integer
  Dim pst_rng As Range
  Dim screen_stop_judge  As Boolean
  Dim i As Integer
  Dim color As Long
  Dim fc As FormatCondition
  If Not exist_sht(new_sht_name) Then
    Set ws_write = Worksheets.Add(After:=ThisWorkbook.Worksheets(Sheets.count))
    ws_write.Name = new_sht_name
  End If
  copy_rng.Copy
  Set ws_write = ThisWorkbook.Worksheets(new_sheet_name)
  ws_write.Cells(pst_rng_left_top_row, pst_rng_left_top_col).PasteSpecial Paste:=xlPasteAll
  pst_row_bottom = pst_rng_left_top_row + copy_rng.Rows.count - 1
  pst_col_right = pst_rng_left_top_col + copy_rng.Columns.count - 1
  Set pst_rng = ws_write.Range(ws_write.Cells(pst_rng_left_top_row, pst_rng_left_top_col), ws_write.Cells(pst_row_bottom, pst_col_right))
  '�������l�ɕϊ�
  pst_rng.Value = copy_rng.Value
  If pst_rng_left_top_row = 1 And pst_rng_left_top_col = 3 Then
    '�����t�������̍폜
    pst_rng.FormatConditions.Delete
    '�w�i�F�̒ǉ�
    For i = 2 To 32
      color = copy_rng(2, i).DisplayFormat.Interior.color
      pst_rng.Cells(2, i).Interior.color = color
      pst_rng.Cells(3, i).Interior.color = color
    Next
  End If
  '�Ζ��\��̏����t�������̕ύX
  If pst_rng.FormatConditions.count > 0 Then
    Set fc = pst_rng.FormatConditions(1)
    If fc.Type = xlExpression Then
      If fc.Formula1 = "=MOD(ROW(),2)=0" Then
        pst_rng.FormatConditions.Delete
        Set fc = pst_rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=MOD(ROW(),2)=1")
        fc.Interior.color = RGB(255, 255, 255)
        Set fc = pst_rng.FormatConditions.Add(Type:=xlTextString, String:="��", TextOperator:=xlContains)
        fc.Interior.color = RGB(255, 192, 0)
        Set fc = pst_rng.FormatConditions.Add(Type:=xlTextString, String:="��", TextOperator:=xlContains)
        fc.Interior.color = RGB(255, 255, 0)
        Set fc = pst_rng.FormatConditions.Add(Type:=xlTextString, String:="��", TextOperator:=xlContains)
        fc.Interior.color = RGB(255, 255, 0)
        Set fc = pst_rng.FormatConditions.Add(Type:=xlTextString, String:="��", TextOperator:=xlContains)
        fc.Interior.color = RGB(146, 208, 80)
        Set fc = pst_rng.FormatConditions.Add(Type:=xlTextString, String:="�[", TextOperator:=xlContains)
        fc.Interior.color = RGB(0, 176, 240)
      End If
    End If
  End If
  pst_rng.EntireColumn.AutoFit
  pst_rng.EntireRow.AutoFit
End Function
Function pic_save(ByVal pic_rng As Range, ByVal pic_name As String, Optional ByVal dir_path As String = "", Optional ByVal sht_name As String)
  Dim sheet_name As String: sheet_name = sht_name
  Dim FileSize As Long
  Dim pic As ChartObject
  Dim picName As String: picName = "\" & pic_name & ".jpg"
  Dim pic_path As String
  Dim event_stop_judge As Boolean
  If Application.EnableEvents = False Then
   event_stop_judge = True
   Application.EnableEvents = True
  End If
  If dir_path = "" Then
    pic_path = ThisWorkbook.Path & picName
  Else
    pic_path = dir_path & picName
  End If
  Dim ws_pic As Worksheet
  Set ws_pic = ThisWorkbook.Worksheets(sheet_name)
  Dim rng As Range: Set rng = pic_rng
  Dim copyRetryCount As Integer: copyRetryCount = 10 '���g���C��
  '�G���[���̓��g���C�����ɔ��(�R�s�[���ɃG���[����������z��̈�, CopyRetry�Ƃ������O�ɂ��Ă���)
  On Error GoTo CopyRetry
  
  '���Z���͈͂��摜�f�[�^�ŃR�s�[����B
  rng.CopyPicture
  
  '���w�肵���Z���͈͂Ɠ����T�C�Y��pic��V�K�쐬���A�ۑ�����B
  Set pic = ws_pic.ChartObjects.Add(0, 0, rng.Width, rng.Height)
  pic.chart.Export pic_path
  FileSize = FileLen(pic_path)
  
  '��pic��FileSize�𒴂���܂Ń��[�v����(�摜�f�[�^���o���オ������I������)
  Do Until FileLen(pic_path) > FileSize
   
    pic.chart.Paste
    pic.chart.Export pic_path
    DoEvents
  Loop
  '���쐬������Apic�폜�B
  pic.Delete
  Set pic = Nothing
  
  If event_stop_judge = True Then
    Application.EnableEvents = False
  End If
  
  Exit Function '���g���C�����ɔ�ԑO�Ƀ��\�b�h���甲����
  
'���g���C����
CopyRetry:
    '��莞�ԑҋ@��A�G���[���O�̏����ɔ��
    copyRetryCount = copyRetryCount - 1
    If copyRetryCount >= 1 Then
        '�c�胊�g���C�񐔂�1���傫���ꍇ�́A�ēx, ���s����
        '100�~���b �o�ߌ�Ď��s
      Application.Wait [Now()] + 100 / 86400000
      DoEvents
      Resume
    End If
  
End Function
Function imgFolderDelete(ByVal dirPath As String)
On Error GoTo deleteErr
  
  Dim objFSO As Object
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Dim strFolderPath As String
  strFolderPath = dirPath
  
  '�t�H���_�폜
  objFSO.DeleteFolder strFolderPath
  
  '�t�H���_���݃`�F�b�N
  If Dir(strFolderPath, vbDirectory) = "" Then
    MsgBox "VBA�̃G���[�ɂ�芮�����Ȃ������Ζ��\�摜�W�̃t�H���_-�폜�ɐ������܂���!", _
         Buttons:=vbInformation, Title:="�������̋Ζ��\�摜�W�̃t�H���_-�폜�����̂���"
  End If
  
  Exit Function
  
deleteErr:
  MsgBox "�������Ȃ������Ζ��\�摜�W�̃t�H���_-�폜���ɃG���[��������, �폜�Ɏ��s���܂���..." & vbCrLf & _
         "�G���[�����͈ȉ��̒ʂ�ł��B" & vbCrLf & vbCrLf & _
         "----------------------------------" & vbCrLf & _
         "�G���[�ԍ��F" & Err.Number & vbCrLf & _
         "�G���[�ڍׁF" & Err.Description & vbCrLf & _
         "----------------------------------", _
         Buttons:=vbCritical, Title:="���s�҂̕������̃G���[�\��"
 
End Function
Sub �Ζ��\�摜�ۑ�()
  offset = 0
  new_sheet_name = "�Ζ��\�摜�ۑ��p�V�[�g"
  Dim back_up_dir_path As String
  Dim sheet_names(2) As String
  Dim picture_range(3 - 1, 2 - 1, 2 - 1) As Range
  Dim picture_name As String
  Dim ws As Worksheet
  Dim common_rng(1) As Range
  Dim table_date As String
  Dim sht_name_index As Integer
  Dim sht_name As Variant
  Dim AB As Integer
  Dim name_schedule As Integer
  Dim number_of_people As Integer
  Dim i As Integer
  Dim team As String
  Dim write_offset As Integer
  Dim active_ws As Worksheet
  Dim err_non_people As Boolean
  Dim err_team_name As String
  sheet_names(0) = "�`�[���Ԓ����O�����Ζ��\"
  sheet_names(1) = "��]�D��`�[���Ԓ����㑍���Ζ��\"
  sheet_names(2) = "��]��񂵃`�[���Ԓ����㑍���Ζ��\"
  
  'On Error GoTo ErrLabel
  
  '�A�N�e�B�u�V�[�g�̎擾
  Set active_ws = ActiveSheet
 
  'VBA�������J�n(���Ɍx�����b�Z�[�W�̒�~)
  Call high_speeding(True)
  '�u�b�N�̕ی�̉���
  Call book_pro_unpro(False)
  '�V�[�g�̕ی�̉���
  For Each sht_name In sheet_names
    Call pro_unpro(False, sht_name)
  Next
  
  '���ʕ����̑��
    '�Ζ��\�Ώی��̓ǂݎ��
  Set ws = ThisWorkbook.Worksheets("�`�[���Ԓ����O�����Ζ��\")
  table_date = ws.Cells(9, 4)
  table_date = Left(table_date, Len(table_date) - 3)
  table_date = Replace(table_date, "/", "�N") & "��"
    '�\�̏�̋��ʕ����͈̔͊i�[
  Set common_rng(0) = ws.Range(ws.Cells(16, 4), ws.Cells(18, 5))
  Set common_rng(1) = ws.Range(ws.Cells(16, 10), ws.Cells(18, 41))
  
  '�V�[�g���ɈقȂ�Z�����e(�͈�)�̓ǂݎ��
  For sht_name_index = 0 To 2
    'sht_AB_loop_num=1,2�̎���sheet_name(0)���w�肷��悤�ɂ���
    Set ws = ThisWorkbook.Worksheets(sheet_names(sht_name_index))
    'name,number_of_people_reading
    Call read_A_B_change("A")
B_team_return_point_num_people:
    For i = offset To offset + 2 * (30 - 1) Step 2
      If ws.Cells(i, 5) <> "" Then
        number_of_people = number_of_people + 1
      End If
    If number_of_people = 0 Then
      err_non_people = True
      GoTo ErrLabel
    End If
    Next
    If offset = 19 Then
      Set picture_range(sht_name_index, 0, 0) = ws.Range(ws.Cells(offset, 5), ws.Cells(offset + number_of_people * 2 - 1, 5))
      Set picture_range(sht_name_index, 0, 1) = ws.Range(ws.Cells(offset, 10), ws.Cells(offset + number_of_people * 2 - 1, 41))
    Else
      Set picture_range(sht_name_index, 1, 0) = ws.Range(ws.Cells(offset, 5), ws.Cells(offset + number_of_people * 2 - 1, 5))
      Set picture_range(sht_name_index, 1, 1) = ws.Range(ws.Cells(offset, 10), ws.Cells(offset + number_of_people * 2 - 1, 41))
    End If
    If read_A_B_change("B") Then
      'A�`�[��(��̗�)�̐l���𐔂�����, �Z���͈͂��i�[���Ă���
      number_of_people = 0
      GoTo B_team_return_point_num_people
    End If
    number_of_people = 0
  Next
    
  '�o�b�N�A�b�v�t�H���_�̃p�X���쐬����(�o�b�N�A�b�v�p�̃t�H���_��I�������t�H���_�Ɠ����K�w�ɍ쐬�����)
  back_up_dir_path = ThisWorkbook.Path & "\" & table_date & "�p�Ζ��\�摜�W_" & Format(Now(), "YYYY�NMM��DD��HH��MM��SS�b") & "�쐬"
  MkDir (back_up_dir_path)
    
  '�\��t���Ɖ摜�ۑ�
  ThisWorkbook.Worksheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
  Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
  ws.Name = new_sheet_name
  For sht_name_index = 0 To 2
    Call rng_new_sht_pst(new_sheet_name, common_rng(0), 1, 1)
    Call rng_new_sht_pst(new_sheet_name, common_rng(1), 1, 3)
    Set ws = ThisWorkbook.Worksheets(new_sheet_name)
    For AB = 0 To 1
      If AB = 0 Then
        team = "A�`�[��"
        write_offset = 3
      Else
        team = "B�`�[��"
        write_offset = 3 + picture_range(sht_name_index, 0, 0).Rows.count
      End If
      ws.Range(ws.Cells(write_offset + 1, 1), ws.Cells(write_offset + picture_range(sht_name_index, AB, 0).Rows.count, 1)).Merge
      ws.Range(ws.Cells(write_offset + 1, 1), ws.Cells(write_offset + picture_range(sht_name_index, AB, 0).Rows.count, 1)).BorderAround weight:=xlMedium, LineStyle:=xlContinuous
      ws.Cells(write_offset + 1, 1) = team
      For name_schedule = 0 To 1
        Call rng_new_sht_pst(new_sheet_name, picture_range(sht_name_index, AB, name_schedule), write_offset + 1, 2 + name_schedule)
        If name_schedule = 1 Then
          ws.Range(ws.Cells(write_offset + 1, 2), ws.Cells(write_offset + picture_range(sht_name_index, AB, 0).Rows.count, 34)).BorderAround weight:=xlMedium, LineStyle:=xlContinuous
        End If
      Next
    Next
    Call pic_save(ws.Range(ws.Cells(1, 1), ws.Cells(3 + picture_range(sht_name_index, 0, 0).Rows.count + picture_range(sht_name_index, 1, 0).Rows.count, 34)), _
    table_date & "�p" & sheet_names(sht_name_index) & "_" & Format(Now(), "YYYY�NMM��DD��HH��MM��SS�b") & "�쐬", back_up_dir_path, _
    new_sheet_name)
    ws.Delete
  Next
  
  '�V�[�g�̕ی�̍ėL����
  For Each sht_name In sheet_names
    Call pro_unpro(True, sht_name)
  Next
  '�u�b�N�̕ی�̍ėL����
  Call book_pro_unpro(True)
  'VBA�������I��(���Ɍx�����b�Z�[�W�̍ĊJ)
  Call high_speeding(False)
  
  '���s��̃A�N�e�B�u�V�[�g�����s�O�̃A�N�e�B�u�V�[�g�ƈقȂ����ꍇ,
  '���s�O�̃A�N�e�B�u�V�[�g���ēx , �A�N�e�B�u�ɂ���
  If active_ws.Name <> ActiveSheet.Name Then
    active_ws.Activate
  End If
  
  If Dir(back_up_dir_path, vbDirectory) <> "" Then
    MsgBox "�Ζ��\�摜�ۑ��������܂����I" & vbCrLf & "���̉�ʂł��m�F�������I", Buttons:=vbInformation, Title:="���s�҂̕��ւ̃��b�Z�[�W"
    '�t�H���_���J��
    Shell "C:\Windows\Explorer.exe " & back_up_dir_path, vbNormalFocus
  End If
  
  Exit Sub
  
ErrLabel:
  '�Ζ��\�摜�ۑ��p�V�[�g����������, �폜����
  If exist_sht(new_sheet_name) Then
    ThisWorkbook.Worksheets(new_sheet_name).Delete
  End If
   '�V�[�g�̕ی�̍ėL����
  For Each sht_name In sheet_names
    Call pro_unpro(True, sht_name)
  Next
  '�u�b�N�̕ی�̍ėL����
  Call book_pro_unpro(True)
  'VBA�������I��(���Ɍx�����b�Z�[�W�̍ĊJ)
  Call high_speeding(False)
  '���s��̃A�N�e�B�u�V�[�g�����s�O�̃A�N�e�B�u�V�[�g�ƈقȂ����ꍇ,
  '���s�O�̃A�N�e�B�u�V�[�g���ēx , �A�N�e�B�u�ɂ���
  If active_ws.Name <> ActiveSheet.Name Then
    active_ws.Activate
  End If
  If Not err_non_people Then
    MsgBox "�u�Ζ��\�摜�ۑ��vVBA���s���ɃG���[���������܂����I" _
    & vbCrLf & "�ēx, �u�Ζ��\�摜�ۑ��v�̃{�^��������, ���s���Ă݂ĉ�����!" _
    & vbCrLf & "���ɕ�����, ���s�����̂�, �摜�ۑ����o���Ȃ��ꍇ��, �X�N���[���V���b�g�ɂ��Ζ��\�摜�ۑ����������������B" _
    & vbCrLf & "�G���[�����A�v��: " & Err.Source _
    & vbCrLf & "�G���[�ԍ�: " & Err.Number _
    & vbCrLf & "�G���[���e: " & Err.Description, _
    Buttons:=vbCritical, Title:="���s�҂̕������̃G���[�\��"
  Else
    If offset = 19 Then
      err_team_name = "A�`�[��"
    ElseIf offset = 79 Then
      err_team_name = "B�`�[��"
    End If
    MsgBox "�ȉ��̎菇�@�`�D���s���ĉ������I" _
    & vbCrLf & "�@" & err_team_name & "�̋Ζ��҂̎������u" & err_team_name & "�p�Ζ���]�\�v�œ��͂��� " _
    & vbCrLf & "�A" & "�u" & err_team_name & "�p�Ζ��\�����쐬���s�v�{�^��������" _
    & vbCrLf & "�B" & "�u�`�[���Ԓ����O�����Ζ��\�v������������" _
    & vbCrLf & "�C" & "�u�`�[���Ԓ����Ζ��\�����쐬���s�v�{�^��������," _
    & vbCrLf & "�@�u��]�D��(���)�����Ζ��\�v������������" _
    & vbCrLf & "�D" & "���́u�Ζ��\�摜�ۑ��v�{�^�����ēx, ����", _
    Buttons:=vbCritical, Title:="���s�҂̕��ւ̃��b�Z�[�W"
  End If
  'back_up_dir_path��, �������ꂽ�t�H���_�[�����������Ă���
  If back_up_dir_path <> "" Then
    '�摜�ۑ��p�̃t�H���_����������, �폜����
    If Dir(back_up_dir_path, vbDirectory) <> "" Then
      Call imgFolderDelete(back_up_dir_path)
    End If
  End If
End Sub
  


