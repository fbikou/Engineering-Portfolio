Attribute VB_Name = "Module7"
Option Explicit
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
Private Sub Auto_Open()

    Dim Write_Sheet(4) As String
    Dim sheet_name As Variant
    Dim ws As Worksheet
    Dim QT As QueryTable
    Write_Sheet(0) = "A�`�[���p�Ζ���]�\"
    'Write_Sheet(1) = "B�`�[���p�Ζ���]�\"
    'Write_Sheet(2) = "�`�[���Ԓ����O�����Ζ��\"
    'Write_Sheet(3) = "��]�D��`�[���Ԓ����㑍���Ζ��\"
    'Write_Sheet(4) = "��]��񂵃`�[���Ԓ����㑍���Ζ��\"
    Dim eligible_year As String
    Dim wrong_holidays_sheets As New Collection
    Dim msg_for_sheet_name As String
    
    On Error GoTo ErrLabel
    
    'VBA���������J�n����
    Call high_speeding(True)
      
    '�ݒ肵�Ă���V�[�g���̏j���̑Ώ۔N��S�Ċm�F��, �Ⴄ�N���Ώۂ������ꍇ, ���̃V�[�g����ǉ�����
    For Each sheet_name In Write_Sheet
      If sheet_name <> "" Then
        Set ws = ThisWorkbook.Sheets(sheet_name)
        eligible_year = ws.Cells(160, 53)
        If CStr(Year(Now)) <> Left(eligible_year, InStr(eligible_year, "�N") - 1) Then
          wrong_holidays_sheets.Add sheet_name
        End If
      End If
    Next
  
    '�ݒ肵�Ă���S�ẴV�[�g���̏j�����u�b�N���J�������̔N�̏j����������
    If wrong_holidays_sheets.count = 0 Then
      'VBA���������I����
      Call high_speeding(False)
      '�j���̎����X�V���֌W�Ȃ��n�_�܂ňړ�����
      GoTo holiday_end_point
    End If
    
    '�ȍ~, '�ݒ肵�Ă���S�ẴV�[�g��1�V�[�g�ȏ�, �j�����u�b�N���J�������̔N�̏j���ƈ�����ꍇ�̂ݎ��s����
    For Each sheet_name In wrong_holidays_sheets
      '�V�[�g�̕ی����������
      Call pro_unpro(False, sheet_name)
      Set ws = ThisWorkbook.Sheets(sheet_name)
      '�L�ڂ���j���̑Ώ۔N�̏�������
      ws.Cells(159, 53) = "�Ώ۔N"
      ws.Cells(160, 53) = CStr(Year(Now)) & "�N"
      ws.Range(ws.Cells(159, 53), ws.Cells(160, 53)).Borders.LineStyle = xlContinuous
      '�j�����擾��, ��������
      Set QT = ws.QueryTables.Add _
      (Connection:="URL;https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html", Destination:=ws.Range("BA162")) '�j���̃f�[�^��������t�{��HP��URL���w��
      With QT
        '�s���ǉ������Z�����㏑������ݒ�ɂ��Ă���
        .RefreshStyle = xlOverwriteCells
        '.AdjustColumnWidth = True
        .WebFormatting = xlWebFormattingNone
        .WebSelectionType = xlSpecifiedTables
        .WebTables = 1
        .Refresh BackgroundQuery:=False
        .Delete
      End With
      ws.Range(ws.Cells(162, 53), ws.Cells(189, 55)).Borders.LineStyle = xlContinuous
      ws.Columns("BA:BC").Columns.AutoFit
      '�V�[�g�̕ی���ēx,�L���ɂ���
      Call pro_unpro(True, sheet_name)
    Next
    
    '���b�Z�[�W�p�ɖ��X�V���X�V�����V�[�g�̖��O�𕶎���Ƃ��Ēǉ�����
    'For Each sheet_name In wrong_holidays_sheets
    '  msg_for_sheet_name = msg_for_sheet_name & vbCrLf & sheet_name
    'Next
    
    'VBA���������I������
    Call high_speeding(False)
    MsgBox "�N�ԏj���������X�V���܂����I", Buttons:=vbInformation, Title:="�t�@�C�����J�������ւ̃��b�Z�[�W"
    'MsgBox "�ȉ��̃V�[�g�ł̔N�ԏj���������X�V���܂����I" & msg_for_sheet_name, Buttons:=vbInformation, Title:="�t�@�C�����J�������ւ̃��b�Z�[�W"
    
holiday_end_point:
   
    If Application.CellDragAndDrop = True Then
      Application.CellDragAndDrop = False
    End If
    
    Exit Sub
    
ErrLabel:
  For Each sheet_name In Write_Sheet
    '�V�[�g�̕ی���ēx,�L���ɂ���
    Call pro_unpro(True, sheet_name)
  Next
  'VBA���������I������
  Call high_speeding(False)
  MsgBox "�u�N�ԏj�������X�V�v(Auto_Open)VBA���s���ɃG���[���������܂����I" _
  & vbCrLf & "�V�[�g�uA�`�[���p�Ζ���]�\�v�́u�j��, ���t��v�m�F���v�����m�F�̏�, ����Excel�t�@�C�����u�㏑���ۑ��v��, �ēx, �J���ĉ������I" _
  & vbCrLf & "��, ���ɕ�����A�����Ă���̂�, �{�v���O����������ɋ@�\���Ȃ��ꍇ��, �N�ԏj���̎蓮�X�V���������������B" _
  & vbCrLf & "�G���[�����A�v��: " & Err.Source _
  & vbCrLf & "�G���[�ԍ�: " & Err.Number _
  & vbCrLf & "�G���[���e: " & Err.Description, _
  Buttons:=vbCritical, Title:="�t�@�C�����J�����������̃G���[�\��"
End Sub
