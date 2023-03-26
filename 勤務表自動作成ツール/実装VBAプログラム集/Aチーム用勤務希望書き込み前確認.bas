Attribute VB_Name = "Module1"
Option Explicit
Sub Aチーム用勤務希望書き込み前確認()
  Dim ws As Worksheet
  Dim number_of_people As Integer
  Dim check_people_num As Integer 'number_of_peopleのループ用の変数
  Dim check_cell_num As Integer
  Dim check_sub_title(1) As String
  check_sub_title(0) = "人数設定"
  check_sub_title(1) = "連続回数設定"
  Dim check_sub_title_num As Integer
  Dim check_title_for_three_change(1) As String
  Dim check_title_num_for_three_change As Integer
  
  On Error GoTo ErrLabel
  
'パラメータ設定の完了確認

  '読み込み用アクティブシート指定
  Set ws = ThisWorkbook.Sheets("Aチーム用勤務希望表")
  'name,number_of_people_reading
  For check_people_num = 19 To 19 + 2 * (30 - 1) Step 2 'For i=19 To 19+2*(30(人)-1) Step 2
    If ws.Cells(check_people_num, 5) <> "" Then
      number_of_people = number_of_people + 1
    End If
  Next
  
  '勤務制度の設定の完了確認
  If ws.Cells(1, 4) <> "二交代制" And ws.Cells(1, 4) <> "三交代制" Then '二交代制, 三交代制のどちらも入力されていない場合
    MsgBox "「勤務制度」に, 「二交代制」,又は,「三交代制」と入力して下さい ！", Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
    Exit Sub
  End If
  
  '日勤の人数設定の完了確認
  If ws.Cells(2, 4) = "" Then
    MsgBox "日勤の人数設定をして下さい！", _
    Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
    Exit Sub
  ElseIf VarType(ws.Cells(2, 4)) <> 5 Then
    MsgBox "日勤の人数設定の欄には数字を入力して下さい！" _
    & vbCrLf & "尚, 文字列で入力しないで下さい！", _
    Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
    Exit Sub
  '勤務表作成年月確認
  ElseIf Not ws.Cells(9, 4) Like "20*/*/*" Then
    MsgBox "何年何月の勤務表を作成するかをD4セルに入力して下さい！" _
    & vbCrLf & "尚, 形式は以下のように書いて下さい！" _
    & vbCrLf & "(例)2023(年:書いて下さい！)/1(月:書いて下さい！)/1(日:1のまま変更しないで下さい！)", _
    Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
    Exit Sub
  End If
  
  '二交代制のパラメータ設定の完了確認
  If ws.Cells(1, 4) = "二交代制" Then
    '1日1チーム当たりの当直の人数設定, 連続日数設定の完了確認
    For check_cell_num = 3 To 4
      If ws.Cells(check_cell_num, 4) = "" Then
        MsgBox "当直の" & check_sub_title(check_sub_title_num) & "をして下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Sub
      ElseIf VarType(ws.Cells(check_cell_num, 4)) <> 5 Then
        MsgBox "当直の" & check_sub_title(check_sub_title_num) & "の欄には数字を入力して下さい！" _
        & vbCrLf & "尚, 文字列で入力しないで下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Sub
      End If
      check_sub_title_num = check_sub_title_num + 1
    Next
    '1カ月の当直の最低回数の設定完了確認
    For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
      If ws.Cells(check_people_num, 6) = "" Then
        MsgBox ws.Cells(check_people_num, 5) & "さんの1カ月あたりの当直の最低回数を設定して下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Sub
      ElseIf VarType(ws.Cells(check_people_num, 6)) <> 5 Then
        MsgBox ws.Cells(check_people_num, 5) & "さんの1カ月あたりの当直の最低回数の欄には数字を入力して下さい！" _
        & vbCrLf & "尚, 文字列で入力しないで下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Sub
      End If
    Next
    
  '三交代制のパラメータ設定の完了確認
  ElseIf ws.Cells(1, 4) = "三交代制" Then
    check_title_for_three_change(0) = "準夜勤"
    check_title_for_three_change(1) = "深夜勤"
    '1日1チーム当たりの準(深)夜勤の人数設定, 連続日数設定の完了確認
    For check_cell_num = 5 To 8
      If ws.Cells(check_cell_num, 4) = "" Then
        MsgBox check_title_for_three_change(check_title_num_for_three_change) & "の" & check_sub_title(check_sub_title_num) & "をして下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Sub
      ElseIf VarType(ws.Cells(check_cell_num, 4)) <> 5 Then
        MsgBox check_title_for_three_change(check_title_num_for_three_change) & "の" & check_sub_title(check_sub_title_num) & "の欄には数字を入力して下さい！" _
        & vbCrLf & "尚, 文字列で入力しないで下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Sub
      End If
      check_sub_title_num = check_sub_title_num + 1
      If check_cell_num = 6 Then
        check_sub_title_num = 0
        check_title_num_for_three_change = check_title_num_for_three_change + 1
      End If
    Next
    '1カ月の準(深)夜勤の最低回数の設定完了確認
    check_title_num_for_three_change = 0
    For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
      For check_cell_num = 7 To 8
        If ws.Cells(check_people_num, check_cell_num) = "" Then
          MsgBox ws.Cells(check_people_num, 5) & "さんの1カ月あたりの" & check_title_for_three_change(check_title_num_for_three_change) & "の最低回数を設定して下さい！", _
          Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
          Exit Sub
        ElseIf VarType(ws.Cells(check_people_num, check_cell_num)) <> 5 Then
          MsgBox ws.Cells(check_people_num, 5) & "さんの1カ月あたりの" & check_title_for_three_change(check_title_num_for_three_change) & "の最低回数の欄には数字を入力して下さい！" _
          & vbCrLf & "尚, 文字列で入力しないで下さい！", _
          Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
          Exit Sub
        End If
        check_title_num_for_three_change = check_title_num_for_three_change + 1
      Next
      check_title_num_for_three_change = 0
    Next
  End If
 '個人の1カ月あたりの休みの設定完了確認
  
  For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
    If ws.Cells(check_people_num, 9) = "" Then
      MsgBox ws.Cells(check_people_num, 5) & "さんの1カ月あたりの休みの最低回数を設定して下さい！", _
      Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
      Exit Sub
    ElseIf VarType(ws.Cells(check_people_num, 9)) <> 5 Then
      MsgBox ws.Cells(check_people_num, 5) & "さんの1カ月あたりの休みの最低回数の欄には数字を入力して下さい！" _
      & vbCrLf & "尚, 文字列で入力しないで下さい！", _
      Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
      Exit Sub
    End If
  Next
  
  MsgBox "これでAチーム勤務希望表の準備完了です！" _
      & vbCrLf & "希望を書き込んだ後, 勤務表自動作成を実行して下さい！", _
      Buttons:=vbInformation, Title:="設定者の方へのメッセージ"
  Exit Sub
  
ErrLabel:
  MsgBox "「Aチーム用勤務希望書き込み前確認」VBA実行中にエラーが発生しました！" _
  & vbCrLf & "シート「Aチーム用勤務希望表」をご確認の上, 再度, 実行して下さい！" _
  & vbCrLf & "エラー発生アプリ: " & Err.Source _
  & vbCrLf & "エラー番号: " & Err.Number _
  & vbCrLf & "エラー内容: " & Err.Description, _
  Buttons:=vbCritical, Title:="実行者の方向けのエラー表示"
End Sub
