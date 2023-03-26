Attribute VB_Name = "Module4"
Option Explicit
'Private変数の定義(同一モジュール内ではPublic変数と同じ様に関数をまたいで使える)
  'シート指定用変数
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
'ループ変数の定義
Private j As Integer 'number_of_dayのループ用のj
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
Function high_speeding(ByVal start_end As Boolean) 'VBA開始の時はTrue, 終了の時はFalseでVBA高速化を開始・終了する
  If start_end = True Then
    ' [VBA高速化]
    '(上書き保存の警告メッセージを含む)警告メッセージを無視するように設定する
    Application.DisplayAlerts = False
    '描画を停止する
    Application.ScreenUpdating = False
    'イベントを抑制する
    Application.EnableEvents = False
    'ステータスバーを無効にする
    Application.DisplayStatusBar = False
    '自動計算の停止
    Application.Calculation = xlCalculationManual
  Else
    ' [VBA高速化]
    '警告メッセージを表示させるように直す
    Application.DisplayAlerts = True
    '描画を再開する
    Application.ScreenUpdating = True
     'イベントの抑制を解放する
    Application.EnableEvents = True
    'ステータスバーを有効にする
    Application.DisplayStatusBar = True
    '自動計算再開
    Application.Calculation = xlCalculationAutomatic
  End If
End Function
Function Before_building_check() As String
  Before_building_check = "設定不足, エラー, 日数不足有り"
'Function Before_building_check内の変数の定義
  Dim check_people_num As Integer 'number_of_peopleのループ用の変数
  Dim check_cell_num As Integer
  Dim check_sub_title(1) As String
  check_sub_title(0) = "人数設定"
  check_sub_title(1) = "連続回数設定"
  Dim check_sub_title_num As Integer
  Dim check_title_for_three_change(1) As String
  Dim check_title_num_for_three_change As Integer
'パラメータ設定の完了確認

  '勤務制度の設定の完了確認
  If ws.Cells(1, 4) <> "二交代制" And ws.Cells(1, 4) <> "三交代制" Then '二交代制, 三交代制のどちらも入力されていない場合
    MsgBox "「勤務制度」に, 「二交代制」,又は,「三交代制」と入力して下さい ！", Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
    Exit Function
  End If
  
  '日勤の人数設定の完了確認
  If ws.Cells(2, 4) = "" Then
    MsgBox "日勤の人数設定をして下さい！", _
    Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
    Exit Function
  ElseIf VarType(ws.Cells(2, 4)) <> 5 Then
    MsgBox "日勤の人数設定の欄には数字を入力して下さい！" _
    & vbCrLf & "尚, 文字列で入力しないで下さい！", _
    Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
    Exit Function
  '勤務表作成年月確認
  ElseIf Not ws.Cells(9, 4) Like "20*/*/*" Then
    MsgBox "何年何月の勤務表を作成するかをD4セルに入力して下さい！" _
    & vbCrLf & "尚, 形式は以下のように書いて下さい！" _
    & vbCrLf & "(例)2023(年:書いて下さい！)/1(月:書いて下さい！)/1(日:1のまま変更しないで下さい！)", _
    Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
    Exit Function
  End If
  
  '二交代制のパラメータ設定の完了確認
  If ws.Cells(1, 4) = "二交代制" Then
    '1日1チーム当たりの当直の人数設定, 連続日数設定の完了確認
    For check_cell_num = 3 To 4
      If ws.Cells(check_cell_num, 4) = "" Then
        MsgBox "当直の" & check_sub_title(check_sub_title_num) & "をして下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Function
      ElseIf VarType(ws.Cells(check_cell_num, 4)) <> 5 Then
        MsgBox "当直の" & check_sub_title(check_sub_title_num) & "の欄には数字を入力して下さい！" _
        & vbCrLf & "尚, 文字列で入力しないで下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Function
      End If
      check_sub_title_num = check_sub_title_num + 1
    Next
    '1カ月の当直の最低回数の設定完了確認
    For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
      If ws.Cells(check_people_num, 6) = "" Then
        MsgBox Names((check_people_num - 19) / 2 + 1) & "さんの1カ月あたりの当直の最低回数を設定して下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Function
      ElseIf VarType(ws.Cells(check_people_num, 6)) <> 5 Then
        MsgBox Names((check_people_num - 19) / 2 + 1) & "さんの1カ月あたりの当直の最低回数の欄には数字を入力して下さい！" _
        & vbCrLf & "尚, 文字列で入力しないで下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Function
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
        Exit Function
      ElseIf VarType(ws.Cells(check_cell_num, 4)) <> 5 Then
        MsgBox check_title_for_three_change(check_title_num_for_three_change) & "の" & check_sub_title(check_sub_title_num) & "の欄には数字を入力して下さい！" _
        & vbCrLf & "尚, 文字列で入力しないで下さい！", _
        Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
        Exit Function
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
          MsgBox Names((check_people_num - 19) / 2 + 1) & "さんの1カ月あたりの" & check_title_for_three_change(check_title_num_for_three_change) & "の最低回数を設定して下さい！", _
          Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
          Exit Function
        ElseIf VarType(ws.Cells(check_people_num, check_cell_num)) <> 5 Then
          MsgBox Names((check_people_num - 19) / 2 + 1) & "さんの1カ月あたりの" & check_title_for_three_change(check_title_num_for_three_change) & "の最低回数の欄には数字を入力して下さい！" _
          & vbCrLf & "尚, 文字列で入力しないで下さい！", _
          Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
          Exit Function
        End If
        check_title_num_for_three_change = check_title_num_for_three_change + 1
      Next
      check_title_num_for_three_change = 0
    Next
  End If
 '個人の1カ月あたりの休みの設定完了確認
  
  For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
    If ws.Cells(check_people_num, 9) = "" Then
      MsgBox Names((check_people_num - 19) / 2 + 1) & "さんの1カ月あたりの休みの最低回数を設定して下さい！", _
      Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
      Exit Function
    ElseIf VarType(ws.Cells(check_people_num, 9)) <> 5 Then
      MsgBox Names((check_people_num - 19) / 2 + 1) & "さんの1カ月あたりの休みの最低回数の欄には数字を入力して下さい！" _
      & vbCrLf & "尚, 文字列で入力しないで下さい！", _
      Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
      Exit Function
    End If
  Next
  
'エラー確認
  For check_people_num = 19 To 19 + 2 * (number_of_people - 1) Step 2
    If ws.Cells(check_people_num, 52) <> "" Then
      MsgBox Names((check_people_num - 19) / 2 + 1) & "さんの希望欄でエラーが発生しています！", _
      Buttons:=vbCritical, Title:="希望入力者の方向けのエラー表示"
      Exit Function
    End If
  Next
  
  Before_building_check = "設定不足, エラー, 日数不足無し"
End Function
Function excess_miss(ByVal day_num_excess_miss As String) As Integer '「不足している」時は-で返し, 「余っている」時は+で, 「(指定日数に)等しい」時は0で返す
  If InStr(day_num_excess_miss, "不足しています") > 0 Then
    excess_miss = -Val(Left(day_num_excess_miss, InStr(day_num_excess_miss, "日") - 1))
  ElseIf InStr(day_num_excess_miss, "余っています") > 0 Then
    excess_miss = Val(Left(day_num_excess_miss, InStr(day_num_excess_miss, "日") - 1))
  ElseIf InStr(day_num_excess_miss, "等しい") > 0 Then
    excess_miss = 0
  End If
End Function
Function after_find_function(ByVal classify_num As String, ByVal change_need_value As Integer)
  max_person_num = classify_num
  max_finding = change_need_value
  fit_object_exist = "有り"
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
  'この関数は「二交代制」の場合のみに使い, 日勤, 休みを「当直→明け→休み」に変える時に「明け→休み」の部分を実現できるか確認する関数である
  '指定日がその月に含まれない(つまり, 月末より後ろ)か, 又は, その月に含まれるが, その後, 二日間の勤務形式は固定希望ではなく, かつ, 日勤は人数が余っていたら,Trueを返す
  Dim Another_size_out As Boolean
  Dim hope_or_not_judge As Boolean
  Dim Surplus_or_equal As Boolean
  If Another_Day_Num > Another_compare_num Then
    Another_size_out = True
  Else
    Another_size_out = False
  End If
  If Another_size_out = False Then
    '固定希望日かの確認
    If hope_or_not(person_num, Another_Day_Num) <> 1 Then
      hope_or_not_judge = False
    Else
      hope_or_not_judge = True
    End If
    'その日の勤務形式が日勤なら, その日の日勤の合計人数が余っているかの確認
    'その日の勤務形式が休みなら, その日の日勤の合計人数が不足していないかの確認
    If (request_table(person_num, Another_Day_Num) = "日" And request_aroud_results_day(0, Another_Day_Num) > 0) Or _
    (request_table(person_num, Another_Day_Num) = "休" And request_aroud_results_day(0, Another_Day_Num) >= 0) Then
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
  '変更後の勤務形式が日勤だった場合,excess_miss_indexに0を設定し, 余り不足数の評価を加えないことにする
  If excess_miss_index = 0 + 1 Then
    change_need_value = request_aroud_results_month(month_sum_index, person_num)
  Else
    change_need_value = request_aroud_results_month(month_sum_index, person_num) - request_aroud_results_month(excess_miss_index, person_num) * 100
  End If
End Function
Function day_work_classifying(ByVal Work_Name As String) As Variant '「日勤, 当直, 準夜勤, 深夜勤, 休みの方の中で」「月合計が余っている, かつ, 最大の方+他の条件」を探索する時にその方の要素数を取得する為に使う
  Dim day_work_classify_num As Integer
  Dim classify_people_num As Integer
  max_person_num = 0
  max_finding = 0
  fit_object_exist = "無し"
  Dim answer(1) As Variant
  Dim after_work_classify_num As Integer
  Dim need_value As Integer
  '日勤,休みは二交代制, 三交代制の両方の場合であるので, ここでnum分類
  If Work_Name = "日" Then
    day_work_classify_num = 0
  End If
  
  If work_name_after = "日" Then
    after_work_classify_num = 0
  End If
  
  If Work_Name = "休" Then
    day_work_classify_num = UBound(request_aroud_results_month, 1) - 1
  End If
  
  If work_name_after = "休" Then
    after_work_classify_num = UBound(request_aroud_results_month, 1) - 1
  End If
  
  If how_work = "二交代制" Then
    If Work_Name = "当" Then
      day_work_classify_num = 1
    End If
    If work_name_after = "当" Then
      after_work_classify_num = 1
    End If
  
    For classify_people_num = 0 To number_of_people - 1
      '固定勤務希望でなければ
      If hope_or_not(classify_people_num, j) = 0 Then
        need_value = change_need_value(day_work_classify_num, after_work_classify_num + 1, classify_people_num)
        If Work_Name = "日" Then
          '日勤の探索の場合, 余り, 不足日数の表示はない為, 合計値の比較のみを行う
          '日勤の方の中で月合計が最大の方の人数番号をmax_person_numに代入し, 月合計値をmax_findingに代入する
          '尚, 日勤を当直に変える時は3日前に当直がないことと後の3日間は当直がないこと,
          '後の二日間が明け→休みに変えていいこと(つまり, 後の二日間の日勤の1日の余りがいること)を確認する
          If request_table(classify_people_num, j) = "日" Then
            If max_finding = 0 Or need_value > max_finding Then
              If work_name_after = "当" Then '当直に変える場合
                If Conditional_Branching(classify_people_num, j - 3, "当", "not_equal", 0, "small") And _
                Conditional_Branching(classify_people_num, j + 1, "当", "not_equal", number_of_day - 1, "big") And _
                Conditional_Branching(classify_people_num, j + 2, "当", "not_equal", number_of_day - 1, "big") And _
                Conditional_Branching(classify_people_num, j + 3, "当", "not_equal", number_of_day - 1, "big") And _
                After_day_Conditional_Branching(classify_people_num, j + 1, number_of_day - 1) And _
                After_day_Conditional_Branching(classify_people_num, j + 2, number_of_day - 1) Then
                  Call after_find_function(classify_people_num, need_value)
                End If
              Else '日勤を休みに変える場合は「当直→明け→休み」の変更可能確認はせずに変える
                Call after_find_function(classify_people_num, need_value)
              End If
            End If
          End If
        ElseIf Work_Name = "当" Then
          If request_table(classify_people_num, j) = Work_Name And request_aroud_results_month(day_work_classify_num + 1, classify_people_num) > 0 And _
          (max_finding = 0 Or need_value > max_finding) Then
          '当直の探索の場合, 余り, 不足日数の表示もある為, 余っていることを確認した上で, 合計値の比較のみを行う
          '当直の方の中で月合計が最大の方の人数番号をmax_person_numに代入し, 月合計値をmax_findingに代入する
            Call after_find_function(classify_people_num, need_value)
          End If
        ElseIf Work_Name = "休" Then
          If request_table(classify_people_num, j) = Work_Name And request_aroud_results_month(day_work_classify_num + 1, classify_people_num) > 0 And _
            (max_finding = 0 Or need_value > max_finding) Then
          '休みの探索の場合, 余り, 不足日数の表示もある為, 余っていることを確認した上で, 合計値の比較のみを行う
          '休みの方の中で月合計が最大の方の人数番号をmax_person_numに代入し, 月合計値をmax_findingに代入する
          
          '休みを日勤に変える時は
          '前2日間に当直が無いことを確認する
            If work_name_after = "日" Then
              If Conditional_Branching(classify_people_num, j - 2, "当", "not_equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "当", "not_equal", 0, "small") Then
                Call after_find_function(classify_people_num, need_value)
              End If
          '休みを当直に変える時は
          '2,3日前に当直がない, かつ, 後の3日間に当直がないことと
          '後の二日間が明け→休みに変えていいこと(つまり, 後の二日間の日勤の1日の余りがいること)を確認する
            ElseIf work_name_after = "当" Then
              If Conditional_Branching(classify_people_num, j - 2, "当", "not_equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 3, "当", "not_equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j + 1, "当", "not_equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "当", "not_equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 3, "当", "not_equal", number_of_day - 1, "big") And _
              After_day_Conditional_Branching(classify_people_num, j + 1, number_of_day - 1) And _
              After_day_Conditional_Branching(classify_people_num, j + 2, number_of_day - 1) Then
                Call after_find_function(classify_people_num, need_value)
              End If
            End If
          End If
        End If
      End If
    Next
  
  ElseIf how_work = "三交代制" Then
    If Work_Name = "準" Then
      day_work_classify_num = 1
    ElseIf Work_Name = "深" Then
      day_work_classify_num = 3
    End If
    If work_name_after = "準" Then
      after_work_classify_num = 1
    ElseIf work_name_after = "深" Then
      after_work_classify_num = 3
    End If
    
    For classify_people_num = 0 To number_of_people - 1
      '固定勤務希望でなければ
      If hope_or_not(classify_people_num, j) = 0 Then
        need_value = change_need_value(day_work_classify_num, after_work_classify_num + 1, classify_people_num)
        '日勤の探索の場合, 余り, 不足日数の表示はない為, 合計値の比較のみを行う
        '日勤の方の中で月合計が最大の方の人数番号をmax_person_numに代入し, 月合計値をmax_findingに代入する
        If Work_Name = "日" Then
          If request_table(classify_people_num, j) = "日" And _
          (max_finding = 0 Or need_value > max_finding) Then
          '尚, 日勤を準夜勤に変える時は, 前日が休みか日勤 かつ, (翌日が休み, もしくは, 翌日→2日後が「準→休」である)
          'ことを確認する
            If work_name_before = "日" And work_name_after = "準" Then
              If (Conditional_Branching(classify_people_num, j - 1, "休", "equal", 0, "small") Or _
              Conditional_Branching(classify_people_num, j - 1, "日", "equal", 0, "small")) And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") Or _
              (Conditional_Branching(classify_people_num, j + 1, "準", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "休", "equal", number_of_day - 1, "big")) _
              ) Then
                Call after_find_function(classify_people_num, need_value)
              End If
            '尚, 日勤を深夜勤に変える時は, 前日が休み かつ, (翌日が休み, もしくは, 翌日→2日後が「準→休」である)
            'ことを確認する
            ElseIf work_name_before = "日" And work_name_after = "深" Then
              If Conditional_Branching(classify_people_num, j - 1, "休", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") Or _
              (Conditional_Branching(classify_people_num, j + 1, "準", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "休", "equal", number_of_day - 1, "big")) _
              ) Then
                Call after_find_function(classify_people_num, need_value)
              End If
            Else '日勤を休みに変える時は条件はそのまま
                Call after_find_function(classify_people_num, need_value)
            End If
          End If
        '準夜勤,深夜勤, 休み(日勤以外)の探索の場合, 余り, 不足日数の表示もある為, 余っていることを確認した上で, 合計値の比較を行う
        '準夜勤,深夜勤, 休み(日勤以外)の方の中で月合計が最大の方の人数番号をmax_person_numに代入し, 月合計値をmax_findingに代入する
        ElseIf Work_Name = "準" Then
          If request_table(classify_people_num, j) = "準" And request_aroud_results_month(day_work_classify_num + 1, classify_people_num) > 0 And _
          (max_finding = 0 Or need_value > max_finding) Then
          '尚, 準夜勤を日勤に変える時は, 前日が準夜勤, 又は, 深夜勤務ではないことを確認する
            If work_name_before = "準" And work_name_after = "日" Then
              If Conditional_Branching(classify_people_num, j - 1, "準", "not_equal", 0, "small") And _
                Conditional_Branching(classify_people_num, j - 1, "深", "not_equal", 0, "small") Then
                  Call after_find_function(classify_people_num, need_value)
              End If
            '尚, 準夜勤を深夜勤に変える時は, (前日が休みであり,かつ, (翌日が休み, 又は, 翌日→2日後が準→休みである))
            '又は, ((2日前→1日前が休→深), かつ, 1日後が休である)ことを確認する
            ElseIf work_name_before = "準" And work_name_after = "深" Then
              If _
              ( _
              Conditional_Branching(classify_people_num, j - 2, "休", "equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "深", "equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") _
              ) Or _
              ( _
              Conditional_Branching(classify_people_num, j - 1, "休", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") Or _
              (Conditional_Branching(classify_people_num, j + 1, "準", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "休", "equal", number_of_day - 1, "big")) _
              ) _
              ) Then
                Call after_find_function(classify_people_num, need_value)
              End If
            Else '準夜勤を休みに変える時は条件はそのまま
              Call after_find_function(classify_people_num, need_value)
            End If
          End If
        
        ElseIf Work_Name = "深" Then
          If request_table(classify_people_num, j) = "深" And request_aroud_results_month(day_work_classify_num + 1, classify_people_num) > 0 And _
          (max_finding = 0 Or need_value > max_finding) Then
          '尚, 深夜勤を日勤に変える時は, 前日が休み, かつ, (翌日が休みか翌日→2日後, 準夜勤→休みである)ことを確認する
            If work_name_before = "深" And work_name_after = "日" Then
              If Conditional_Branching(classify_people_num, j - 1, "休", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") Or _
              (Conditional_Branching(classify_people_num, j + 1, "準", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "休", "equal", number_of_day - 1, "big")) _
              ) Then
                Call after_find_function(classify_people_num, need_value)
              End If
            '尚, 深夜勤を準夜勤に変える時は, (前日が休み, かつ, (翌日が休みか翌日→2日後, 準夜勤→休みである))か
            '(2日前→前日が休→深, かつ, 翌日が休み)であることを確認する
            ElseIf work_name_before = "深" And work_name_after = "準" Then
              If ( _
              Conditional_Branching(classify_people_num, j - 1, "休", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") Or _
              (Conditional_Branching(classify_people_num, j + 1, "準", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "休", "equal", number_of_day - 1, "big")) _
              ) _
              ) _
              Or _
              ( _
              (Conditional_Branching(classify_people_num, j - 2, "休", "equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "深", "equal", 0, "small")) And _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") _
              ) _
              Then
                Call after_find_function(classify_people_num, need_value)
              End If
            Else '深夜勤を休みに変える時は条件はそのまま
                Call after_find_function(classify_people_num, need_value)
            End If
          End If
          
        ElseIf Work_Name = "休" Then
          If request_table(classify_people_num, j) = "休" And request_aroud_results_month(day_work_classify_num + 1, classify_people_num) > 0 And _
          (max_finding = 0 Or need_value > max_finding) Then
          '尚, 休みを日勤に変える時は, (前日が準, 深ではなく,) かつ, (翌日が深ではない)ことを確認する
            If work_name_before = "休" And work_name_after = "日" Then
              If ( _
              Conditional_Branching(classify_people_num, j - 1, "準", "not_equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "深", "not_equal", 0, "small") _
              ) And _
              Conditional_Branching(classify_people_num, j + 1, "深", "not_equal", number_of_day - 1, "big") Then
                Call after_find_function(classify_people_num, need_value)
              End If
            '尚, 休みを準夜勤に変える時は, (前日が休み, かつ, (翌日が休みか翌日→2日後, 準夜勤→休みである))か
            '(2日前→前日が休→深, かつ, 翌日が休み)であることを確認する
            ElseIf work_name_before = "休" And work_name_after = "準" Then
              If ( _
              Conditional_Branching(classify_people_num, j - 1, "休", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") Or _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "準", "equal", number_of_day - 1, "big") And _
              Conditional_Branching(classify_people_num, j + 2, "休", "equal", number_of_day - 1, "big") _
              ) _
              ) _
              ) _
              Or _
              ( _
              (Conditional_Branching(classify_people_num, j - 2, "休", "equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "深", "equal", 0, "small")) And _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") _
              ) _
              Then
                Call after_find_function(classify_people_num, need_value)
              End If
           '尚, 休みを深夜勤に変える時は,
           '(前日が休み, かつ, (翌日が休みか翌日→2日後, (準夜勤か深夜)→休みである))か
           '(2日前→前日が休→深, かつ, 翌日が休み)であることを確認する
            ElseIf work_name_before = "休" And work_name_after = "深" Then
              If ( _
              Conditional_Branching(classify_people_num, j - 1, "休", "equal", 0, "small") And _
              ( _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") Or _
              ( _
              (Conditional_Branching(classify_people_num, j + 1, "準", "equal", number_of_day - 1, "big") Or Conditional_Branching(classify_people_num, j + 1, "深", "equal", number_of_day - 1, "big")) And _
              Conditional_Branching(classify_people_num, j + 2, "休", "equal", number_of_day - 1, "big") _
              ) _
              ) _
              ) _
              Or _
              ( _
              (Conditional_Branching(classify_people_num, j - 2, "休", "equal", 0, "small") And _
              Conditional_Branching(classify_people_num, j - 1, "深", "equal", 0, "small")) And _
              Conditional_Branching(classify_people_num, j + 1, "休", "equal", number_of_day - 1, "big") _
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
  
  If day_work_classifying(work_name_before)(1) = "有り" Then
    keep_num_for_change = day_work_classifying(work_name_before)(0)
    request_table(keep_num_for_change, j) = work_name_after
  Else
    GoTo Point_Change_End
  End If
  
  '日勤か休みを扱う場合, 二交代制, 三交代制のどちらも日勤と休みはあるので,
  '最初にchange_before_num, change_after_numを割り当てる
  If work_name_before = "日" Then
      change_before_num = 0
  End If
  
  If work_name_after = "日" Then
      change_after_num = 0
  End If
  
  If work_name_before = "休" Then
      change_before_num = UBound(request_aroud_results_month, 1) - 1
  End If
  
  If work_name_after = "休" Then
      change_after_num = UBound(request_aroud_results_month, 1) - 1
  End If
  
  'chage_person_numを算出する
  
  change_person_num = keep_num_for_change
  
  If how_work = "二交代制" Then
  '二交代制における変更前の勤務時間帯が変更されることによる1日の余り・不足人数表示と月合計と月合計の余り・不足人数表示の変更
    If work_name_before = "当" Then
       change_before_num = 1
    End If
    
    If work_name_before <> "休" Then '変更前が日勤,当直(休み以外)の場合, 1日の余り・不足人数表示を-1する
      request_aroud_results_day(change_before_num, j) = request_aroud_results_day(change_before_num, j) - 1
    End If
    
    request_aroud_results_month(change_before_num, change_person_num) = _
    request_aroud_results_month(change_before_num, change_person_num) - 1
    
    If work_name_before <> "日" Then '変更前が当直,休み(日勤以外)の探索の場合, 月合計の余り, 不足日数の表示を-1する
      request_aroud_results_month(change_before_num + 1, change_person_num) = _
      request_aroud_results_month(change_before_num + 1, change_person_num) - 1
    End If
      
    '変更前が当直で, 月末日よりjが小さい時, 変更前の当直を日勤か休みに変えると同時に, 翌日の明けを日勤か休みに変える
    If work_name_before = "当" And j < number_of_day - 1 Then
      request_table(change_person_num, j + 1) = work_name_after
      If work_name_after = "日" Then
        request_aroud_results_day(0, j + 1) = request_aroud_results_day(0, j + 1) + 1 '1日の日勤の余り・不足人数表示を+1する
        request_aroud_results_month(0, change_person_num) = _
        request_aroud_results_month(0, change_person_num) + 1 '日勤の月合計を+1する
      ElseIf work_name_after = "休" Then
        request_aroud_results_month(3, change_person_num) = _
        request_aroud_results_month(3, change_person_num) + 1 '休みの月合計を+1する
        request_aroud_results_month(4, change_person_num) = _
        request_aroud_results_month(4, change_person_num) + 1 '休みの月合計の余り・不足人数表示を+1する
      End If
    End If
    
  '二交代制における変更後の勤務時間帯が変更されることによる1日の余り・不足人数表示と月合計と月合計の余り・不足人数表示の変更
    If work_name_after = "当" Then
      change_after_num = 1
    End If
    
    If work_name_after <> "休" Then '変更後が日勤,当直(休み以外)の場合, 1日の余り・不足人数表示を+1する
      request_aroud_results_day(change_after_num, j) = request_aroud_results_day(change_after_num, j) + 1
    End If
    
    '変更後の勤務時間帯の月合計を+1する
    request_aroud_results_month(change_after_num, change_person_num) = _
    request_aroud_results_month(change_after_num, change_person_num) + 1
    
    If work_name_after <> "日" Then '変更後が当直,休み(日勤以外)の探索の場合, 月合計の余り, 不足日数の表示を+1する
      request_aroud_results_month(change_after_num + 1, change_person_num) = _
      request_aroud_results_month(change_after_num + 1, change_person_num) + 1
    End If
      
    '変更後が当直で, 月末日よりjより小さい,つまり,月末まで最低でも1日は空いている時,
    '日勤か休みを当直(変更後)に変えると同時に, 翌日の日勤か休みを「明」け に変える
    If work_name_after = "当" And j < number_of_day - 1 Then
      If request_table(change_person_num, j + 1) = "日" Then
        request_aroud_results_day(0, j + 1) = request_aroud_results_day(0, j + 1) - 1 '1日の日勤の余り・不足人数表示を-1する
        request_aroud_results_month(0, change_person_num) = _
        request_aroud_results_month(0, change_person_num) - 1 '日勤の月合計を-1する
      ElseIf request_table(change_person_num, j + 1) = "休" Then
        request_aroud_results_month(3, change_person_num) = _
        request_aroud_results_month(3, change_person_num) - 1 '休みの月合計を-1する
        request_aroud_results_month(4, change_person_num) = _
        request_aroud_results_month(4, change_person_num) - 1 '休みの月合計の余り・不足人数表示を+1する
      End If
      request_table(change_person_num, j + 1) = "明"
      '月末日よりj-1より小さい,つまり,月末まで最低でも2日は空いている時,
      '日勤か休みを当直(変更後)に変えると同時に, 2日後の日勤を「休」み に変える
      If j < (number_of_day - 1) - 1 Then
        If request_table(change_person_num, j + 2) = "日" Then
          request_aroud_results_day(0, j + 2) = request_aroud_results_day(0, j + 2) - 1 '1日の日勤の余り・不足人数表示を-1する
          request_aroud_results_month(0, change_person_num) = _
          request_aroud_results_month(0, change_person_num) - 1 '日勤の月合計を-1する
        End If
        request_table(change_person_num, j + 2) = "休"
      End If
    End If
  
  ElseIf how_work = "三交代制" Then
    If work_name_before = "準" Then
      change_before_num = 1
    ElseIf work_name_before = "深" Then
      change_before_num = 3
    End If
    
    '変更前が日勤, 準夜勤, 深夜勤の場合(1日の余り・不足表示を-1する)
    If work_name_before = "日" Or work_name_before = "準" Then
      request_aroud_results_day(change_before_num, j) = request_aroud_results_day(change_before_num, j) - 1
    ElseIf work_name_before = "深" Then
      request_aroud_results_day(2, j) = request_aroud_results_day(2, j) - 1
    End If
    
    request_aroud_results_month(change_before_num, change_person_num) = _
    request_aroud_results_month(change_before_num, change_person_num) - 1
    
    If work_name_before <> "日" Then '変更前が準夜勤, 深夜勤,休み(日勤以外)の探索の場合, 月合計の余り, 不足日数の表示もある
      request_aroud_results_month(change_before_num + 1, change_person_num) = _
      request_aroud_results_month(change_before_num + 1, change_person_num) - 1
    End If

    If work_name_after = "準" Then
      change_after_num = 1
    ElseIf work_name_after = "深" Then
      change_after_num = 3
    End If
    
    '変更後が日勤, 準夜勤, 深夜勤の場合(1日の余り・不足表示を+1する)
    If work_name_after = "日" Or work_name_after = "準" Then
      request_aroud_results_day(change_after_num, j) = request_aroud_results_day(change_after_num, j) + 1
    ElseIf work_name_after = "深" Then
      request_aroud_results_day(2, j) = request_aroud_results_day(2, j) + 1
    End If
    
    request_aroud_results_month(change_after_num, change_person_num) = _
    request_aroud_results_month(change_after_num, change_person_num) + 1
    
    If work_name_after <> "日" Then '変更後が準夜勤, 深夜勤,休み(日勤以外)の探索の場合, 月合計の余り, 不足日数の表示もある
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
  If how_work = "二交代制" Then
    If insufficient_work_name = "日" Then
     insufficient_num = 0
     compensation_work(0) = "当"
    ElseIf insufficient_work_name = "当" Then
      insufficient_num = 1
      compensation_work(0) = "日"
    End If
    Call work_name_decide(compensation_work(0), insufficient_work_name)
    If request_aroud_results_day(inversion(insufficient_num), j) > 0 Then 'その日の補填勤務時間帯人数が余っていたら,
      While request_aroud_results_day(insufficient_num, j) < 0 And _
      day_work_classifying(compensation_work(0))(1) = "有り" And _
      request_aroud_results_day(inversion(insufficient_num), j) > 0
      '1日のその人数不足が0になるか補填勤務時間帯人数の月合計で余っている方がいなくなるかその日の補填勤務時間帯人数が余らなくなる時まで
       Call change_write
       '補填勤務時間帯の方の中で月合計が余っていて, かつ, 月合計が最大の補填勤務時間帯の方を不足勤務に変える
      Wend
      If request_aroud_results_day(insufficient_num, j) = 0 Then 'その日の不足勤務が不足していくなっていたら, パス
      ElseIf day_work_classifying(compensation_work(0))(1) = "無し" Or _
      request_aroud_results_day(inversion(insufficient_num), j) <= 0 Then
      'その日の不足勤務がまだ不足していて, 補填勤務時間帯の月合計が余っている人かその日の補填勤務時間帯の余り人がいなかったら,
        Call work_name_decide("休", insufficient_work_name)
        While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("休")(1) = "有り"
        '1日の不足勤務の不足人数が0になるか休みの月合計で余っている方がいなくなる時まで
          Call change_write
        Wend
      End If
    Else 'その日の補填勤務時間帯人数が余っていなかったら,
      Call work_name_decide("休", insufficient_work_name)
      While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("休")(1) = "有り"
        '1日の不足勤務の不足人数が0になるか休みの月合計で余っている方がいなくなる時まで
        Call change_write
      Wend
    End If
  ElseIf how_work = "三交代制" Then
    Dim compensation_num(1) As Integer
    Dim three_change_loop_num As Integer
    If insufficient_work_name = "日" Then
      insufficient_num = 0
      compensation_num(0) = 1
      compensation_num(1) = 2
      compensation_work(0) = "準"
      compensation_work(1) = "深"
    ElseIf insufficient_work_name = "準" Then
      insufficient_num = 1
      compensation_num(0) = 0
      compensation_num(1) = 2
      compensation_work(0) = "日"
      compensation_work(1) = "深"
    ElseIf insufficient_work_name = "深" Then
      insufficient_num = 2
      compensation_num(0) = 0
      compensation_num(1) = 1
      compensation_work(0) = "日"
      compensation_work(1) = "準"
    End If
    
    For three_change_loop_num = 0 To 1
      If request_aroud_results_day(compensation_num(three_change_loop_num), j) > 0 And _
      request_aroud_results_day(compensation_num(inversion(three_change_loop_num)), j) <= 0 Then '最初(二つ目)の方の補填勤務時間帯人数のみが余っていたら,
        Call work_name_decide(compensation_work(three_change_loop_num), insufficient_work_name)
        While request_aroud_results_day(insufficient_num, j) < 0 And _
        day_work_classifying(compensation_work(three_change_loop_num))(1) = "有り" And _
        request_aroud_results_day(compensation_num(three_change_loop_num), j) > 0
        '1日のその人数不足が0になるか最初(二つ目)の方の補填勤務時間帯人数の月合計で余っている方がいなくなるか
        'その日の最初(二つ目)の方の補填勤務時間帯人数が余らなくなる時まで
         Call change_write
         '補填勤務時間帯の方の中で月合計が余っていて, かつ, 月合計が最大の補填勤務時間帯の方を不足勤務に変える
        Wend
        If request_aroud_results_day(insufficient_num, j) = 0 Then 'その日の不足勤務が不足していくなっていたら, パス
        ElseIf day_work_classifying(compensation_work(three_change_loop_num))(1) = "無し" Or request_aroud_results_day(compensation_num(three_change_loop_num), j) <= 0 Then
        'その日の不足勤務がまだ不足していて, 補填勤務時間帯の月合計が余っている人かその日の補填勤務時間帯の余り人がいなかったら,
          Call work_name_decide("休", insufficient_work_name)
          While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("休")(1) = "有り"
          '1日の不足勤務の不足人数が0になるか休みの月合計で余っている方がいなくなる時まで
            Call change_write
          Wend
        End If
      End If
    Next
    
    If request_aroud_results_day(compensation_num(0), j) > 0 And _
    request_aroud_results_day(compensation_num(1), j) > 0 Then 'その日の補填勤務時間帯人数がどちらも余っていたら,
      If insufficient_work_name = "日" Then '補填勤務時間帯が準夜勤と深夜勤務である場合, (不足勤務時間帯が日勤の場合)
        If request_aroud_results_day(compensation_num(0), j) > request_aroud_results_day(compensation_num(1), j) Then '1日あたりで準夜勤が深夜勤より余っていたら,
          While request_aroud_results_day(insufficient_num, j) < 0 And _
          day_work_classifying("準")(1) = "有り" And _
          request_aroud_results_day(compensation_num(0), j) > request_aroud_results_day(compensation_num(1), j) '日勤が不足し, 月合計の観点から変更可能な準夜勤がいて, 1日あたり準夜勤が深夜勤より余っていたら,
            Call work_name_decide("準", "日")
            Call change_write  'まず, 準夜勤を補填に充てる
          Wend
          If request_aroud_results_day(insufficient_num, j) = 0 Then
            GoTo Point_End
          ElseIf day_work_classifying("準")(1) = "無し" Then
            GoTo Midnight_Start
          ElseIf request_aroud_results_day(compensation_num(0), j) = request_aroud_results_day(compensation_num(1), j) Then
           GoTo Semi_Night_Start
          End If
        ElseIf request_aroud_results_day(compensation_num(0), j) < request_aroud_results_day(compensation_num(1), j) Then '1日あたりで深夜勤が準夜勤より余っていたら,
          While request_aroud_results_day(insufficient_num, j) < 0 And _
          day_work_classifying("深")(1) = "有り" And _
          request_aroud_results_day(compensation_num(1), j) > request_aroud_results_day(compensation_num(0), j) '日勤が不足し, 月合計の観点から変更可能な深夜勤がいて, 1日あたり深夜勤が準夜勤より余っていたら,
            Call work_name_decide("深", "日")
            Call change_write  '深夜勤を補填に充てる
          Wend
          If request_aroud_results_day(insufficient_num, j) = 0 Then
            GoTo Point_End
          Else
            GoTo Semi_Night_Start
          End If
        ElseIf request_aroud_results_day(compensation_num(0), j) = request_aroud_results_day(compensation_num(1), j) Then '1日あたりの余り人数で深夜勤=準夜勤なら
Semi_Night_Start:
          Call work_name_decide("準", "日")
          Call change_write  'まず, 準夜勤を補填に充てる
          ElseIf request_aroud_results_day(compensation_num(1), j) > request_aroud_results_day(compensation_num(0), j) Then '1日あたり深夜勤が準夜勤より余っていたら,
            GoTo Midnight_Start
          ElseIf request_aroud_results_day(insufficient_num, j) = 0 Then
            GoTo Point_End
          ElseIf day_work_classifying("深")(1) = "無し" Then
            If day_work_classifying("準")(1) = "有り" Then
              GoTo Semi_Night_Start
            Else
Holiday_start:
              Call work_name_decide("休", "日")
              While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("休")(1) = "有り"
                Call change_write
              Wend
              GoTo Point_End
            End If
          Else
Midnight_Start:
            Call work_name_decide("深", "日")
            Call change_write  'まず, 深夜勤を補填に充てる
            If request_aroud_results_day(insufficient_num, j) = 0 Then
              GoTo Point_End
            ElseIf request_aroud_results_day(1, j) = 0 And request_aroud_results_day(2, j) = 0 Then '準夜勤と深夜勤の両方とも１日の勤務人数に余りがなくなり, 変更不可になったら
              GoTo Holiday_start
            ElseIf day_work_classifying("準")(1) = "無し" Then
              If day_work_classifying("深")(1) = "有り" Then
                GoTo Midnight_Start
              Else
                GoTo Holiday_start
              End If
            End If
          End If
        End If
Point_End:
      Else '補填勤務時間帯が日勤と準(深)夜勤である場合
        Call work_name_decide(compensation_work(1), insufficient_work_name) 'まず, 準(深)夜勤を補填に充てる
        While request_aroud_results_day(insufficient_num, j) < 0 And _
        day_work_classifying(compensation_work(1))(1) = "有り" And _
        request_aroud_results_day(compensation_num(1), j) > 0
        '1日のその人数不足が0になるか準(深)夜勤の補填勤務時間帯人数の月合計で余っている方がいなくなるか
        'その日の準(深)夜勤の補填勤務時間帯人数が余らなくなる時まで
          Call change_write
         '準(深)夜勤の方の中で月合計が余っていて, かつ, 月合計が最大の準(深)夜勤の方を不足勤務に変える
        Wend
        If request_aroud_results_day(insufficient_num, j) = 0 Then
        ElseIf day_work_classifying(compensation_work(1))(1) = "無し" Or request_aroud_results_day(compensation_num(1), j) = 0 Then
          Call work_name_decide(compensation_work(0), insufficient_work_name) '次に, 日勤を補填に充てる
          While request_aroud_results_day(insufficient_num, j) < 0 And _
          day_work_classifying(compensation_work(0))(1) = "有り" And _
          request_aroud_results_day(compensation_num(0), j) > 0
          '1日のその人数不足が0になるか日勤の補填勤務時間帯人数の月合計で余っている方がいなくなるか
          'その日の日勤の方の補填勤務時間帯人数が余らなくなる時まで
            Call change_write
           '日勤の方の中で月合計が余っていて, かつ, 月合計が最大の日勤の方を不足勤務に変える
          Wend
          If request_aroud_results_day(insufficient_num, j) = 0 Then
          ElseIf day_work_classifying(compensation_work(0))(1) = "有り" And request_aroud_results_day(compensation_num(0), j) > 0 Then
            Call work_name_decide("休", insufficient_work_name)
            While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("休")(1) = "有り"
            '1日の不足勤務の不足人数が0になるか休みの月合計で余っている方がいなくなる時まで
              Call change_write
            Wend
          End If
        End If
      End If
    
    If request_aroud_results_day(compensation_num(0), j) <= 0 And _
    request_aroud_results_day(compensation_num(1), j) <= 0 Then 'その日の補填勤務時間帯人数がどちらも余っていなかったら,
      Call work_name_decide("休", insufficient_work_name)
      While request_aroud_results_day(insufficient_num, j) < 0 And day_work_classifying("休")(1) = "有り"
        '1日の不足勤務の不足人数が0になるか休みの月合計で余っている方がいなくなる時まで
          Call change_write
      Wend
    End If
  End If
End Function
Sub Bチーム用勤務表自動作成実行()
'Private変数の再初期化
  'グローバル変数の定義
  '読み込み用シート指定
  Set ws = ThisWorkbook.Sheets("Bチーム用勤務希望表")
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
  'ループ変数の定義
  j = 0 'number_of_dayのループ用のj
  
'読み込み時用の1日1人勤務形式格納変数
  Dim work_content As String
'Naming
  Dim i As Integer 'number_of_peopleのループ用のi
  '以下, デバッグ用変数
  'Dim dbug_j As Integer
  'Dim dbug_titles As Integer
'精度向上用の評価値変数
  Dim loss As Long
  Dim weight As Integer
  Dim save_loss As Long
  Dim save_request_for_evaluate() As String
  Dim distinction_first_0 As Boolean 'Trueになったら, 一度は評価したことを示す
  
'備考読み込み用変数
  Dim remarks_table() As String
  
  On Error GoTo ErrLabel
  
'exectuing
  
  'name,number_of_people_reading
  For i = 19 To 19 + 2 * (30 - 1) Step 2 'For i=19 To 19+2*(30(人)-1) Step 2
    If ws.Cells(i, 5) <> "" Then
      Names.Add ws.Cells(i, 5)
      number_of_people = number_of_people + 1
    End If
  Next
  '名前が無かったら, 終了する
  If number_of_people = 0 Then
    MsgBox "勤務者の氏名を入力して下さい！", _
    Buttons:=vbCritical, Title:="実行者の方向けのエラー表示"
    Exit Sub
  End If
  
  'number_of_day_reading
  For j = 11 To 41 'For j = 11(K) To 41(AO)
    If ws.Cells(18, j) <> "×" Then
      number_of_day = number_of_day + 1
    End If
  Next
  
  '開始前の設定不足, エラー, 日数不足確認
  If Before_building_check = "設定不足, エラー, 日数不足有り" Then
    Exit Sub
  End If
  
  'VBA高速化開始
  Call high_speeding(True)
  
  '勤務制度把握
  how_work = ws.Cells(1, 4)
  
  'request_aroud_results
  If how_work = "二交代制" Then
    ReDim request_aroud_results_day(2 - 1, number_of_day - 1)
    ReDim request_aroud_results_month(5 - 1, number_of_people - 1)
    'ReDim request_aroud_results_day((評価指標の数) - 1, number_of_day(people) - 1)
    '二交代制の1日毎の評価指標
    'D1.1日の日勤の余り, 不足人数表示
    'D2.1日の当直の余り, 不足人数表示
    '二交代制の1か月毎の評価指標
    'M1.1カ月の日勤の合計
    'M2.1カ月の当直の合計
    'M3.1カ月の当直の余り, 不足日数表示
    'M4.1カ月の休みの合計
    'M5.1カ月の休みの余り, 不足日数表示
    For j = 11 To 11 + number_of_day - 1
      'D2.1日の当直の余り, 不足人数表示
      request_aroud_results_day(1, j - 11) = ws.Cells(12, j)
    Next
    For i = 19 To 19 + 2 * (number_of_people - 1) Step 2
    'M2.1カ月の当直の合計
    'M3.1カ月の当直の余り, 不足日数表示
      request_aroud_results_month(1, (i - 19) / 2) = ws.Cells(i, 43)
      request_aroud_results_month(2, (i - 19) / 2) = excess_miss(ws.Cells(i, 44))
    Next
  ElseIf how_work = "三交代制" Then
    ReDim request_aroud_results_day(3 - 1, number_of_day - 1)
    ReDim request_aroud_results_month(7 - 1, number_of_people - 1)
    'ReDim request_aroud_results_day((評価指標の数) - 1, number_of_day(people) - 1)
    '三交代制の1日毎の評価指標
    'D1.1日の日勤の余り, 不足人数表示
    'D2.1日の準夜勤の余り, 不足人数表示
    'D3.1日の深夜勤の余り, 不足人数表示
    '三交代制の1か月毎の評価指標
    'M1.1カ月の日勤の合計
    'M2.1カ月の準夜勤の合計
    'M3.1カ月の準夜勤の余り, 不足日数表示
    'M4.1カ月の深夜勤の合計
    'M5.1カ月の深夜勤の余り, 不足日数表示
    'M6.1カ月の休みの合計
    'M7.1カ月の休みの余り, 不足日数表示
    For j = 11 To 11 + number_of_day - 1
      'D2.1日の準夜勤の余り, 不足人数表示
      request_aroud_results_day(1, j - 11) = ws.Cells(13, j)
      'D3.1日の深夜勤の余り, 不足人数表示
      request_aroud_results_day(2, j - 11) = ws.Cells(14, j)
    Next
    For i = 19 To 19 + 2 * (number_of_people - 1) Step 2
    'M2.1カ月の準夜勤の合計
    'M3.1カ月の準夜勤の余り, 不足日数表示
    'M4.1カ月の深夜勤の合計
    'M5.1カ月の深夜勤の余り, 不足日数表示
      request_aroud_results_month(1, (i - 19) / 2) = ws.Cells(i, 45)
      request_aroud_results_month(2, (i - 19) / 2) = excess_miss(ws.Cells(i, 46))
      request_aroud_results_month(3, (i - 19) / 2) = ws.Cells(i, 47)
      request_aroud_results_month(4, (i - 19) / 2) = excess_miss(ws.Cells(i, 48))
    Next
  End If
  '～以下, 共通部分の代入開始～
  'D1.1日の日勤の余り, 不足人数表示
  For j = 11 To 11 + number_of_day - 1
    request_aroud_results_day(0, j - 11) = ws.Cells(11, j)
  Next
  'M1.1カ月の日勤の合計
  'M4(6).1カ月の休みの合計
  'M5(7).1カ月の休みの余り, 不足日数表示
  For i = 19 To 19 + 2 * (number_of_people - 1) Step 2
    request_aroud_results_month(0, (i - 19) / 2) = ws.Cells(i, 42)
    request_aroud_results_month(UBound(request_aroud_results_month, 1) - 1, (i - 19) / 2) = ws.Cells(i, 49)
    request_aroud_results_month(UBound(request_aroud_results_month, 1), (i - 19) / 2) = excess_miss(ws.Cells(i, 50))
  Next
  '～以上, 共通部分の代入終了～
  
  'request_table_readingとhope_or_notの格納
  ReDim request_table(number_of_people - 1, number_of_day - 1)
  ReDim hope_or_not(number_of_people - 1, number_of_day - 1)
  ReDim remarks_table(number_of_people - 1, number_of_day - 1)
  For i = 19 To 19 + 2 * (number_of_people - 1) Step 2
    For j = 11 To 11 + number_of_day - 1
      work_content = ws.Cells(i, j)
      If how_work = "二交代制" Then
        If work_content <> "" And work_content <> "日" And work_content <> "休" And work_content <> "当" And work_content <> "明" Then
          'VBA高速化を終了し
          Call high_speeding(False)
          'エラーメッセージを出力し
          MsgBox Names((i - 19) / 2 + 1) & "さんの" & CStr(j - 11 + 1) & "日の勤務形式が不適切です！" & vbCrLf & "「日」「休」「当」「明」のどれかを記載して下さい！", _
          Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
          'Subプロジージャを終了する
          Exit Sub
        End If
      ElseIf how_work = "三交代制" Then
        If work_content <> "" And work_content <> "日" And work_content <> "休" And work_content <> "準" And work_content <> "深" Then
          'VBA高速化を終了し
          Call high_speeding(False)
          'エラーメッセージを出力し
          MsgBox Names((i - 19) / 2 + 1) & "さんの" & CStr(j - 11 + 1) & "日の勤務形式が不適切です！" & vbCrLf & "「日」「休」「準」「深」のどれかを記載して下さい！", _
          Buttons:=vbCritical, Title:="設定者の方向けのエラー表示"
          'Subプロジージャを終了する
          Exit Sub
        End If
      End If
      '希望が書かれていたら,
      If work_content <> "" Then
        request_table((i - 19) / 2, j - 11) = work_content
        hope_or_not((i - 19) / 2, j - 11) = 1
      Else
        request_table((i - 19) / 2, j - 11) = "休"
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
  If how_work = "二交代制" Then
    For j = 0 To number_of_day - 1
      If request_aroud_results_day(0, j) < 0 Then '日勤人数が不足していたら,
        Call insufficient_compensation("日")
      End If
    Next
   
    For j = 0 To number_of_day - 1
      Call work_name_decide("当", "日")
      If request_aroud_results_day(1, j) > 0 Then
        While request_aroud_results_day(1, j) > 0 And day_work_classifying("当")(1) = "有り"
         '1日の当直が設定人数を超えていて, かつ, 当直から日勤に変更可能な方がいる場合
          Call change_write
        Wend
      End If
    Next
    
    For j = 0 To number_of_day - 1
      If request_aroud_results_day(1, j) < 0 Then '当直人数が不足していたら
        Call insufficient_compensation("当")
      End If
    Next
    
    For j = 0 To number_of_day - 1
      Call work_name_decide("当", "日")
      If request_aroud_results_day(1, j) > 0 Then
        While request_aroud_results_day(1, j) > 0 And day_work_classifying("当")(1) = "有り"
         '1日の当直が設定人数を超えていて, かつ, 当直から日勤に変更可能な方がいる場合
          Call change_write
        Wend
      End If
    Next
    
    For j = 0 To number_of_day - 1
      Call work_name_decide("日", "休")
      If request_aroud_results_day(0, j) > 0 Then
        While request_aroud_results_day(0, j) > 0 And day_work_classifying("日")(1) = "有り"
         '1日の日勤が設定人数を超えていて, かつ, 日勤から休みに変更可能な方がいる場合
          Call change_write
        Wend
      End If
    Next
    
  ElseIf how_work = "三交代制" Then
    
    For j = 0 To number_of_day - 1
      If request_aroud_results_day(0, j) < 0 Then '1日の日勤人数が不足していたら,
        Call insufficient_compensation("日")
      End If
    Next
    
    For j = 0 To number_of_day - 1
    
      Call work_name_decide("準", "日")
      While request_aroud_results_day(1, j) > 0 And day_work_classifying("準")(1) = "有り"
      '1日の準夜勤が設定人数を超えていて, かつ, 準夜勤から日勤に変更可能な方がいる場合
        Call change_write
      Wend
      
      Call work_name_decide("準", "休")
      If request_aroud_results_day(1, j) = 0 Then
      ElseIf day_work_classifying("準")(1) = "有り" Then
        While request_aroud_results_day(1, j) > 0 And day_work_classifying("準")(1) = "有り"
        '1日の準夜勤が設定人数を超えていて, かつ, 準夜勤から休みに変更可能な方がいる場合
          Call change_write
        Wend
      End If
      
      Call work_name_decide("深", "日")
      While request_aroud_results_day(2, j) > 0 And day_work_classifying("深")(1) = "有り"
      '1日の深夜勤が設定人数を超えていて, かつ, 深夜勤から日勤に変更可能な方がいる場合
        Call change_write
      Wend
      
      Call work_name_decide("深", "休")
      If request_aroud_results_day(2, j) = 0 Then
      ElseIf day_work_classifying("深")(1) = "有り" Then
        While request_aroud_results_day(2, j) > 0 And day_work_classifying("深")(1) = "有り"
        '1日の準夜勤が設定人数を超えていて, かつ, 深夜勤から休みに変更可能な方がいる場合
          Call change_write
        Wend
      End If
      
    Next
    
    For j = 0 To number_of_day - 1
      
      If request_aroud_results_day(1, j) < 0 Then '1日の準夜勤人数が不足していたら,
        Call insufficient_compensation("準")
      End If
      If request_aroud_results_day(2, j) < 0 Then '1日の深夜勤人数が不足していたら,
        Call insufficient_compensation("深")
      End If
    Next
    
    For j = 0 To number_of_day - 1
    
      Call work_name_decide("準", "日")
      While request_aroud_results_day(1, j) > 0 And day_work_classifying("準")(1) = "有り"
      '1日の準夜勤が設定人数を超えていて, かつ, 準夜勤から日勤に変更可能な方がいる場合
        Call change_write
      Wend
      
      Call work_name_decide("準", "休")
      If request_aroud_results_day(1, j) = 0 Then
      ElseIf day_work_classifying("準")(1) = "有り" Then
        While request_aroud_results_day(1, j) > 0 And day_work_classifying("準")(1) = "有り"
        '1日の準夜勤が設定人数を超えていて, かつ, 準夜勤から休みに変更可能な方がいる場合
          Call change_write
        Wend
      End If
      
      Call work_name_decide("深", "日")
      While request_aroud_results_day(2, j) > 0 And day_work_classifying("深")(1) = "有り"
      '1日の深夜勤が設定人数を超えていて, かつ, 深夜勤から日勤に変更可能な方がいる場合
        Call change_write
      Wend
      
      Call work_name_decide("深", "休")
      If request_aroud_results_day(2, j) = 0 Then
      ElseIf day_work_classifying("深")(1) = "有り" Then
        While request_aroud_results_day(2, j) > 0 And day_work_classifying("深")(1) = "有り"
        '1日の準夜勤が設定人数を超えていて, かつ, 深夜勤から休みに変更可能な方がいる場合
          Call change_write
        Wend
      End If
    Next
    
    For j = 0 To number_of_day - 1
      If request_aroud_results_day(1, j) < 0 Then '1日の準夜勤人数が不足していたら,
        Call insufficient_compensation("準")
      End If
      If request_aroud_results_day(2, j) < 0 Then '1日の深夜勤人数が不足していたら,
        Call insufficient_compensation("深")
      End If
    Next
    
    For j = 0 To number_of_day - 1
      Call work_name_decide("日", "休")
      If request_aroud_results_day(0, j) > 0 Then
        While request_aroud_results_day(0, j) > 0 And day_work_classifying("日")(1) = "有り"
         '1日の日勤が設定人数を超えていて, かつ, 日勤から休みに変更可能な方がいる場合
          Call change_write
        Wend
      End If
    Next
    
  End If
  
  '勤務表評価
    '1日の余り・不足人数に対する評価
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
  '日勤の月合計はlossの評価には関係ないので, 1から始める
  For i = 1 To UBound(request_aroud_results_month, 1)
    For j = 0 To number_of_people - 1
      If request_aroud_results_month(i, j) = 0 Then
        weight = 0
      Else
        If i = 2 Or (how_work = "三交代制" And i = 4) Or i = UBound(request_aroud_results_month, 1) Then
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
  
  '書き込み用シート指定
  Set ws = ThisWorkbook.Sheets("チーム間調整前総合勤務表")
  'シート保護の解除
  Call pro_unpro(False, "チーム間調整前総合勤務表")
  '削除
  Range(ws.Cells(60 + 19, 11), ws.Cells(60 + 78, 41)).ClearContents
  '太字→普通の細さにする
  Range(ws.Cells(60 + 19, 11), ws.Cells(60 + 78, 41)).Font.Bold = False
  '斜体→直体にする
  Range(ws.Cells(60 + 19, 11), ws.Cells(60 + 78, 41)).Font.Italic = False
  '字の色を赤→黒に変える
  Range(ws.Cells(60 + 19, 11), ws.Cells(60 + 78, 41)).Font.color = RGB(0, 0, 0)
  '書き込み
  For i = 60 + 19 To 60 + 19 + 2 * (number_of_people - 1) Step 2
    For j = 11 To 11 + number_of_day - 1
      '固定希望日ではなく, VBA上で自動作成した勤務形式ならば
      If hope_or_not((i - 60 - 19) / 2, j - 11) <> 1 Then
        '太字に変えて
        ws.Cells(i, j).Font.Bold = True
        '斜体に変えて
        ws.Cells(i, j).Font.Italic = True
        '字の色を赤に変える
        ws.Cells(i, j).Font.color = RGB(255, 0, 0)
      End If
      ws.Cells(i, j) = request_table((i - 60 - 19) / 2, j - 11)
      ws.Cells(i + 1, j) = remarks_table((i - 60 - 19) / 2, j - 11)
    Next
  Next
  
  '勤務表自動作成最終日時更新
  ws.Cells(21, 1) = "Bチーム:" & Format(Now(), "YYYY年MM月DD日HH時MM分")
  'シート保護の再有効化
  Call pro_unpro(True, "チーム間調整前総合勤務表")
  'VBA高速化を終了し
  Call high_speeding(False)
  
  '完了メッセージを表示する
  MsgBox "Bチーム単体での勤務表の作成が完了しました！" _
      & vbCrLf & "シート「チーム間調整前総合勤務表」をご確認下さい！", _
      Buttons:=vbInformation, Title:="実行者の方へのメッセージ"
  Exit Sub
   
ErrLabel:
  'シートの保護の再有効化
  Call pro_unpro(True, "チーム間調整前総合勤務表")
  'VBA高速化を終了する
  Call high_speeding(False)
  MsgBox "「Bチーム用勤務表自動作成実行」VBA実行中にエラーが発生しました！" _
  & vbCrLf & "シート「Bチーム用勤務希望表」をご確認の上, 再度, 実行して下さい！" _
  & vbCrLf & "エラー発生アプリ: " & Err.Source _
  & vbCrLf & "エラー番号: " & Err.Number _
  & vbCrLf & "エラー内容: " & Err.Description, _
  Buttons:=vbCritical, Title:="実行者の方向けのエラー表示"
End Sub



