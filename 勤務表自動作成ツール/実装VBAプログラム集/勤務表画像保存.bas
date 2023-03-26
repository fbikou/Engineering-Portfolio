Attribute VB_Name = "Module6"
Option Explicit
Private offset As Integer
Private new_sheet_name As String
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
            exist_sht = True ' 存在する
            Exit Function
        End If
    Next
    ' 存在しない
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
  '数式→値に変換
  pst_rng.Value = copy_rng.Value
  If pst_rng_left_top_row = 1 And pst_rng_left_top_col = 3 Then
    '条件付き書式の削除
    pst_rng.FormatConditions.Delete
    '背景色の追加
    For i = 2 To 32
      color = copy_rng(2, i).DisplayFormat.Interior.color
      pst_rng.Cells(2, i).Interior.color = color
      pst_rng.Cells(3, i).Interior.color = color
    Next
  End If
  '勤務予定の条件付き書式の変更
  If pst_rng.FormatConditions.count > 0 Then
    Set fc = pst_rng.FormatConditions(1)
    If fc.Type = xlExpression Then
      If fc.Formula1 = "=MOD(ROW(),2)=0" Then
        pst_rng.FormatConditions.Delete
        Set fc = pst_rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=MOD(ROW(),2)=1")
        fc.Interior.color = RGB(255, 255, 255)
        Set fc = pst_rng.FormatConditions.Add(Type:=xlTextString, String:="日", TextOperator:=xlContains)
        fc.Interior.color = RGB(255, 192, 0)
        Set fc = pst_rng.FormatConditions.Add(Type:=xlTextString, String:="当", TextOperator:=xlContains)
        fc.Interior.color = RGB(255, 255, 0)
        Set fc = pst_rng.FormatConditions.Add(Type:=xlTextString, String:="明", TextOperator:=xlContains)
        fc.Interior.color = RGB(255, 255, 0)
        Set fc = pst_rng.FormatConditions.Add(Type:=xlTextString, String:="準", TextOperator:=xlContains)
        fc.Interior.color = RGB(146, 208, 80)
        Set fc = pst_rng.FormatConditions.Add(Type:=xlTextString, String:="深", TextOperator:=xlContains)
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
  Dim copyRetryCount As Integer: copyRetryCount = 10 'リトライ回数
  'エラー時はリトライ処理に飛ぶ(コピー時にエラーが発生する想定の為, CopyRetryという名前にしている)
  On Error GoTo CopyRetry
  
  '■セル範囲を画像データでコピーする。
  rng.CopyPicture
  
  '■指定したセル範囲と同じサイズのpicを新規作成し、保存する。
  Set pic = ws_pic.ChartObjects.Add(0, 0, rng.Width, rng.Height)
  pic.chart.Export pic_path
  FileSize = FileLen(pic_path)
  
  '■picのFileSizeを超えるまでループする(画像データが出来上がったら終了する)
  Do Until FileLen(pic_path) > FileSize
   
    pic.chart.Paste
    pic.chart.Export pic_path
    DoEvents
  Loop
  '■作成完了後、pic削除。
  pic.Delete
  Set pic = Nothing
  
  If event_stop_judge = True Then
    Application.EnableEvents = False
  End If
  
  Exit Function 'リトライ処理に飛ぶ前にメソッドから抜ける
  
'リトライ処理
CopyRetry:
    '一定時間待機後、エラー直前の処理に飛ぶ
    copyRetryCount = copyRetryCount - 1
    If copyRetryCount >= 1 Then
        '残りリトライ回数が1より大きい場合は、再度, 実行する
        '100ミリ秒 経過後再試行
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
  
  'フォルダ削除
  objFSO.DeleteFolder strFolderPath
  
  'フォルダ存在チェック
  If Dir(strFolderPath, vbDirectory) = "" Then
    MsgBox "VBAのエラーにより完成しなかった勤務表画像集のフォルダ-削除に成功しました!", _
         Buttons:=vbInformation, Title:="未完成の勤務表画像集のフォルダ-削除成功のご報告"
  End If
  
  Exit Function
  
deleteErr:
  MsgBox "完成しなかった勤務表画像集のフォルダ-削除時にエラーが発生し, 削除に失敗しました..." & vbCrLf & _
         "エラー原因は以下の通りです。" & vbCrLf & vbCrLf & _
         "----------------------------------" & vbCrLf & _
         "エラー番号：" & Err.Number & vbCrLf & _
         "エラー詳細：" & Err.Description & vbCrLf & _
         "----------------------------------", _
         Buttons:=vbCritical, Title:="実行者の方向けのエラー表示"
 
End Function
Sub 勤務表画像保存()
  offset = 0
  new_sheet_name = "勤務表画像保存用シート"
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
  sheet_names(0) = "チーム間調整前総合勤務表"
  sheet_names(1) = "希望優先チーム間調整後総合勤務表"
  sheet_names(2) = "希望後回しチーム間調整後総合勤務表"
  
  'On Error GoTo ErrLabel
  
  'アクティブシートの取得
  Set active_ws = ActiveSheet
 
  'VBA高速化開始(特に警告メッセージの停止)
  Call high_speeding(True)
  'ブックの保護の解除
  Call book_pro_unpro(False)
  'シートの保護の解除
  For Each sht_name In sheet_names
    Call pro_unpro(False, sht_name)
  Next
  
  '共通部分の代入
    '勤務表対象月の読み取り
  Set ws = ThisWorkbook.Worksheets("チーム間調整前総合勤務表")
  table_date = ws.Cells(9, 4)
  table_date = Left(table_date, Len(table_date) - 3)
  table_date = Replace(table_date, "/", "年") & "月"
    '表の上の共通部分の範囲格納
  Set common_rng(0) = ws.Range(ws.Cells(16, 4), ws.Cells(18, 5))
  Set common_rng(1) = ws.Range(ws.Cells(16, 10), ws.Cells(18, 41))
  
  'シート毎に異なるセル内容(範囲)の読み取り
  For sht_name_index = 0 To 2
    'sht_AB_loop_num=1,2の時はsheet_name(0)を指定するようにする
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
      'Aチーム(上の欄)の人数を数えた後, セル範囲を格納しておく
      number_of_people = 0
      GoTo B_team_return_point_num_people
    End If
    number_of_people = 0
  Next
    
  'バックアップフォルダのパスを作成する(バックアップ用のフォルダを選択したフォルダと同じ階層に作成する為)
  back_up_dir_path = ThisWorkbook.Path & "\" & table_date & "用勤務表画像集_" & Format(Now(), "YYYY年MM月DD日HH時MM分SS秒") & "作成"
  MkDir (back_up_dir_path)
    
  '貼り付けと画像保存
  ThisWorkbook.Worksheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
  Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
  ws.Name = new_sheet_name
  For sht_name_index = 0 To 2
    Call rng_new_sht_pst(new_sheet_name, common_rng(0), 1, 1)
    Call rng_new_sht_pst(new_sheet_name, common_rng(1), 1, 3)
    Set ws = ThisWorkbook.Worksheets(new_sheet_name)
    For AB = 0 To 1
      If AB = 0 Then
        team = "Aチーム"
        write_offset = 3
      Else
        team = "Bチーム"
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
    table_date & "用" & sheet_names(sht_name_index) & "_" & Format(Now(), "YYYY年MM月DD日HH時MM分SS秒") & "作成", back_up_dir_path, _
    new_sheet_name)
    ws.Delete
  Next
  
  'シートの保護の再有効化
  For Each sht_name In sheet_names
    Call pro_unpro(True, sht_name)
  Next
  'ブックの保護の再有効化
  Call book_pro_unpro(True)
  'VBA高速化終了(特に警告メッセージの再開)
  Call high_speeding(False)
  
  '実行後のアクティブシートが実行前のアクティブシートと異なった場合,
  '実行前のアクティブシートを再度 , アクティブにする
  If active_ws.Name <> ActiveSheet.Name Then
    active_ws.Activate
  End If
  
  If Dir(back_up_dir_path, vbDirectory) <> "" Then
    MsgBox "勤務表画像保存完了しました！" & vbCrLf & "次の画面でご確認下さい！", Buttons:=vbInformation, Title:="実行者の方へのメッセージ"
    'フォルダを開く
    Shell "C:\Windows\Explorer.exe " & back_up_dir_path, vbNormalFocus
  End If
  
  Exit Sub
  
ErrLabel:
  '勤務表画像保存用シートがあったら, 削除する
  If exist_sht(new_sheet_name) Then
    ThisWorkbook.Worksheets(new_sheet_name).Delete
  End If
   'シートの保護の再有効化
  For Each sht_name In sheet_names
    Call pro_unpro(True, sht_name)
  Next
  'ブックの保護の再有効化
  Call book_pro_unpro(True)
  'VBA高速化終了(特に警告メッセージの再開)
  Call high_speeding(False)
  '実行後のアクティブシートが実行前のアクティブシートと異なった場合,
  '実行前のアクティブシートを再度 , アクティブにする
  If active_ws.Name <> ActiveSheet.Name Then
    active_ws.Activate
  End If
  If Not err_non_people Then
    MsgBox "「勤務表画像保存」VBA実行中にエラーが発生しました！" _
    & vbCrLf & "再度, 「勤務表画像保存」のボタンを押し, 実行してみて下さい!" _
    & vbCrLf & "既に複数回, 実行したのに, 画像保存が出来ない場合は, スクリーンショットによる勤務表画像保存をご検討下さい。" _
    & vbCrLf & "エラー発生アプリ: " & Err.Source _
    & vbCrLf & "エラー番号: " & Err.Number _
    & vbCrLf & "エラー内容: " & Err.Description, _
    Buttons:=vbCritical, Title:="実行者の方向けのエラー表示"
  Else
    If offset = 19 Then
      err_team_name = "Aチーム"
    ElseIf offset = 79 Then
      err_team_name = "Bチーム"
    End If
    MsgBox "以下の手順①～⑤を行って下さい！" _
    & vbCrLf & "①" & err_team_name & "の勤務者の氏名を「" & err_team_name & "用勤務希望表」で入力する " _
    & vbCrLf & "②" & "「" & err_team_name & "用勤務表自動作成実行」ボタンを押す" _
    & vbCrLf & "③" & "「チーム間調整前総合勤務表」を完成させる" _
    & vbCrLf & "④" & "「チーム間調整勤務表自動作成実行」ボタンを押し," _
    & vbCrLf & "　「希望優先(後回し)総合勤務表」も完成させる" _
    & vbCrLf & "⑤" & "この「勤務表画像保存」ボタンを再度, 押す", _
    Buttons:=vbCritical, Title:="実行者の方へのメッセージ"
  End If
  'back_up_dir_pathに, 生成されたフォルダー名が代入されていて
  If back_up_dir_path <> "" Then
    '画像保存用のフォルダがあったら, 削除する
    If Dir(back_up_dir_path, vbDirectory) <> "" Then
      Call imgFolderDelete(back_up_dir_path)
    End If
  End If
End Sub
  


