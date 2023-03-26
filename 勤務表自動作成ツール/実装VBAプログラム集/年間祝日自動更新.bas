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
Private Sub Auto_Open()

    Dim Write_Sheet(4) As String
    Dim sheet_name As Variant
    Dim ws As Worksheet
    Dim QT As QueryTable
    Write_Sheet(0) = "Aチーム用勤務希望表"
    'Write_Sheet(1) = "Bチーム用勤務希望表"
    'Write_Sheet(2) = "チーム間調整前総合勤務表"
    'Write_Sheet(3) = "希望優先チーム間調整後総合勤務表"
    'Write_Sheet(4) = "希望後回しチーム間調整後総合勤務表"
    Dim eligible_year As String
    Dim wrong_holidays_sheets As New Collection
    Dim msg_for_sheet_name As String
    
    On Error GoTo ErrLabel
    
    'VBA高速化を開始する
    Call high_speeding(True)
      
    '設定しているシート内の祝日の対象年を全て確認し, 違う年が対象だった場合, そのシート名を追加する
    For Each sheet_name In Write_Sheet
      If sheet_name <> "" Then
        Set ws = ThisWorkbook.Sheets(sheet_name)
        eligible_year = ws.Cells(160, 53)
        If CStr(Year(Now)) <> Left(eligible_year, InStr(eligible_year, "年") - 1) Then
          wrong_holidays_sheets.Add sheet_name
        End If
      End If
    Next
  
    '設定している全てのシート内の祝日がブックを開いた時の年の祝日だったら
    If wrong_holidays_sheets.count = 0 Then
      'VBA高速化を終了し
      Call high_speeding(False)
      '祝日の自動更新が関係ない地点まで移動する
      GoTo holiday_end_point
    End If
    
    '以降, '設定している全てのシートで1シート以上, 祝日がブックを開いた時の年の祝日と違った場合のみ実行する
    For Each sheet_name In wrong_holidays_sheets
      'シートの保護を解除する
      Call pro_unpro(False, sheet_name)
      Set ws = ThisWorkbook.Sheets(sheet_name)
      '記載する祝日の対象年の書き込み
      ws.Cells(159, 53) = "対象年"
      ws.Cells(160, 53) = CStr(Year(Now)) & "年"
      ws.Range(ws.Cells(159, 53), ws.Cells(160, 53)).Borders.LineStyle = xlContinuous
      '祝日を取得し, 書き込む
      Set QT = ws.QueryTables.Add _
      (Connection:="URL;https://www8.cao.go.jp/chosei/shukujitsu/gaiyou.html", Destination:=ws.Range("BA162")) '祝日のデータがある内閣府のHPのURLを指定
      With QT
        '行列を追加せずセルを上書きする設定にしている
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
      'シートの保護を再度,有効にする
      Call pro_unpro(True, sheet_name)
    Next
    
    'メッセージ用に未更新→更新したシートの名前を文字列として追加する
    'For Each sheet_name In wrong_holidays_sheets
    '  msg_for_sheet_name = msg_for_sheet_name & vbCrLf & sheet_name
    'Next
    
    'VBA高速化を終了する
    Call high_speeding(False)
    MsgBox "年間祝日を自動更新しました！", Buttons:=vbInformation, Title:="ファイルを開いた方へのメッセージ"
    'MsgBox "以下のシートでの年間祝日を自動更新しました！" & msg_for_sheet_name, Buttons:=vbInformation, Title:="ファイルを開いた方へのメッセージ"
    
holiday_end_point:
   
    If Application.CellDragAndDrop = True Then
      Application.CellDragAndDrop = False
    End If
    
    Exit Sub
    
ErrLabel:
  For Each sheet_name In Write_Sheet
    'シートの保護を再度,有効にする
    Call pro_unpro(True, sheet_name)
  Next
  'VBA高速化を終了する
  Call high_speeding(False)
  MsgBox "「年間祝日自動更新」(Auto_Open)VBA実行中にエラーが発生しました！" _
  & vbCrLf & "シート「Aチーム用勤務希望表」の「祝日, 日付一致確認欄」をご確認の上, このExcelファイルを「上書き保存」し, 再度, 開いて下さい！" _
  & vbCrLf & "尚, 既に複数回、試しているのに, 本プログラムが正常に機能しない場合は, 年間祝日の手動更新をご検討下さい。" _
  & vbCrLf & "エラー発生アプリ: " & Err.Source _
  & vbCrLf & "エラー番号: " & Err.Number _
  & vbCrLf & "エラー内容: " & Err.Description, _
  Buttons:=vbCritical, Title:="ファイルを開いた方向けのエラー表示"
End Sub
