Attribute VB_Name = "m見出し改訂"
Option Explicit

'==================== API宣言 ====================
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" ( _
        ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, _
        ByVal uType As Long, ByVal wLanguageID As Long, ByVal dwMilliseconds As Long) As Long
#Else
    Private Declare Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" ( _
        ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, _
        ByVal uType As Long, ByVal wLanguageID As Long, ByVal dwMilliseconds As Long) As Long
#End If

'==================== 設 定 ====================
Private Const FILE_PATTERN As String = "*.xls*"         ' 対象拡張子
Private Const OLD_FOLDER As String = "■旧書式"          ' 元ファイルを保管するフォルダ名
'================================================

Public Sub 見出し改訂_フォルダ一括()
    Dim fldr As FileDialog, folderPath As String
    Dim logWs As Worksheet, logRow As Long
    Dim secLevel As MsoAutomationSecurity
    Dim t0 As Double
    Dim duplicates As String

    ' フォルダ選択
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "一括整形するフォルダを選択"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' 事前チェック：同名で拡張子違いのファイルがあるか
    duplicates = CheckDuplicateBaseNames(folderPath)
    If Len(duplicates) > 0 Then
        MsgBox "以下のファイルは同じ名前で拡張子が異なるため、処理できません。" & vbCrLf & _
               "手動で整理してから再実行してください。" & vbCrLf & vbCrLf & _
               duplicates, vbExclamation, "重複ファイル検出"
        Exit Sub
    End If

    ' ログシート用意
    Set logWs = PrepareLogSheet("見出し改訂Log")
    logRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
    t0 = Timer

    ' 安全系・高速化
    secLevel = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "見出し改訂中…"

    On Error GoTo Fin

    ' 再帰または単層で処理
    ProcessFolder folderPath, logWs, logRow

Fin:
    ' 復帰
    Application.AutomationSecurity = secLevel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False

    If Err.Number <> 0 Then
        MessageBoxTimeout 0, "一部エラーが発生しました: " & Err.Description, "エラー", vbExclamation, 0, 2000
    Else
        MessageBoxTimeout 0, "見出し改訂 完了", "完了", vbInformation, 0, 2000
    End If
End Sub

Private Sub ProcessFolder(ByVal folderPath As String, ByVal logWs As Worksheet, ByRef logRow As Long)
    Dim f As String
    Dim wb As Workbook, thisPath As String
    Dim originalPath As String, tempPath As String
    Dim oldFolderPath As String, oldFilePath As String
    Dim baseName As String, ext As String
    Dim sht As Worksheet, processed As Boolean
    Dim fso As Object
    Dim files As Collection
    Dim fileName As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    thisPath = ThisWorkbook.FullName

    ' 「旧書式」フォルダのパス
    oldFolderPath = folderPath & OLD_FOLDER & "\"

    ' ファイル一覧を先に全て取得（Dir()ループ中のファイル変更問題を回避）
    Set files = New Collection
    f = Dir(folderPath & FILE_PATTERN)
    Do While Len(f) > 0
        files.Add f
        f = Dir()
    Loop

    ' 取得済みリストをループして処理
    For Each fileName In files
        f = CStr(fileName)
        ' 自分自身はスキップ
        If folderPath & f <> thisPath Then
            On Error GoTo HandleErr

            originalPath = folderPath & f
            baseName = fso.GetBaseName(f)
            ext = LCase$(fso.GetExtensionName(f))

            ' 既に「旧書式」フォルダに同名ファイルが存在する場合はスキップ
            oldFilePath = oldFolderPath & f
            If fso.FileExists(oldFilePath) Then
                GoTo NextFile
            End If

            ' ファイルを開く
            Set wb = Workbooks.Open(originalPath, ReadOnly:=False, UpdateLinks:=0)

            ' 全シートをループ、2行目までにデータがあるシートのみ処理
            Dim processedCount As Long
            processedCount = 0
            For Each sht In wb.Worksheets
                If HasDataInFirstTwoRows(sht) Then
                    ProcessOneSheet sht
                    processedCount = processedCount + 1
                End If
            Next
            processed = (processedCount > 0)

            ' 保存前に全シートのセル選択とスクロール位置をリセット
            Dim shtReset As Worksheet
            For Each shtReset In wb.Worksheets
                shtReset.Activate
                shtReset.Range("A1").Select
                ActiveWindow.ScrollRow = 1
                ActiveWindow.ScrollColumn = 1
            Next
            wb.Worksheets(1).Activate  ' 最初のシートをアクティブに

            ' 一時ファイルとして保存
            tempPath = folderPath & baseName & "_temp_" & Format$(Now, "hhmmss") & ".xlsx"
            wb.SaveAs Filename:=tempPath, FileFormat:=xlOpenXMLWorkbook
            wb.Close SaveChanges:=False

            ' 「旧書式」フォルダがなければ作成
            If Not fso.FolderExists(oldFolderPath) Then
                fso.CreateFolder oldFolderPath
            End If

            ' 元ファイルを「旧書式」フォルダに移動
            fso.MoveFile originalPath, oldFilePath

            ' 一時ファイルを元のファイル名（.xlsx）にリネーム
            Dim newPath As String
            newPath = folderPath & baseName & ".xlsx"
            fso.MoveFile tempPath, newPath

            ' ログ
            logRow = logRow + 1
            logWs.Cells(logRow, 1).Value = baseName
            logWs.Cells(logRow, 2).Value = ext
            logWs.Cells(logRow, 3).Value = "xlsx"
            logWs.Cells(logRow, 4).Value = processedCount
            logWs.Cells(logRow, 5).Value = folderPath
        End If
NextFile:
        On Error GoTo 0
        DoEvents
    Next fileName
    Exit Sub

HandleErr:
    On Error GoTo 0
    If Not wb Is Nothing Then On Error Resume Next: wb.Close SaveChanges:=False: On Error GoTo 0
    ' 一時ファイルが残っていたら削除
    If Len(tempPath) > 0 Then
        If fso.FileExists(tempPath) Then fso.DeleteFile tempPath
    End If
    Resume NextFile
End Sub

Private Sub ProcessOneSheet(ByVal ws As Worksheet)
    On Error GoTo Fail
    FormatFrontPage ws
    Exit Sub
Fail:
    Err.Raise Err.Number, "ProcessOneSheet(" & ws.Name & ")", Err.Description
End Sub

' 2行目までにデータがあるかチェック（結合セルも左上セルで判定される）
Private Function HasDataInFirstTwoRows(ByVal ws As Worksheet) As Boolean
    Dim cell As Range
    On Error Resume Next
    For Each cell In ws.Range("1:2").Cells
        If Len(Trim$(CStr(cell.Value))) > 0 Then
            HasDataInFirstTwoRows = True
            Exit Function
        End If
    Next
    HasDataInFirstTwoRows = False
    On Error GoTo 0
End Function

' 同じベース名で拡張子が異なるファイルを検出
Private Function CheckDuplicateBaseNames(ByVal folderPath As String) As String
    Dim fso As Object, f As String
    Dim dict As Object  ' Dictionary: ベース名 → 拡張子のDictionary
    Dim baseName As String, ext As String
    Dim result As String
    Dim key As Variant, exts As Object, extKey As Variant
    Dim extList As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set dict = CreateObject("Scripting.Dictionary")

    ' ファイルをスキャンしてベース名ごとに拡張子を収集
    f = Dir(folderPath & FILE_PATTERN)
    Do While Len(f) > 0
        baseName = fso.GetBaseName(f)
        ext = LCase$(fso.GetExtensionName(f))

        If Not dict.Exists(baseName) Then
            Set dict(baseName) = CreateObject("Scripting.Dictionary")
        End If
        dict(baseName)(ext) = True  ' 拡張子をキーとして登録
        f = Dir()
    Loop

    ' 複数の拡張子があるベース名を抽出
    result = ""
    For Each key In dict.Keys
        Set exts = dict(key)
        If exts.Count > 1 Then
            extList = ""
            For Each extKey In exts.Keys
                If Len(extList) > 0 Then extList = extList & ", "
                extList = extList & "." & extKey
            Next
            result = result & key & " (" & extList & ")" & vbCrLf
        End If
    Next

    CheckDuplicateBaseNames = result
End Function

Private Function PrepareLogSheet(ByVal Name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(Name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = Name
        ' 見出しと書式を設定（新規作成時のみ）
        ws.[A1:E1].Value = Array("ファイル名", "変換前", "変換後", "シート数", "フォルダパス")
        ' 列幅設定
        ws.Columns("A").ColumnWidth = 30
        ws.Columns("B").ColumnWidth = 5
        ws.Columns("C").ColumnWidth = 5
        ws.Columns("D").ColumnWidth = 5
        ws.Columns("E").ColumnWidth = 50
        ' 縮小して全体を表示
        ws.Columns("A:E").ShrinkToFit = True
    End If
    Set PrepareLogSheet = ws
End Function

'==================== 整形本体 ====================
Private Sub FormatFrontPage(ByVal ws As Worksheet)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanFail

    ' 1) 2〜6行コピー→1行目の前に挿入
    ws.Rows("2:6").Copy
    ws.Rows("1:1").Insert Shift:=xlDown

    ' 2) 7〜11行：結合解除＋値&書式クリア
    With ws.Rows("7:11")
        .UnMerge
        .Clear
    End With

    ' 3) B7:AK11：結合解除、内部罫線クリア、フォント設定
    With ws.Range("B7:AK11")
        .UnMerge
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        With .Font
            .Name = "ＭＳ Ｐゴシック"
            .Size = 10
            .Bold = False
        End With
    End With

    ' 4) 7行目と8行目の間に1行挿入
    ws.Rows("8:8").Insert Shift:=xlDown

    ' 5) 範囲結合・中央・縮小
    Dim rngs As Variant, i As Long
    rngs = Array( _
        "B7:D8", "B9:D10", "B11:D12", _
        "E7:K8", "E9:K10", "E11:K12", _
        "L7:AB8", "L9:AB12", _
        "AC7:AE8", "AC9:AE9", "AC10:AE12", _
        "AF7:AK8", "AF9:AH9", "AF10:AH12", _
        "AI9:AK9", "AI10:AK12")
    For i = LBound(rngs) To UBound(rngs)
        With ws.Range(rngs(i))
            .UnMerge
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .ShrinkToFit = True
        End With
    Next i

    ' 6) ラベル
    ws.Range("B7").Value = "制定日"
    ws.Range("B9").Value = "改定日"
    ws.Range("B11").Value = "文書番号"
    ws.Range("E7").Value = "作成日"
    ws.Range("L7").Value = "文書名"
    ws.Range("AC7").Value = "頁"
    ws.Range("AC9:AE9").Value = "承認"
    ws.Range("AF9:AH9").Value = "照査"
    ws.Range("AI9:AK9").Value = "作成"

    ' 6.5) E7,E9の書式設定（日付用）
    With ws.Range("E7")
        .Font.Size = 11
        .NumberFormatLocal = "yyyy""年""m""月""d""日"""
    End With
    With ws.Range("E9")
        .Font.Size = 11
        .NumberFormatLocal = "yyyy""年""m""月""d""日"""
    End With

    ' 7) L9にB1の文字列、フォント指定
    ws.Range("L9").Value = ws.Range("B1").Value
    With ws.Range("L9").Font
        .Name = "HG創英角ｺﾞｼｯｸUB"
        .Bold = False
        .Size = 18
    End With

    ' 8) E7に AC1,AF1,AI1 から "yyyy年m月d日"
    Dim y As Long, m As Long, d As Long
    y = GetYearValue(ws.Range("AC1").Value)
    m = GetMonthValue(ws.Range("AF1").Value)
    d = GetDayValue(ws.Range("AI1").Value)
    If y > 0 And m > 0 And d > 0 Then
        ws.Range("E7").Value = DateSerial(y, m, d)
        ws.Range("E7").NumberFormatLocal = "yyyy""年""m""月""d""日"""
    Else
        ws.Range("E7").Value = CStr(y) & "年" & CStr(m) & "月" & CStr(d) & "日"
    End If

    ' 9) B7:AK12 罫線：内側=通常、外枠=中太線
    With ws.Range("B7:AK12")
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
    End With

    ' 10) 1〜5行を削除
    ws.Rows("1:5").Delete Shift:=xlUp

    ' 11) 印刷設定（行削除後の最終座標で指定）
    With ws.PageSetup
        .PrintArea = "B2:AK66"
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        .Zoom = 100
    End With

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Err.Raise Err.Number, "FormatFrontPage(" & ws.Name & ")", Err.Description
End Sub

'---- ヘルパー：年・月・日 ----
Private Function GetYearValue(ByVal v As Variant) As Long
    On Error GoTo Fallback
    If IsDate(v) Then GetYearValue = Year(CDate(v)): Exit Function
    If IsNumeric(v) Then
        Dim n As Double: n = CDbl(v)
        If n >= 1900 And n < 10000 Then GetYearValue = CLng(n): Exit Function
    End If
Fallback:
    GetYearValue = Val(v)
End Function

Private Function GetMonthValue(ByVal v As Variant) As Long
    On Error GoTo Fallback
    If IsDate(v) Then GetMonthValue = Month(CDate(v)): Exit Function
    If IsNumeric(v) Then GetMonthValue = CLng(v): Exit Function
Fallback:
    GetMonthValue = Val(v)
End Function

Private Function GetDayValue(ByVal v As Variant) As Long
    On Error GoTo Fallback
    If IsDate(v) Then GetDayValue = Day(CDate(v)): Exit Function
    If IsNumeric(v) Then GetDayValue = CLng(v): Exit Function
Fallback:
    GetDayValue = Val(v)
End Function
