VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1
   Caption         =   "入力支援"
   ClientHeight    =   4008
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7356
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm1 - インポートコード
Option Explicit

' クラスハンドラ
Private dateHandler As CTextBoxEvent    ' 日付用ハンドラ
Private monthHandler As CTextBoxEvent   ' 月間後用ハンドラ
Private lotHandler As CTextBoxEvent     ' ロット用ハンドラ
Private qtyHandler As CTextBoxEvent     ' 数量用ハンドラ

' フォーム初期化時の処理
Private Sub UserForm_Initialize()
    ' コンボボックス1（工程）の初期化
    With ComboBox1
        .AddItem "成形"
        .AddItem "塗装"
        .AddItem "モール"
        .AddItem "加工"
        .IMEMode = fmIMEModeHiragana ' 全角ひらがなモード
    End With

    ' コンボボックス2のIMEモード設定
    ComboBox2.IMEMode = fmIMEModeHiragana ' 全角ひらがなモード

    ' テキストボックスのIMEモード設定と制御設定
    With TextBox1
        .IMEMode = fmIMEModeDisable ' 半角英数モード
        .Tag = "Date"               ' 日付入力用のタグ設定
    End With

    With TextBox2
        .IMEMode = fmIMEModeDisable ' 半角英数モード
        .Tag = "Month"              ' 月入力用のタグ設定
    End With

    With TextBox3
        .IMEMode = fmIMEModeDisable ' 半角英数モード
        .Tag = "Lot"                ' ロット入力用のタグ設定
    End With

    With TextBox4
        .IMEMode = fmIMEModeDisable ' 半角英数モード
        .Tag = "Quantity"           ' 数量入力用のタグ設定
    End With

    ' TextBox5（品番展開）のフォント設定とマルチライン設定
    With TextBox5
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 12
        .Font.Bold = True
        .MultiLine = True           ' マルチライン設定を有効にする
        .IMEMode = fmIMEModeDisable ' 半角英数モード
    End With

    ' クラスハンドラの初期化
    InitializeTextBoxHandlers

    ' ComboBox1にフォーカスを設定
    ComboBox1.SetFocus

    ' 初期IME設定
    SetJapaneseIME ComboBox1
End Sub

' テキストボックスハンドラの初期化
Private Sub InitializeTextBoxHandlers()
    ' 日付入力用ハンドラ
    Set dateHandler = New CTextBoxEvent
    Set dateHandler.TB = TextBox1

    ' 月間後用ハンドラ
    Set monthHandler = New CTextBoxEvent
    Set monthHandler.TB = TextBox2

    ' ロット用ハンドラ
    Set lotHandler = New CTextBoxEvent
    Set lotHandler.TB = TextBox3

    ' 数量用ハンドラ
    Set qtyHandler = New CTextBoxEvent
    Set qtyHandler.TB = TextBox4
End Sub

' ユーザーフォームのキーダウンイベント
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' IME制御（キー押下時IMEモードを強制再設定）
    HandleFormKeyDown Me

    ' タブキーによるフォーカス移動時の入力検証
    If KeyCode = vbKeyTab Then
        Dim ctrl As MSForms.Control
        Set ctrl = Me.ActiveControl

        If TypeOf ctrl Is MSForms.TextBox Then
            Select Case ctrl.Tag
                Case "Date"
                    ' 日付入力チェック
                    If ctrl.Value <> "" Then
                        Dim dateValue As String
                        dateValue = StrConv(ctrl.Value, vbNarrow)
                        If Not IsDate(dateValue) Then
                            MsgBox "日付は「yyyy/m/d」または「m/d」形式で入力してください。", vbExclamation
                            KeyCode = 0
                            ctrl.Value = ""
                            ctrl.SetFocus
                        End If
                    End If

                Case "Month"
                    ' 月間後入力チェック
                    If ctrl.Value <> "" Then
                        If Not IsNumeric(ctrl.Value) Or Val(ctrl.Value) <> Int(Val(ctrl.Value)) Or Val(ctrl.Value) < 0 Or Val(ctrl.Value) > 12 Then
                            MsgBox "月間後は0〜12の数字を入力してください。", vbExclamation
                            KeyCode = 0
                            ctrl.Value = ""
                            ctrl.SetFocus
                        End If
                    End If

                Case "Lot"
                    ' ロット入力チェック
                    If ctrl.Value <> "" Then
                        If Not IsNumeric(ctrl.Value) Or Val(ctrl.Value) <> Int(Val(ctrl.Value)) Or Val(ctrl.Value) <= 0 Then
                            MsgBox "ロットは正の整数を入力してください。", vbExclamation
                            KeyCode = 0
                            ctrl.Value = ""
                            ctrl.SetFocus
                        End If
                    End If

                Case "Quantity"
                    ' 数量入力チェック
                    If ctrl.Value <> "" Then
                        If Not IsNumeric(ctrl.Value) Or Val(ctrl.Value) <> Int(Val(ctrl.Value)) Or Val(ctrl.Value) <= 0 Then
                            MsgBox "数量は正の整数を入力してください。", vbExclamation
                            KeyCode = 0
                            ctrl.Value = ""
                            ctrl.SetFocus
                        End If
                    End If
            End Select
        End If
    End If
End Sub

' ComboBox1にフォーカスが移った時
Private Sub ComboBox1_GotFocus()
    ComboBox1.IMEMode = fmIMEModeHiragana ' 全角ひらがなモード
    ' IMEモードを強制的に日本語入力ONへ
    SetJapaneseIME ComboBox1
End Sub

' ComboBox1からのフォーカス移動を検知
Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' CapsLockなどのキー入力時に日本語入力モードを維持
    SetJapaneseIME ComboBox1
End Sub

' ComboBox1のキー押下を検知
Private Sub ComboBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' キー入力中も日本語入力モードを維持
    SetJapaneseIME ComboBox1
End Sub

' ComboBox2にフォーカスが移った時
Private Sub ComboBox2_GotFocus()
    ComboBox2.IMEMode = fmIMEModeHiragana ' 全角ひらがなモード
    ' IMEモードを強制的に日本語入力ONへ
    SetJapaneseIME ComboBox2
End Sub

' ComboBox2からのフォーカス移動を検知
Private Sub ComboBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' CapsLockなどのキー入力時に日本語入力モードを維持
    SetJapaneseIME ComboBox2
End Sub

' ComboBox2のキー押下を検知
Private Sub ComboBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' キー入力中も日本語入力モードを維持
    SetJapaneseIME ComboBox2
End Sub

' TextBox1にフォーカスが移った時
Private Sub TextBox1_GotFocus()
    TextBox1.IMEMode = fmIMEModeDisable ' 半角英数モード
    ' IMEモードを強制的に日本語入力OFFへ
    SetAlphaIME TextBox1
End Sub

' TextBox2にフォーカスが移った時
Private Sub TextBox2_GotFocus()
    TextBox2.IMEMode = fmIMEModeDisable ' 半角英数モード
    ' IMEモードを強制的に日本語入力OFFへ
    SetAlphaIME TextBox2
End Sub

' TextBox3にフォーカスが移った時
Private Sub TextBox3_GotFocus()
    TextBox3.IMEMode = fmIMEModeDisable ' 半角英数モード
    ' IMEモードを強制的に日本語入力OFFへ
    SetAlphaIME TextBox3
End Sub

' TextBox4にフォーカスが移った時
Private Sub TextBox4_GotFocus()
    TextBox4.IMEMode = fmIMEModeDisable ' 半角英数モード
    ' IMEモードを強制的に日本語入力OFFへ
    SetAlphaIME TextBox4
End Sub

' コンボボックス1（工程）の変更処理
Private Sub ComboBox1_Change()
    ' コンボボックス2（品番）の初期化
    ComboBox2.Clear
    ComboBox3.Clear

    Select Case ComboBox1.Value
        Case "成形", "塗装"
            ComboBox2.AddItem "ノアFr"
            ComboBox2.AddItem "ノアRr"
            ComboBox2.AddItem "アルFr"
            ComboBox2.AddItem "アルRr"
        Case "モール"
            ComboBox2.AddItem "アル"     ' 新規に追加
            ComboBox2.AddItem "アルFr"
            ComboBox2.AddItem "アルRr"
        Case "加工"
            ComboBox2.AddItem "ノア"
            ComboBox2.AddItem "アル"
    End Select
End Sub

' ComboBox1の入力検証（フォーカス離脱時）
Private Sub ComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ComboBox1.Value = "" Then Exit Sub

    Dim validValue As Boolean
    validValue = False

    ' 有効な値かチェック
    Select Case ComboBox1.Value
        Case "成形", "塗装", "モール", "加工"
            validValue = True
    End Select

    If Not validValue Then
        MsgBox "工程は「成形」「塗装」「モール」「加工」から選択してください。", vbExclamation
        Cancel = True
        ComboBox1.SetFocus
    End If
End Sub

' コンボボックス2（品番）の変更処理
Private Sub ComboBox2_Change()
    ' コンボボックス3（品番末尾）の初期化
    ComboBox3.Clear

    If InStr(ComboBox2.Value, "ノア") > 0 Then
        ComboBox3.AddItem "30"
        ComboBox3.AddItem "40"
        ComboBox3.AddItem "50"
        ComboBox3.AddItem "60"
    ElseIf InStr(ComboBox2.Value, "アル") > 0 Then
        ComboBox3.AddItem "20"
        ComboBox3.AddItem "30"
        ComboBox3.AddItem "40"
        ComboBox3.AddItem "50"
        ComboBox3.AddItem "60"
    End If
End Sub

' ComboBox2の入力検証（フォーカス離脱時）
Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ComboBox2.Value = "" Then Exit Sub
    If ComboBox1.Value = "" Then Exit Sub

    Dim validValue As Boolean
    validValue = False

    ' ComboBox1の値に応じて有効な値かチェック
    Select Case ComboBox1.Value
        Case "成形", "塗装"
            Select Case ComboBox2.Value
                Case "ノアFr", "ノアRr", "アルFr", "アルRr"
                    validValue = True
            End Select

        Case "モール"
            Select Case ComboBox2.Value
                Case "アル", "アルFr", "アルRr"  ' 「アル」を追加
                    validValue = True
            End Select

        Case "加工"
            Select Case ComboBox2.Value
                Case "ノア", "アル"
                    validValue = True
            End Select
    End Select

    If Not validValue Then
        MsgBox "選択された工程に対して無効な品番です。リストから選択してください。", vbExclamation
        Cancel = True
        ComboBox2.SetFocus
    End If
End Sub

' ComboBox3の入力検証（フォーカス離脱時）
Private Sub ComboBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ComboBox3.Value = "" Then Exit Sub
    If ComboBox2.Value = "" Then Exit Sub

    Dim validValue As Boolean
    validValue = False

    ' ComboBox2の値に応じて有効な値かチェック
    If InStr(ComboBox2.Value, "ノア") > 0 Then
        Select Case ComboBox3.Value
            Case "30", "40", "50", "60"
                validValue = True
        End Select
    ElseIf InStr(ComboBox2.Value, "アル") > 0 Then
        Select Case ComboBox3.Value
            Case "20", "30", "40", "50", "60"
                validValue = True
        End Select
    End If

    If Not validValue Then
        MsgBox "選択された品番に対して無効な品番末尾です。リストから選択してください。", vbExclamation
        Cancel = True
        ComboBox3.SetFocus
    End If
End Sub

' 日付テキストボックスのフォーカス離脱処理
Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' 入力内容が空の場合は何もしない
    If TextBox1.Value = "" Then Exit Sub

    Dim dateValue As String
    dateValue = TextBox1.Value

    ' 全角があれば半角に変換
    dateValue = StrConv(dateValue, vbNarrow)

    ' yyyy/m/d形式または m/d形式かチェック
    If IsDate(dateValue) Then
        ' 日付形式の場合、yyyy/m/d形式に変換
        TextBox1.Value = Format(CDate(dateValue), "yyyy/m/d")
    Else
        ' 日付形式でない場合はエラーメッセージ
        MsgBox "日付は「yyyy/m/d」または「m/d」形式で入力してください。", vbExclamation
        TextBox1.Value = ""
        Cancel = True  ' フォーカスを維持
        TextBox1.SetFocus
    End If
End Sub

' TextBox2（月間後）の入力検証
Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox2.Value = "" Then Exit Sub

    ' 全角を半角に変換
    TextBox2.Value = StrConv(TextBox2.Value, vbNarrow)

    ' 数値かチェック
    If Not IsNumeric(TextBox2.Value) Then
        MsgBox "月間後は0から12までの数字を入力してください。", vbExclamation
        TextBox2.Value = ""
        Cancel = True
        TextBox2.SetFocus
        Exit Sub
    End If

    ' 整数かチェック
    If Int(Val(TextBox2.Value)) <> Val(TextBox2.Value) Then
        MsgBox "月間後は0から12までの数字を入力してください。", vbExclamation
        TextBox2.Value = ""
        Cancel = True
        TextBox2.SetFocus
        Exit Sub
    End If

    ' 範囲内かチェック
    Dim monthValue As Integer
    monthValue = CInt(TextBox2.Value)

    If monthValue < 0 Or monthValue > 12 Then
        MsgBox "月間後は0から12までの数字を入力してください。", vbExclamation
        TextBox2.Value = ""
        Cancel = True
        TextBox2.SetFocus
    End If
End Sub

' TextBox3（ロット）の入力検証
Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox3.Value = "" Then Exit Sub

    ' 全角を半角に変換
    TextBox3.Value = StrConv(TextBox3.Value, vbNarrow)

    ' 数値かチェック
    If Not IsNumeric(TextBox3.Value) Then
        MsgBox "ロットは正の整数を入力してください。", vbExclamation
        TextBox3.Value = ""
        Cancel = True
        TextBox3.SetFocus
        Exit Sub
    End If

    ' 正の整数かチェック
    If Int(Val(TextBox3.Value)) <> Val(TextBox3.Value) Or Val(TextBox3.Value) <= 0 Then
        MsgBox "ロットは正の整数を入力してください。", vbExclamation
        TextBox3.Value = ""
        Cancel = True
        TextBox3.SetFocus
    End If
End Sub

' TextBox4（数量）の入力検証 - KeyPressイベントを追加
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 数字とBackspaceのみ許可
    Select Case KeyAscii
        Case 48 To 57  ' 0-9許可する
        Case 8         ' Backspace許可する
        Case Else
            KeyAscii = 0  ' それ以外は無効化
    End Select
End Sub

' TextBox4（数量）の入力検証
Private Sub TextBox4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox4.Value = "" Then Exit Sub

    ' 全角を半角に変換
    TextBox4.Value = StrConv(TextBox4.Value, vbNarrow)

    ' 数値かチェック
    If Not IsNumeric(TextBox4.Value) Then
        MsgBox "数量は正の整数を入力してください。", vbExclamation
        TextBox4.Value = ""
        Cancel = True
        TextBox4.SetFocus
        Exit Sub
    End If

    ' 正の整数かチェック
    If Int(Val(TextBox4.Value)) <> Val(TextBox4.Value) Or Val(TextBox4.Value) <= 0 Then
        MsgBox "数量は正の整数を入力してください。", vbExclamation
        TextBox4.Value = ""
        Cancel = True
        TextBox4.SetFocus
    End If
End Sub

' 展開ボタンクリック時の処理
Private Sub CommandButton1_Click()
    ' 入力値の検証
    If TextBox1.Value = "" Then
        MsgBox "日付を入力してください。", vbExclamation
        TextBox1.SetFocus
        Exit Sub
    End If

    ' 日付の最終確認（フォーマット変換を含む）
    Dim dateValue As String
    dateValue = StrConv(TextBox1.Value, vbNarrow)  ' 全角→半角変換

    If IsDate(dateValue) Then
        TextBox1.Value = Format(CDate(dateValue), "yyyy/m/d")
    Else
        MsgBox "日付は「yyyy/m/d」または「m/d」形式で入力してください。", vbExclamation
        TextBox1.SetFocus
        Exit Sub
    End If

    If ComboBox1.Value = "" Then
        MsgBox "工程を選択してください。", vbExclamation
        ComboBox1.SetFocus
        Exit Sub
    End If

    If ComboBox2.Value = "" Then
        MsgBox "品番を選択してください。", vbExclamation
        ComboBox2.SetFocus
        Exit Sub
    End If

    If ComboBox3.Value = "" Then
        MsgBox "品番末尾を選択してください。", vbExclamation
        ComboBox3.SetFocus
        Exit Sub
    End If

    If TextBox2.Value = "" Then
        MsgBox "月間後を入力してください。", vbExclamation
        TextBox2.SetFocus
        Exit Sub
    End If

    If TextBox3.Value = "" Then
        MsgBox "ロットを入力してください。", vbExclamation
        TextBox3.SetFocus
        Exit Sub
    End If

    If TextBox4.Value = "" Then
        MsgBox "数量を入力してください。", vbExclamation
        TextBox4.SetFocus
        Exit Sub
    End If

    ' 進捗表示
    Application.StatusBar = "品番展開を生成中..."

    ' 品番展開の生成（改行対応版）
    Dim displayText As String

    ' 工程と品番の組み合わせによって展開方法を変更
    If ComboBox1.Value = "加工" Or (ComboBox1.Value = "モール" And ComboBox2.Value = "アル") Then
        ' 4行表示 - 改行を確実に入れる
        displayText = TextBox1.Value & "-" & ComboBox2.Value & "FrLH-" & ComboBox3.Value & "-" & _
                      TextBox2.Value & "-" & TextBox3.Value & "-" & ComboBox1.Value & "-" & TextBox4.Value

        ' 各行の間に確実に改行を入れる
        displayText = displayText & vbNewLine

        displayText = displayText & TextBox1.Value & "-" & ComboBox2.Value & "FrRH-" & ComboBox3.Value & "-" & _
                      TextBox2.Value & "-" & TextBox3.Value & "-" & ComboBox1.Value & "-" & TextBox4.Value

        displayText = displayText & vbNewLine

        displayText = displayText & TextBox1.Value & "-" & ComboBox2.Value & "RrLH-" & ComboBox3.Value & "-" & _
                      TextBox2.Value & "-" & TextBox3.Value & "-" & ComboBox1.Value & "-" & TextBox4.Value

        displayText = displayText & vbNewLine

        displayText = displayText & TextBox1.Value & "-" & ComboBox2.Value & "RrRH-" & ComboBox3.Value & "-" & _
                      TextBox2.Value & "-" & TextBox3.Value & "-" & ComboBox1.Value & "-" & TextBox4.Value
    Else
        ' 2行表示 - 改行を確実に入れる
        displayText = TextBox1.Value & "-" & ComboBox2.Value & "LH-" & ComboBox3.Value & "-" & _
                      TextBox2.Value & "-" & TextBox3.Value & "-" & ComboBox1.Value & "-" & TextBox4.Value

        ' 必ず改行を入れる
        displayText = displayText & vbNewLine

        displayText = displayText & TextBox1.Value & "-" & ComboBox2.Value & "RH-" & ComboBox3.Value & "-" & _
                      TextBox2.Value & "-" & TextBox3.Value & "-" & ComboBox1.Value & "-" & TextBox4.Value
    End If

    ' TextBox5に表示
    TextBox5.Value = displayText

    ' ステータスバーをクリア
    Application.StatusBar = False

    ' 転記ボタンにフォーカスを設定
    CommandButton2.SetFocus
End Sub

' 転記ボタンクリック時の処理
Private Sub CommandButton2_Click()
    ' 品番展開が未実行の場合
    If TextBox5.Value = "" Then
        MsgBox "先に展開ボタンを押して品番を展開してください。", vbExclamation
        Exit Sub
    End If

    ' パフォーマンス最適化
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 進捗表示
    Application.StatusBar = "「_ロット数量全」テーブルを検索中..."

    ' _ロット数量全テーブルの参照（修正版）
    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim tableFound As Boolean
    Dim insertData() As Variant
    Dim lines() As String
    Dim firstEmptyRow As Long
    Dim lastRow As Long
    Dim i As Long, j As Long

    tableFound = False

    On Error Resume Next
    ' すべてのワークシートを検索してテーブルを探す
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = "_ロット数量全" Then
                tableFound = True
                Exit For
            End If
        Next tbl
        If tableFound Then Exit For
    Next ws

    ' テーブルが存在しない場合はアンダースコアなしでも検索
    If Not tableFound Then
        For Each ws In ThisWorkbook.Worksheets
            For Each tbl In ws.ListObjects
                If tbl.Name = "ロット数量全" Then
                    tableFound = True
                    Exit For
                End If
            Next tbl
            If tableFound Then Exit For
        Next ws
    End If
    On Error GoTo 0

    If Not tableFound Then
        MsgBox "ロット数量全テーブルが見つかりません。テーブル名を確認してください。", vbCritical
        GoTo CleanExit
    End If

    Application.StatusBar = "空き行を検索中..."

    ' 最初の空白行を検索
    Dim dateCol As Range
    Set dateCol = tbl.ListColumns("日付").DataBodyRange

    firstEmptyRow = 0
    If Not dateCol Is Nothing Then
        Dim dateValues As Variant
        dateValues = dateCol.Value

        If Not IsEmpty(dateValues) Then
            For i = 1 To UBound(dateValues, 1)
                If IsEmpty(dateValues(i, 1)) Then
                    firstEmptyRow = i
                    Exit For
                End If
            Next i
        End If
    End If

    ' 空白行がない場合は新規行を追加
    If firstEmptyRow = 0 Then
        firstEmptyRow = tbl.ListRows.Count + 1
    End If

    ' 転記するデータを行ごとに分割
    lines = Split(TextBox5.Value, vbCrLf)

    ' 転記するデータの行数
    Dim lineCount As Long
    lineCount = UBound(lines) + 1

    ' 必要な行数を確保
    Dim currentRows As Long
    currentRows = tbl.ListRows.Count

    If firstEmptyRow + lineCount - 1 > currentRows Then
        ' 行が足りない→行を追加
        Dim rowsToAdd As Long
        rowsToAdd = (firstEmptyRow + lineCount - 1) - currentRows

        Application.StatusBar = "テーブルに行を追加中..."

        ' 複数行を一度に追加する場合の処理
        For i = 1 To rowsToAdd
            tbl.ListRows.Add
        Next i
    End If

    ' 配列を初期化（転記する行数 x 7列のサイズ）
    ReDim insertData(1 To lineCount, 1 To 7)

    Application.StatusBar = "データを転記用に準備中..."

    ' データを一括で配列に格納
    For i = 0 To UBound(lines)
        Dim parts() As String
        parts = Split(lines(i), "-")

        ' 配列にデータを格納
        insertData(i + 1, 1) = parts(0)  ' 日付
        insertData(i + 1, 2) = parts(1)  ' 品番
        insertData(i + 1, 3) = parts(2)  ' 品番末尾
        insertData(i + 1, 4) = parts(3)  ' 月間後
        insertData(i + 1, 5) = parts(4)  ' ロット
        insertData(i + 1, 6) = parts(5)  ' 工程
        insertData(i + 1, 7) = parts(6)  ' ロット量
    Next i

    Application.StatusBar = "データを転記中..."

    ' データを一括で転記
    Dim targetRange As Range
    Set targetRange = tbl.ListRows(firstEmptyRow).Range.Resize(lineCount, 7)
    targetRange.Value = insertData

    ' 最終行を保存（後で次のセルを選択するため）
    lastRow = firstEmptyRow + lineCount - 1

    ' 入力欄をクリア（工程と日付は保持）
    ClearInputFields

    ' アクティブシートのセルC1を選択
    ws.Activate
    ws.Range("C1").Select

    Application.StatusBar = "転記完了"
    Application.Run "ステータスバークリア"

CleanExit:
    ' 設定を元に戻す
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        Application.StatusBar = "エラー: " & Err.Description
        Application.Run "ステータスバークリア"
    End If
End Sub

' 終了ボタンクリック時の処理
Private Sub CommandButton3_Click()
    ' _ロット数量全テーブルの参照
    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim tableFound As Boolean
    Dim lastFilledRow As Long

    Application.StatusBar = "テーブル検索中..."

    tableFound = False

    On Error Resume Next
    ' すべてのワークシートを検索してテーブルを探す
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = "_ロット数量全" Then
                tableFound = True
                Exit For
            End If
        Next tbl
        If tableFound Then Exit For
    Next ws

    ' テーブルが存在しない場合はアンダースコアなしでも検索
    If Not tableFound Then
        For Each ws In ThisWorkbook.Worksheets
            For Each tbl In ws.ListObjects
                If tbl.Name = "ロット数量全" Then
                    tableFound = True
                    Exit For
                End If
            Next tbl
            If tableFound Then Exit For
        Next ws
    End If
    On Error GoTo 0

    ' テーブルが見つかった場合
    If tableFound Then
        Application.StatusBar = "最終入力行を検索中..."

        ' 日付列の最後の入力行を検索
        Dim dateCol As Range
        Set dateCol = tbl.ListColumns("日付").DataBodyRange

        lastFilledRow = 0
        If Not dateCol Is Nothing Then
            Dim i As Long
            For i = 1 To dateCol.Rows.Count
                If Not IsEmpty(dateCol.Cells(i, 1).Value) Then
                    lastFilledRow = i
                End If
            Next i
        End If

        ' 最後の入力行の次のセルを選択
        If lastFilledRow > 0 Then
            ws.Activate
            ' もし最後の行より下に行がある場合
            If lastFilledRow < tbl.ListRows.Count Then
                tbl.ListColumns("日付").DataBodyRange.Cells(lastFilledRow + 1, 1).Select
            Else
                ' 新しい行を追加して選択
                tbl.ListRows.Add
                tbl.ListColumns("日付").DataBodyRange.Cells(lastFilledRow + 1, 1).Select
            End If
        End If
    End If

    Application.StatusBar = False

    ' フォームを閉じる
    Unload Me
End Sub

' 入力欄をクリアする関数（工程と日付を保持）
Private Sub ClearInputFields()
    ' ComboBox1（工程）とTextBox1（日付）は保持
    TextBox2.Value = ""    ' 月間後
    TextBox3.Value = ""    ' ロット
    TextBox4.Value = ""    ' 数量
    TextBox5.Value = ""    ' 品番展開結果
    ComboBox2.Clear        ' 品番
    ComboBox3.Clear        ' 品番末尾

    ' 工程選択に応じて品番コンボボックスを初期化
    If ComboBox1.Value <> "" Then
        ComboBox1_Change
    End If

    ' TextBox1（日付）にフォーカスを設定
    TextBox1.SetFocus

    ' IMEモードを強制的に半角英数に設定
    SetAlphaIME TextBox1
End Sub
