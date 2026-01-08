Attribute VB_Name = "mRunAccess_月別分割"
Option Explicit

' ========================================
' マクロ名: Access月別分割マクロ実行
' 処理概要: ExcelからAccessデータベースの月別分割処理を外部実行
'          転送マクロで処理した年の各月DBに対して月別分割を実行
' 作成日: 不明
' 更新日: 2026-01-08（12ファイル処理・進捗表示を復元）
'
' 処理の流れ:
' 1. 転送マクロから処理済み年リストを取得
' 2. 各年の1〜12月DBファイルに対して月別分割マクロを実行
' 3. 進捗表示: 「○年 X/12」形式
' ========================================

' DBパス設定（転送マクロと同じ）
Const DB_BASE_PATH As String = "Z:\全社共有\オート事業部\日報\不良集計\不良集計表\"
Const DB_FILE_PREFIX As String = "不良調査表DB-"

' ============================================
' 補助関数: BuildDBPath
' 役割: 年と月からDBファイルパスを動的生成
' 形式: 不良調査表DB-2025-01.accdb
' ============================================
Function BuildDBPath(yearValue As Integer, monthValue As Integer) As String
    Dim monthStr As String
    monthStr = Format(monthValue, "00")
    BuildDBPath = DB_BASE_PATH & yearValue & "年\" & DB_FILE_PREFIX & yearValue & "-" & monthStr & ".accdb"
End Function

Sub RunAccess_月別分割()
    ' ============================================
    ' 変数宣言
    ' ============================================
    Dim acc As Object
    Dim dbPath As String
    Dim years As Variant
    Dim yearValue As Variant
    Dim monthNum As Integer

    ' ============================================
    ' 処理対象の年を取得
    ' ============================================
    years = GetTransferredYears()

    ' 年リストが空の場合は終了
    If IsArrayEmpty(years) Then
        Application.StatusBar = "月別分割: 処理対象の年がありません"
        Application.Wait Now + TimeValue("00:00:02")
        Application.StatusBar = False
        Exit Sub
    End If

    ' ============================================
    ' 初期設定
    ' ============================================
    Application.ScreenUpdating = False
    Application.EnableCancelKey = 0   ' xlDisable（Ctrl+Break無効化）

    On Error GoTo EH

    ' ============================================
    ' 各年・各月に対して月別分割を実行
    ' ============================================
    For Each yearValue In years
        For monthNum = 1 To 12
            dbPath = BuildDBPath(CInt(yearValue), monthNum)

            ' 進捗表示: 「2025年 1/12」形式
            Application.StatusBar = "月別分割: " & yearValue & "年 " & monthNum & "/12"

            ' Access起動とDB接続
            Set acc = CreateObject("Access.Application")
            acc.Visible = False
            acc.OpenCurrentDatabase dbPath, False

            ' マクロ実行（関数→UIマクロの2段階フォールバック）
            On Error Resume Next
            acc.Run "月別分割_Run"
            If Err.Number <> 0 Then
                Err.Clear
                acc.DoCmd.RunMacro "月別分割"
            End If
            On Error GoTo EH

            ' Access終了
            acc.CloseCurrentDatabase
            acc.Quit
            Set acc = Nothing

            DoEvents
        Next monthNum
    Next yearValue

    ' ============================================
    ' 完了処理
    ' ============================================
    Application.StatusBar = "月別分割: 完了"
    Application.Wait Now + TimeValue("00:00:02")

CleanUp:
    On Error Resume Next
    If Not acc Is Nothing Then
        acc.CloseCurrentDatabase
        acc.Quit
        Set acc = Nothing
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableCancelKey = 1   ' xlInterrupt（通常状態に復帰）
    Exit Sub

EH:
    Application.StatusBar = "月別分割: 失敗 (" & Err.Number & ") " & Err.Description
    Application.Wait Now + TimeValue("00:00:03")
    Resume CleanUp
End Sub

' ============================================
' 補助関数: IsArrayEmpty
' 役割: 配列が空かどうかを判定
' ============================================
Private Function IsArrayEmpty(arr As Variant) As Boolean
    On Error Resume Next

    ' 配列でない場合
    If Not IsArray(arr) Then
        IsArrayEmpty = True
        Exit Function
    End If

    ' 空配列の場合、UBound < LBound になる
    If UBound(arr) < LBound(arr) Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If

    On Error GoTo 0
End Function
