Attribute VB_Name = "mフィルター解除"
Option Explicit

' ========================================
' マクロ名: フィルター解除
' 処理概要: 全テーブルのフィルターを解除（全行を表示）
' 対象テーブル: _完成品, _core, _slitter, _acf
' 解除対象:
'   - 行の非表示（EntireRow.Hidden）
'   - オートフィルター（AutoFilter）
' ========================================

Sub フィルター解除()
    ' --------------------------------------------
    ' 元の設定を保存
    ' --------------------------------------------
    Dim origScreenUpdating As Boolean
    Dim origEnableEvents As Boolean
    origScreenUpdating = Application.ScreenUpdating
    origEnableEvents = Application.EnableEvents

    ' --------------------------------------------
    ' 画面更新・イベント抑制（高速化）
    ' --------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject

    ' 対象テーブル名
    Dim tables As Variant
    tables = Array("_完成品", "_core", "_slitter", "_acf")

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' 全テーブルのフィルターを解除
    ' --------------------------------------------
    Dim tblName As Variant
    For Each tblName In tables
        Set tbl = FindTableByPattern(ws, CStr(tblName))

        ' テーブルが見つからない場合はスキップ
        If tbl Is Nothing Then GoTo NextTable

        ' データ行がある場合のみ処理
        If Not tbl.DataBodyRange Is Nothing Then
            ' 行の非表示を解除
            tbl.DataBodyRange.EntireRow.Hidden = False
        End If

        ' オートフィルターをリセット（フィルター適用中の場合のみ）
        If Not tbl.AutoFilter Is Nothing Then
            If tbl.AutoFilter.FilterMode Then
                tbl.AutoFilter.ShowAllData
            End If
        End If

NextTable:
    Next tblName

    ' --------------------------------------------
    ' フィルター参照セルをリセット
    ' --------------------------------------------
    ws.Range("B3").Value = ""
    ws.Range("C3").Value = ""
    ws.Range("D3").Value = ""
    ws.Range("E3").Value = "全項目"

    ' --------------------------------------------
    ' 垂直スクロールのみ先頭に移動（水平位置は維持）
    ' --------------------------------------------
    ActiveWindow.ScrollRow = 1

    GoTo Cleanup

ErrorHandler:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical

Cleanup:
    Application.ScreenUpdating = origScreenUpdating
    Application.EnableEvents = origEnableEvents
End Sub
