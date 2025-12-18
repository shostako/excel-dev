Attribute VB_Name = "m完成品フィルター"
Option Explicit

' ========================================
' マクロ名: 完成品フィルター
' 処理概要: B3セルの値でオートフィルターを適用
' 参照セル: B3（フィルター条件）
' フィルター対象:
'   - _完成品テーブル「製品名」列 → B3そのまま
'   - _core, _slitter, _acfテーブル「小部品」列 → B3末尾4字除去
' 条件分岐:
'   - 「全品番」→ フィルター解除
'   - それ以外 → 指定値 + 「稼働日」「合計」でフィルター
' 複合フィルター: 他のフィルターは維持（オートフィルター方式）
' ========================================

Sub 完成品フィルター()
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
    Dim filterValue As String
    Dim filterValueTrimmed As String
    Dim colIndex As Long
    Dim filterArray(0 To 2) As String
    Dim filterArrayTrimmed(0 To 2) As String

    ' 小部品テーブル名
    Dim subTables As Variant
    subTables = Array("_core", "_slitter", "_acf")

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' フィルター条件の取得
    ' --------------------------------------------
    filterValue = ws.Range("B3").Value

    ' 末尾4字除去（小部品テーブル用）
    If Len(filterValue) > 4 Then
        filterValueTrimmed = Left(filterValue, Len(filterValue) - 4)
    Else
        filterValueTrimmed = filterValue
    End If

    ' フィルター条件配列（「稼働日」「合計」を追加）
    filterArray(0) = filterValue
    filterArray(1) = "稼働日"
    filterArray(2) = "合計"

    filterArrayTrimmed(0) = filterValueTrimmed
    filterArrayTrimmed(1) = "稼働日"
    filterArrayTrimmed(2) = "合計"

    ' --------------------------------------------
    ' _完成品テーブル：製品名列にオートフィルター
    ' --------------------------------------------
    Set tbl = FindTableByPattern(ws, "_完成品")
    If tbl Is Nothing Then
        Err.Raise vbObjectError + 1, , "テーブル '_完成品' が見つかりません"
    End If

    colIndex = tbl.ListColumns("製品名").Index

    If filterValue = "全品番" Or filterValue = "" Then
        ' フィルター解除
        If Not tbl.AutoFilter Is Nothing Then
            If tbl.AutoFilter.FilterMode Then
                tbl.Range.AutoFilter Field:=colIndex
            End If
        End If
    Else
        ' フィルター適用（「稼働日」「合計」を含む）
        tbl.Range.AutoFilter Field:=colIndex, _
            Criteria1:=filterArray, _
            Operator:=xlFilterValues
    End If

    ' --------------------------------------------
    ' 小部品テーブル：小部品列にオートフィルター（末尾4字除去）
    ' --------------------------------------------
    Dim tblName As Variant
    For Each tblName In subTables
        Set tbl = FindTableByPattern(ws, CStr(tblName))
        If tbl Is Nothing Then
            Debug.Print "警告: テーブル '" & tblName & "' が見つかりません"
            GoTo NextTable
        End If

        colIndex = tbl.ListColumns("小部品").Index

        If filterValue = "全品番" Or filterValue = "" Then
            ' フィルター解除
            If Not tbl.AutoFilter Is Nothing Then
                If tbl.AutoFilter.FilterMode Then
                    tbl.Range.AutoFilter Field:=colIndex
                End If
            End If
        Else
            ' フィルター適用（末尾4字除去した値 + 「稼働日」「合計」）
            tbl.Range.AutoFilter Field:=colIndex, _
                Criteria1:=filterArrayTrimmed, _
                Operator:=xlFilterValues
        End If
NextTable:
    Next tblName

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
