Attribute VB_Name = "mTransferYears"
Option Explicit

' ========================================
' モジュール名: mTransferYears
' 処理概要: 転送処理で処理した年のリストを管理
' 用途: 転送マクロと月別分割マクロ間で処理対象年を共有
' 作成日: 2025-01-05
'
' 使用方法:
' 1. 転送マクロの最初で InitTransferredYears を呼ぶ
' 2. 各年の転送成功時に AddTransferredYear を呼ぶ
' 3. 月別分割マクロで GetTransferredYears を呼んで年リストを取得
' ========================================

' 転送処理で処理した年のリスト（Dictionaryで重複排除）
Public g_TransferredYears As Object

' ============================================
' 初期化: 年リストをクリアして新規作成
' ============================================
Public Sub InitTransferredYears()
    Set g_TransferredYears = CreateObject("Scripting.Dictionary")
End Sub

' ============================================
' 年を追加: 処理した年をリストに追加
' 引数: yearValue - 年（Integer）
' ============================================
Public Sub AddTransferredYear(yearValue As Integer)
    If g_TransferredYears Is Nothing Then
        InitTransferredYears
    End If
    If Not g_TransferredYears.Exists(yearValue) Then
        g_TransferredYears.Add yearValue, True
    End If
End Sub

' ============================================
' クリア: 年リストを空にする
' ============================================
Public Sub ClearTransferredYears()
    If Not g_TransferredYears Is Nothing Then
        g_TransferredYears.RemoveAll
    End If
End Sub

' ============================================
' 年リストを取得: 処理した年の配列を返す
' 戻り値: 年の配列（Variant）、空の場合は空配列
' ============================================
Public Function GetTransferredYears() As Variant
    If g_TransferredYears Is Nothing Then
        GetTransferredYears = Array()
    ElseIf g_TransferredYears.Count = 0 Then
        GetTransferredYears = Array()
    Else
        GetTransferredYears = g_TransferredYears.Keys
    End If
End Function

' ============================================
' 年リストの件数を取得
' 戻り値: 年の件数（Long）
' ============================================
Public Function GetTransferredYearsCount() As Long
    If g_TransferredYears Is Nothing Then
        GetTransferredYearsCount = 0
    Else
        GetTransferredYearsCount = g_TransferredYears.Count
    End If
End Function
