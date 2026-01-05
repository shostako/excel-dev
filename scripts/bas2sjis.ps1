<#
.SYNOPSIS
    UTF-8のbasファイルをShift-JISに変換してmacros/に出力

.DESCRIPTION
    src/ディレクトリのUTF-8エンコードのVBAファイルを
    Excel VBEでインポート可能なShift-JIS形式に変換する

.PARAMETER InputFile
    変換元のbasファイルパス

.EXAMPLE
    .\scripts\bas2sjis.ps1 src\m期間集計_通称別a.bas
#>

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$InputFile
)

$ErrorActionPreference = "Stop"

# 入力ファイルの存在確認
if (-not (Test-Path $InputFile)) {
    Write-Error "エラー: ファイル '$InputFile' が見つかりません"
    exit 1
}

# ファイル名抽出（拡張子を維持）
$FileName = [System.IO.Path]::GetFileName($InputFile)
$OutputFile = "macros\$FileName"

# macros ディレクトリの存在確認・作成
if (-not (Test-Path "macros")) {
    New-Item -ItemType Directory -Path "macros" | Out-Null
    Write-Host "macros ディレクトリを作成しました"
}

try {
    # UTF-8でファイル読み込み（BOMあり/なし両対応）
    $content = Get-Content -Path $InputFile -Encoding UTF8 -Raw

    # CRLF→LF変換してからCRLFに統一（Windows/Excel用）
    $content = $content -replace "`r`n", "`n"
    $content = $content -replace "`n", "`r`n"

    # Shift-JIS（CP932）で出力
    $sjisEncoding = [System.Text.Encoding]::GetEncoding("shift_jis")
    [System.IO.File]::WriteAllText($OutputFile, $content, $sjisEncoding)

    Write-Host "変換完了: $InputFile → $OutputFile"
}
catch {
    Write-Error "エラー: 変換に失敗しました - $($_.Exception.Message)"
    exit 1
}
