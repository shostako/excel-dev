# 文字エンコーディング管理

## 参考マクロの読み込み（超重要！）

**必須事項**: inboxフォルダの参考マクロファイルは**必ず文字コード変換してから読むこと**

```powershell
# 読む前に必ず実行（PowerShell）
Get-Content -Path "inbox/ファイル名.bas" -Encoding oem | Select-Object -First 100

# または Git Bash使用
iconv -f SHIFT-JIS -t UTF-8 "inbox/ファイル名.bas" | head -100
```

**理由**:
- ユーザーがExcelからエクスポートしたファイルは**Shift-JIS**
- そのまま読むと文字化けして**列名を誤認識**する
- 過去の事故例：文字化けした列名で修正→本番で動作しない

**禁止事項**:
- 文字化けしたまま読み進めること
- 文字化けした内容を基に修正すること
- UTF-8版があってもオリジナルの確認を怠ること

---

## エンコーディング体系

| ディレクトリ | エンコーディング | 用途 |
|-------------|-----------------|------|
| inbox/ | Shift-JIS | Excelエクスポート（参考マクロ） |
| src/ | UTF-8 | Claude編集用 |
| macros/ | Shift-JIS | Excel取り込み用 |

---

## bas2sjisスクリプト

### 用途と仕組み

- **目的**: UTF-8のbasファイルをShift-JISに変換
- **出力先**: `macros/`ディレクトリ
- **CRLF対応**: 自動的にCRLF→LF変換を実行

### 使用方法

```powershell
# PowerShell版
.\scripts\bas2sjis.ps1 src\マクロ名.bas

# 変換結果の確認
Get-ChildItem macros\
```

### 変換プロセス

1. CRLF→LF変換（一時ファイル使用）
2. UTF-8→Shift-JIS変換
3. `macros/`ディレクトリに出力

---

## 読み込み時の確認手順

```powershell
# 参考マクロファイル読み込み前に必ず実行
Get-Content -Path "inbox/ファイル名.bas" -Encoding oem | Select-Object -First 100
```

---

## トラブルシューティング

### 文字化け問題

- **原因**: エンコーディング不一致
- **対策**: 適切なエンコーディング指定で読み込み
- **確認**: ファイルの先頭部分を確認

### 変換失敗

- **原因**: CRLF行末、特殊文字
- **対策**: 変換オプション使用
- **確認**: 変換結果ファイルの内容確認
