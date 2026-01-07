# excel-auto 進捗管理

**最終更新**: 2026-01-07 11:48
**直近Gitコミット**: f08e925 feat: UserForm日付入力yyyy/m/d対応（年末年始対策）

---

## 現在の状態

UserFormのテーブル名・列名参照エラーを修正。ルールに「参照名の確認（推測禁止）」を追加。

### ファイル場所

- ソース: `src/` (UTF-8)
- Excel用: `macros/` (Shift-JIS)
- バックアップ: `inbox/backup/`

---

## 完了済み

- [x] Windows環境用プロジェクトセットアップ
- [x] ルール設定（01-04共通 + 11-13 VBA固有）
- [x] scripts/bas2sjis.ps1 動作確認
- [x] GitHub連携（https://github.com/shostako/excel-dev）
- [x] docs/excel-knowledge/ ナレッジ整備
- [x] m見出し改訂.bas 開発完了（印刷設定・日付書式含む）
- [x] データ転送マクロ年別自動振り分け機能（2026-01-05）
  - mロット数量データ転送ADO.bas
  - mゾーン別データ転送ADO.bas
- [x] 月別分割マクロ年別対応（2026-01-05）
  - mTransferYears.bas（新規）
  - mRunAccess_月別分割.bas（改修）
- [x] rulesディレクトリ最適化（2026-01-05）
  - 10-vba-required-reading.md 削除（機能しない参照指示）
  - 11-vba-coding-standards.md 拡充（失敗パターン、コメント標準、黄金律）
- [x] docs/excel-knowledge/ 整理（2026-01-05）
  - 冗長ファイル削除（rulesに統合済みの5ファイル）
  - フォルダ構造簡素化（failures/, m-language/, examples/）
- [x] UserForm日付入力yyyy/m/d対応（2026-01-05）
  - CTextBoxEvent.cls, UserForm1.frm, UserForm2.frm改修
  - bas2sjis.ps1を.cls/.frm対応に改修
- [x] UserFormテーブル名・列名参照エラー修正（2026-01-07）
  - UserForm1.frm: _ロット数量全 → _ロット数量S
  - UserForm2.frm: _不良集計ゾーン別全 → _不良集計ゾーン別S、工程→発見、戻めし→差戻し
  - 12-vba-encoding.md: 「参照名の確認（推測禁止）」ルール追加

## 次セッションへの引き継ぎ

特になし。
