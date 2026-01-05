# excel-auto 進捗管理

**最終更新**: 2026-01-05 21:30
**直近Gitコミット**: (pending) feat: UserForm日付入力yyyy/m/d対応

---

## 現在の状態

UserFormの日付入力を年末年始対応に改修。m/d入力時にシステム日付の西暦を使用し、表示はyyyy/m/d形式で確認可能。明示的にyyyy/m/dで入力すればその西暦を使用。

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

## 次セッションへの引き継ぎ

特になし。
