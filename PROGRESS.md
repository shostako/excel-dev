# excel-auto 進捗管理

**最終更新**: 2026-01-05 17:30
**直近Gitコミット**: 6a5fe73 feat: 月別分割マクロを年別対応に改修

---

## 現在の状態

データ転送マクロの年別自動振り分け機能を実装完了。2025年/2026年の年跨ぎデータにも対応。

## アクティブタスク

**rulesディレクトリ構造改善**（次セッションで実施）

docs/excel-knowledge/の必読ファイル内容をrules/に統合する。現状の「別ファイルを読め」という指示は仕組みとして機能しないため。

### ファイル場所

- ソース: `src/` (UTF-8)
- Excel用: `macros/` (Shift-JIS)
- バックアップ: `inbox/backup/`

---

## 完了済み

- [x] Windows環境用プロジェクトセットアップ
- [x] ルール設定（01-04共通 + 10-13 VBA固有）
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

## 次セッションへの引き継ぎ

**必須タスク**: docs/excel-knowledge/の必読ファイル内容をrules/に統合

対象ファイル:
1. `docs/excel-knowledge/failures/001_activate_vs_screenupdating.md`
2. `docs/excel-knowledge/patterns/VBA_OPTIMIZATION_PATTERNS.md`
3. `docs/excel-knowledge/techniques/VBA_BASIC_TECHNIQUES.md`
4. `docs/excel-knowledge/claude-code/EXCEL_MACRO_KNOWLEDGE_BASE.md`

理由: rulesファイルは自動読み込みされるが、rulesから参照された別ファイルは自動読み込みされない。「別ファイルを読め」という指示は仕組みとして機能しない。
