# excel-auto 進捗管理

**最終更新**: 2026-01-10 13:18
**直近Gitコミット**: 84c0cca feat: 手直し転記マクロ改修（加工T/成形T/塗装T）

---

## 現在の状態

月別12ファイル分割運用を継続。年間1ファイル運用はクエリ速度の問題で却下。

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
- [x] パワークエリ年間ファイル対応（2026-01-08）→ 却下
  - 年間1ファイルはクエリ速度が遅いため月別運用継続
  - MissingField.UseNullの知見は有効（スキーマキャッシュ問題対策）
- [x] 月別分割マクロ進捗表示復元（2026-01-08）
  - 12ファイル処理に戻し、進捗表示を「○年 X/12」形式に
- [x] ゾーン別転送マクロ重複確認機能削除（2026-01-08）
  - 無条件INSERTに変更
- [x] /vbaコマンド作成（2026-01-08）
  - ~/.claude/commands/vba.md を新規作成
  - VBAマクロ生成時にexcel-vba-expertスキルを確実に使わせる仕組み
  - 処理概要→選択肢（生成開始/追加情報）→スキル呼び出し→生成のフロー
- [x] セッションワークフロールール追加（2026-01-08）
  - ~/.claude/rules/01-session-workflow.md を新規作成
  - 進捗確認時はcwdのPROGRESS.mdを読むルール
- [x] 廃棄転記マクロ改修（2026-01-10）
  - m転記_廃棄_加工H.bas, m転記_廃棄_成形H.bas, m転記_廃棄_塗装H.bas
  - ソート統一（BubbleSort→QuickSort）
  - 空白期間対応（データなしでもテーブル構造を出力）
  - 期間テーブル空白行スキップ（CDate(0)=12/30問題修正）

- [x] 手直し転記マクロ改修（2026-01-10）
  - m転記_手直し_加工T.bas, m転記_手直し_成形T.bas, m転記_手直し_塗装T.bas
  - ソート統一（成形T: BubbleSort→QuickSort）
  - 空白期間対応（データなしでもテーブル構造を出力）
  - 期間テーブル空白行スキップ（CDate(0)=12/30問題修正）

## 次セッションへの引き継ぎ

- 廃棄マクロ3つのコメント充実（成形Hに倣って加工H/塗装Hも詳細コメント追加）
  - Windows環境のEdit/Write失敗で中断
