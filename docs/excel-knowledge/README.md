# Excel Knowledge Base

ExcelマクロとM言語（Power Query）の参考資料。

**注**: 核心的なルール（Activate禁止、コメント標準等）は `.claude/rules/11-vba-coding-standards.md` に統合済み。このフォルダは詳細な参考資料として残している。

## 構成

### failures/ - 失敗事例集
実際に起きたミスとその教訓。「なぜそうなったか」を含めて記録。

| ファイル | 内容 |
|----------|------|
| 005_optimization_restore_calculation.md | 最適化設定の保存・復元による計算式更新漏れ |
| 006_vba_line_continuation_limit.md | VBA行継続文字の24個制限 |
| 007_pivot_filter_clearallfilters_failure.md | ピボットテーブルのClearAllFiltersが効かない問題 |
| 008_listobject_dynamic_table_count.md | ListObjectとRangeの混同による削除漏れ |

### m-language/ - M言語専用
Power Query (M言語)の基礎と落とし穴。

### examples/ - 実装例集
コピペして使えるVBAコード集。

## 基本方針

1. **失敗を隠さない** - ミスこそが最高の教材
2. **理由を明記** - 「なぜ」がわからなければ意味がない
3. **実例重視** - 抽象論より具体的なコード

---

*最終更新: 2026-01-05（冗長ファイル削除、rulesに統合）*