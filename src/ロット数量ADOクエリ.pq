let
    ソース = Access.Database(File.Contents("Z:\全社共有\オート事業部\日報\不良集計\不良集計表\2026年\不良調査表DB-2026.accdb"), [CreateNavigationProperties=true]),
    テーブル = ソース{[Schema="",Item="_ロット数量"]}[Data],
    フィルターされた行 = Table.SelectRows(テーブル, each [日付] <> null and [日付] <> "")
in
    フィルターされた行
