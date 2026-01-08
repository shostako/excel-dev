let
    ソース = Access.Database(File.Contents("Z:\全社共有\オート事業部\日報\不良集計\不良集計表\2026年\不良調査表DB-2026.accdb"), [CreateNavigationProperties=true]),
    テーブル = ソース{[Schema="",Item="_不良集計ゾーン別"]}[Data],
    // MissingField.UseNull: 列が見つからない場合はnullで補完（スキーマキャッシュ問題回避）
    削除された他の列 = Table.SelectColumns(テーブル,{"ID", "日付", "品番", "品番末尾", "注番月", "ロット", "発見", "ゾーン", "番号", "数量", "差戻し"}, MissingField.UseNull),
    // 差戻し列のnullを0に置換
    差戻し補完 = Table.ReplaceValue(削除された他の列, null, 0, Replacer.ReplaceValue, {"差戻し"}),
    変更された型 = Table.TransformColumnTypes(差戻し補完,{{"数量", Int64.Type}, {"差戻し", Int64.Type}})
in
    変更された型
