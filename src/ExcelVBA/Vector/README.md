# クラスモジュール `Vector`

`Vector` は、一次元配列を扱うための VBA クラスモジュールです。  
1 列または 1 行のデータを読み込み、条件マスク生成、型変換、集計、セル出力を行います。

## この README の役割

この README は入口に絞っています。  
メソッド早見表、補足、レシピ、全メソッド解説は GitHub Pages 側に移しています。

## 詳細ドキュメント

- Pages ガイド: [Vector Guide](https://kanko-tech.github.io/VBA_tools/vector.html)
- ソース配置: `src/ExcelVBA/Vector`

## まず何ができるか

- 1 列または 1 行の `Range` を一次元配列として保持する
- `eq` `gt` `is_empty` などで条件マスクを作る
- `cast_to_double_safe` `fill_empty` などで列を整形する
- `sum` `mean` `unique` で列を集計する

## クイックスタート

```vb
Sub Sample_Vector_QuickStart()
    Dim vec As New Vector
    Dim mask As Variant

    vec.read_col_range Sheet1.Range("B2:B6")
    vec.fill_empty 0
    vec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True

    mask = vec.gt(100)

    Debug.Print vec.sum
    Debug.Print vec.mean
End Sub
```
