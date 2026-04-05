# クラスモジュール `Table`

`Table` は、列名を持つ表データを扱うための VBA クラスモジュールです。  
`Matrix` を内部に持ち、`pandas.DataFrame` に近い列ベース操作を提供します。

## 役割

- ヘッダー付き `Range` を列名付きテーブルとして読み込む
- 列名で列を取得する
- 条件に一致する行だけを抽出する
- 条件に一致する行の特定列を書き換える
- 列名ベースで列選択する
- ヘッダー込みでシートへ出力する

## pandas との対応イメージ

- `read_range(..., hasHeader:=True)`: `DataFrame` の読込
- `col("status")`: `df["status"]`
- `filter_by_equals("status", "OK")`: `df[df["status"] == "OK"]`
- `set_by_mask mask, "score", 0`: `df.loc[mask, "score"] = 0`
- `select_columns(Array("name", "score"))`: `df[["name", "score"]]`

## 基本的な使い方

```vb
Sub Sample_Table()
    Dim tbl As New Table
    Dim okRows As Table
    Dim mask As Variant

    tbl.read_range Sheet1.Range("A1:C6"), hasHeader:=True

    Set okRows = tbl.filter_by_equals("status", "OK")
    okRows.to_range Sheet1.Range("E1")

    mask = tbl.col("status")
End Sub
```

## 主なメソッド

### `read_range(ByVal rng As Range, Optional ByVal hasHeader As Boolean = True)`

シート上の表を読み込みます。

- `hasHeader=True` の場合、先頭行を列名として扱います
- `hasHeader=False` の場合、列名は `col1`, `col2`, ... を自動生成します

### `read_matrix(ByVal src As Matrix, ByVal columnNames As Variant)`

既存の `Matrix` と列名配列から `Table` を構築します。

### `column_names() As Variant`

列名一覧を返します。

### `col(ByVal columnName As String) As Variant`

指定列を一次元配列で返します。  
`pandas` の `df["col"]` に近いメソッドです。

### `filter_by_mask(ByVal mask As Variant) As Table`

ブール配列で行を絞り込みます。

### `filter_by_equals(ByVal columnName As String, ByVal matchValue As Variant) As Table`

指定列が特定値に一致する行だけを返します。  
`pandas` の `df[df["col"] == value]` に相当します。

### `select_columns(ByVal columnNames As Variant) As Table`

列名配列で列を選択します。

### `set_by_mask(ByVal mask As Variant, ByVal columnName As String, ByVal newValue As Variant)`

条件に一致した行だけ、対象列の値を更新します。  
`pandas` の `loc` 的な更新の最小版です。

### `set_column(ByVal columnName As String, ByVal values As Variant)`

列全体を一次元配列で置き換えます。

### `to_range(ByVal topLeft As Range, Optional ByVal includeHeader As Boolean = True)`

表をヘッダー付きでシートへ出力します。

## 例: 条件に合う行だけ別列を書き換える

```vb
Sub Sample_Table_Update()
    Dim tbl As New Table
    Dim statusCol As Variant
    Dim mask() As Boolean
    Dim i As Long

    tbl.read_range Sheet1.Range("A1:C6"), hasHeader:=True

    statusCol = tbl.col("status")
    ReDim mask(LBound(statusCol) To UBound(statusCol))

    For i = LBound(statusCol) To UBound(statusCol)
        mask(i) = (statusCol(i) = "NG")
    Next i

    tbl.set_by_mask mask, "score", 0
    tbl.to_range Sheet1.Range("E1")
End Sub
```

## 注意点

- 現段階では複合条件、集計、結合、ソートまでは実装していません
- 列名は重複不可です
- `mask` の長さは行数と一致している必要があります
- 条件に一致する行が 0 件の場合、内部の `Matrix.filter_rows` に合わせてエラーになります
- 比較演算を列単位で簡潔に書けるようにするには、今後 `Vector` と連携した拡張が有効です
