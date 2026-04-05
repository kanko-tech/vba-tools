# クラスモジュール `Matrix`

`Matrix` は、2 次元配列を扱うための VBA クラスモジュールです。  
表データを 2 次元配列として保持し、参照・抽出・更新・出力を行うための基本クラスです。

## 役割

- Excel の矩形 `Range` を 2 次元配列として読み込む
- 2 次元配列をそのまま保持する
- 行数・列数・セル値を参照する
- 行フィルタや列選択を行う
- 列単位の値更新を行う
- ワークシートへまとめて出力する

## 基本的な使い方

```vb
Sub Sample_Matrix()
    Dim mat As New Matrix
    Dim filtered As Matrix
    Dim mask(1 To 3) As Boolean

    mat.read_range Sheet1.Range("A1:C3")

    mask(1) = True
    mask(2) = False
    mask(3) = True

    Set filtered = mat.filter_rows(mask)
    filtered.to_range Sheet1.Range("E1")
End Sub
```

## 主なメソッド

### `read_range(ByVal rng As Range)`

矩形 `Range` を 2 次元配列として読み込みます。

### `read_array(ByVal arr As Variant)`

2 次元配列を読み込みます。  
1 次元配列や 3 次元以上の配列は受け付けません。

### `row_count As Long`

行数を返します。

### `col_count As Long`

列数を返します。

### `item(ByVal rowIndex As Long, ByVal colIndex As Long) As Variant`

指定位置の値を返します。

### `row_values(ByVal rowIndex As Long) As Variant`

指定行を一次元配列として返します。

### `col_values(ByVal colIndex As Long) As Variant`

指定列を一次元配列として返します。

### `filter_rows(ByVal mask As Variant) As Matrix`

ブール配列を使って行を絞り込みます。

### `select_columns(ByVal columnIndexes As Variant) As Matrix`

列番号配列を使って列を選択します。

### `set_column_values(ByVal colIndex As Long, ByVal values As Variant)`

指定列を一次元配列で置き換えます。

### `to_range(ByVal topLeft As Range)`

内容をワークシートにまとめて出力します。

## 注意点

- `Matrix` は列名を持ちません
- 列名付きのテーブル操作は `Table` クラスで扱う想定です
- `filter_rows` の `mask` 長は行数と一致している必要があります
- `set_column_values` の `values` 長は行数と一致している必要があります
- 現段階では `filter_rows` で 0 件になった場合は空の `Matrix` ではなくエラーにしています
