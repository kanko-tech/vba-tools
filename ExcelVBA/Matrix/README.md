# クラスモジュール `Matrix`

`Matrix` は、2 次元配列を扱うための基本クラスです。  
列名を持たない表データを保持し、行抽出・列選択・更新・出力を行います。

## 概要

| 項目 | 内容 |
| --- | --- |
| 主な用途 | 2 次元配列や矩形 `Range` を中間表として扱う |
| 内部データ | `Variant` の 2 次元配列 |
| 添字 | 正規化後は 1 始まり |
| 空表 | `row_count=0` でも列数は保持可能 |
| 他モジュールとの関係 | `Table` の内部表現として利用 |

## 利用前提

| 項目 | 内容 |
| --- | --- |
| 初期化必須 | `read_range` / `read_array` / `read_empty` のいずれかを先に実行 |
| データ形状 | `read_array` は 2 次元配列専用 |
| 未初期化時 | 参照系・操作系・出力系はエラー |
| 出力制約 | `to_range` は行数 0 の `Matrix` では実行不可 |

## 典型的なユースケース

- Excel の表を一括で読み込んで中間データとして保持する
- 行マスクで表を絞り込む
- 必要な列だけを抜き出して次工程に渡す
- `Table` へ渡す前の下処理を行う

## メソッド一覧

| 区分 | メソッド |
| --- | --- |
| 読込 | `read_range` `read_array` `read_empty` |
| 参照 | `is_loaded` `row_count` `col_count` `data` `item` `row_values` `col_values` |
| 行列操作 | `filter_rows` `slice_rows` `select_columns` `set_column_values` `append_row` `append_column` `transpose` |
| 出力 | `to_range` `clear` |

## 読込系

### `read_range(ByVal rng As Range)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 矩形 `Range` を 2 次元配列として読み込む |
| 前提条件 | `rng` が `Nothing` でないこと |
| 入力 | `rng As Range` |
| 実行内容 | 単一セルは `1 x 1`、複数セルは矩形配列として正規化して保持 |
| 出力 | 内部状態 `MATRIX` を更新 |
| ユースケース | シート表の一括読込 |

ユースケース例: シート上の表を読み込み、あとで条件抽出に使う。

```vb
Dim mat As New Matrix
mat.read_range Sheet1.Range("A1:C10")
```

### `read_array(ByVal arr As Variant)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 2 次元配列を読み込む |
| 前提条件 | `arr` が配列であり、2 次元配列であること |
| 入力 | `arr As Variant` |
| 実行内容 | 内部データを 1 始まりの 2 次元配列へ正規化 |
| 出力 | 内部状態 `MATRIX` を更新 |
| 注意点 | 1 次元や 3 次元以上はエラー |
| ユースケース | 外部生成された表データの受入れ |

### `read_empty(ByVal colCount As Long)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 列数だけを持つ空 `Matrix` を作る |
| 前提条件 | `colCount >= 1` |
| 入力 | `colCount As Long` |
| 実行内容 | 行数 0、列数 `colCount` の空表を生成 |
| 出力 | 内部状態 `MATRIX` を更新 |
| ユースケース | フィルタ 0 件結果、空 `Table` の土台 |

## 参照系

### `is_loaded() As Boolean`

| 項目 | 内容 |
| --- | --- |
| 役割 | 読込済みかを返す |
| 戻り値 | `Boolean` |
| ユースケース | 実行前ガード、デバッグ確認 |

### `row_count() As Long`

| 項目 | 内容 |
| --- | --- |
| 役割 | 行数を返す |
| 戻り値 | `Long` |
| ユースケース | 行ループ、結果件数確認 |

### `col_count() As Long`

| 項目 | 内容 |
| --- | --- |
| 役割 | 列数を返す |
| 戻り値 | `Long` |
| ユースケース | 構造確認、列ループ |

### `data() As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 内部 2 次元配列のコピーを返す |
| 前提条件 | 読込済みであること |
| 戻り値 | `Variant` の 2 次元配列 |
| 注意点 | 返却後の変更は内部状態へ直接反映されない |
| ユースケース | 他関数への受渡し、デバッグ確認 |

### `item(ByVal rowIndex As Long, ByVal colIndex As Long) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定セルの値を返す |
| 前提条件 | 読込済みであること、`rowIndex` と `colIndex` が範囲内であること |
| 入力 | `rowIndex As Long` `colIndex As Long` |
| 戻り値 | セル位置の `Variant` 値 |
| ユースケース | 単一点の確認、条件分岐 |

### `row_values(ByVal rowIndex As Long) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定行を一次元配列で返す |
| 前提条件 | 読込済みであること、`rowIndex` が範囲内であること |
| 入力 | `rowIndex As Long` |
| 戻り値 | `Variant` の一次元配列 |
| ユースケース | 特定行の抜出し、行単位処理 |

### `col_values(ByVal colIndex As Long) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定列を一次元配列で返す |
| 前提条件 | 読込済みであること、`colIndex` が範囲内であること、行が 1 件以上あること |
| 入力 | `colIndex As Long` |
| 戻り値 | `Variant` の一次元配列 |
| 注意点 | 行数 0 の `Matrix` ではエラー |
| ユースケース | 列条件生成、列更新データ作成 |

## 行列操作系

### `filter_rows(ByVal mask As Variant) As Matrix`

| 項目 | 内容 |
| --- | --- |
| 役割 | ブール配列で行を絞り込む |
| 前提条件 | 読込済みであること、`mask` が一次元配列であり行数と一致すること |
| 入力 | `mask As Variant` |
| 戻り値 | 新しい `Matrix` |
| 注意点 | 0 件の場合は空 `Matrix` を返す |
| ユースケース | 条件抽出、中間結果の作成 |

ユースケース例: 条件に合う行だけ残した表を作る。

```vb
Dim mat As New Matrix
Dim filtered As Matrix
Dim mask(1 To 3) As Boolean

mat.read_range Sheet1.Range("A1:C3")
mask(1) = True
mask(2) = False
mask(3) = True

Set filtered = mat.filter_rows(mask)
```

### `slice_rows(ByVal rowIndexes As Variant) As Matrix`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定行番号だけを順番通りに取り出す |
| 前提条件 | 読込済みであること、`rowIndexes` が一次元配列であること、各行番号が範囲内であること |
| 入力 | `rowIndexes As Variant` |
| 戻り値 | 新しい `Matrix` |
| 注意点 | 空配列相当なら空 `Matrix` を返す |
| ユースケース | 任意行抜出し、再配列 |

### `select_columns(ByVal columnIndexes As Variant) As Matrix`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定列だけを返す |
| 前提条件 | 読込済みであること、`columnIndexes` が一次元配列であること、各列番号が範囲内であること |
| 入力 | `columnIndexes As Variant` |
| 戻り値 | 新しい `Matrix` |
| 注意点 | 列指定 0 件はエラー |
| ユースケース | 必要列だけの中間表作成 |

ユースケース例: 1 列目と 3 列目だけを抜き出す。

```vb
Dim mat As New Matrix
Dim picked As Matrix

mat.read_range Sheet1.Range("A1:C10")
Set picked = mat.select_columns(Array(1, 3))
```

### `set_column_values(ByVal colIndex As Long, ByVal values As Variant)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定列全体を置き換える |
| 前提条件 | 読込済みであること、`colIndex` が範囲内であること、`values` が一次元配列で行数と一致すること |
| 入力 | `colIndex As Long` `values As Variant` |
| 実行内容 | 対象列を新しい値で上書き |
| 出力 | 内部表を更新 |
| ユースケース | 計算済み列の反映 |

### `append_row(ByVal values As Variant)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 末尾に 1 行追加する |
| 前提条件 | 読込済みであること、`values` が一次元配列で列数と一致すること |
| 入力 | `values As Variant` |
| 実行内容 | 既存表の末尾に新規行を追加 |
| 出力 | 内部表を更新 |
| ユースケース | 集計結果行や追加入力行の付加 |

### `append_column(ByVal values As Variant)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 末尾に 1 列追加する |
| 前提条件 | 読込済みであること |
| 入力 | `values As Variant` |
| 実行内容 | 行がある場合は値配列を追加し、行が 0 件なら空列構造だけを追加 |
| 出力 | 内部表を更新 |
| 注意点 | 行がある場合、`values` は一次元配列で行数と一致する必要がある |
| ユースケース | 計算列、フラグ列の追加 |

ユースケース例: 元の表にフラグ列を追加する。

```vb
Dim mat As New Matrix

mat.read_range Sheet1.Range("A1:C5")
mat.append_column Array("Y", "N", "Y", "N", "Y")
```

### `transpose() As Matrix`

| 項目 | 内容 |
| --- | --- |
| 役割 | 転置した新しい `Matrix` を返す |
| 前提条件 | 読込済みであること、行数 0 でないこと |
| 戻り値 | 新しい `Matrix` |
| 注意点 | 空 `Matrix` の転置は未サポート |
| ユースケース | 行列方向の入替え、見せ方の変更 |

## 出力系

### `to_range(ByVal topLeft As Range)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 表内容をシートへ出力する |
| 前提条件 | 読込済みであること、`topLeft` が `Nothing` でないこと、行数 0 でないこと |
| 入力 | `topLeft As Range` |
| 実行内容 | `topLeft` を左上として矩形範囲へ書込み |
| 出力 | ワークシート上のセル範囲 |
| ユースケース | 中間表や最終結果の確認 |

ユースケース例: 抽出後の表を別領域に表示する。

```vb
Dim mat As New Matrix

mat.read_range Sheet1.Range("A1:C5")
mat.to_range Sheet1.Range("F1")
```

### `clear()`

| 項目 | 内容 |
| --- | --- |
| 役割 | 内部状態を初期化する |
| 実行内容 | 配列、行数、列数、読込状態を初期化 |
| ユースケース | 再利用前のリセット |

## 補足

- 空結果は `read_empty` ベースで表現します
- 行数 0 でも列数は保持されます
- `Table` はこの `Matrix` を内部表現として利用します
