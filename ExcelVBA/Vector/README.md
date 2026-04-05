# クラスモジュール `Vector`

`Vector` は、一次元配列を扱うための VBA クラスモジュールです。  
Excel の 1 列または 1 行の `Range`、または一次元配列を読み込み、型変換やセル出力をまとめて行えます。

## できること

- 1 列の `Range` を一次元配列として読み込む
- 1 行の `Range` を一次元配列として読み込む
- 一次元配列をそのまま保持する
- 配列の要素数や各要素を参照する
- 値を `Double`、`String`、`Date` に安全寄りに変換する
- ワークシートへ縦方向、横方向に出力する

## 前提

- このクラスが扱うのは一次元配列です
- 利用前に `read_col_range`、`read_row_range`、`read_array` のいずれかでデータを読み込んでください
- 未読込の状態で参照系・型変換系・出力系メソッドを呼ぶとエラーになります

## 基本的な使い方

```vb
Sub Sample_Vector()
    Dim vec As New Vector
    
    vec.read_col_range Sheet1.Range("A1:A5")
    
    Debug.Print vec.count
    Debug.Print vec.item(1)
    
    vec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True
    vec.to_range_vertical Sheet1.Range("C1")
End Sub
```

## 読み込み

### `read_col_range(ByVal rng As Range)`

1 列の `Range` を読み込みます。  
複数セルでも単一セルでも読み込めます。

- `rng Is Nothing` の場合はエラー
- 1 列でない場合はエラー

```vb
vec.read_col_range Sheet1.Range("A1:A10")
```

### `read_row_range(ByVal rng As Range)`

1 行の `Range` を読み込みます。

- `rng Is Nothing` の場合はエラー
- 1 行でない場合はエラー

```vb
vec.read_row_range Sheet1.Range("A1:J1")
```

### `read_array(ByVal arr As Variant)`

一次元配列を読み込みます。

- 配列でない場合はエラー
- 2 次元以上の配列はエラー

```vb
Dim arr(1 To 3) As Variant
arr(1) = 10
arr(2) = 20
arr(3) = 30

vec.read_array arr
```

## 参照

### `is_loaded As Boolean`

データが読み込まれているかを返します。

### `data As Variant`

保持している一次元配列のコピーを返します。  
返却値を変更しても、内部の配列本体には直接影響しません。

### `count As Long`

要素数を返します。

### `item(ByVal index As Long) As Variant`

指定した添字の値を返します。

- 範囲外の添字はエラー

### `type_names() As Variant`

各要素の型名を配列で返します。

- `Error` の場合は `"Error"`
- `Empty` の場合は `"Empty"`
- それ以外は `TypeName` の結果

```vb
Dim types As Variant
types = vec.type_names
```

## 型変換

### `cast_to_double_safe(Optional emptyAsZero As Boolean = False, Optional invalidAsZero As Boolean = False, Optional treatDateAsInvalid As Boolean = True)`

各要素を `Double` に寄せて変換します。

- `Empty`、エラー値、空文字は `emptyAsZero=True` なら `0`、そうでなければ `Empty`
- 日付は `treatDateAsInvalid=True` のとき無効値として扱う
- 数値に変換できるものは `CDbl` で変換
- 変換できない値は `invalidAsZero=True` なら `0`、そうでなければ `Empty`

```vb
vec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True
```

### `cast_to_string_safe(Optional emptyAsBlank As Boolean = True, Optional errorAsBlank As Boolean = True)`

各要素を文字列に変換します。

- エラー値は `errorAsBlank=True` なら空文字、そうでなければエラー
- `Empty` は `emptyAsBlank=True` なら空文字、そうでなければ `Empty`
- それ以外は `CStr` で変換

### `cast_to_date_safe(Optional invalidAsEmpty As Boolean = True)`

各要素を日付に変換します。

- エラー値、`Empty`、空文字は `Empty`
- `IsDate` で判定可能な値は `CDate` で変換
- 日付に変換できない値は `invalidAsEmpty=True` なら `Empty`、そうでなければエラー

## 出力

### `to_range_vertical(ByVal topLeft As Range)`

保持している値を縦方向に出力します。

```vb
vec.to_range_vertical Sheet1.Range("E1")
```

### `to_range_horizontal(ByVal topLeft As Range)`

保持している値を横方向に出力します。

```vb
vec.to_range_horizontal Sheet1.Range("E1")
```

### `clear()`

内部データを破棄し、未読込状態に戻します。

## 注意点

- `item` や `count` を使う前に、必ず読み込みを済ませてください
- `read_array` は一次元配列専用です
- 型変換メソッドは配列をその場で書き換えます
- `to_range_vertical` と `to_range_horizontal` は、出力先の左上セルを受け取ります
