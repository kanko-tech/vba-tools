# クラスモジュール `Vector`

`Vector` は、一次元配列を扱うための VBA クラスモジュールです。  
列データの読込、条件マスク生成、型変換、集計、セル出力を 1 つのオブジェクトに集約します。

## 概要

| 項目 | 内容 |
| --- | --- |
| 主な用途 | 1 列または 1 行のデータを一次元配列として扱う |
| 内部データ | `Variant` の一次元配列 |
| 得意な処理 | 型変換、欠損補完、条件マスク生成、集計 |
| 他モジュールとの関係 | `Table.col_vector` から列データを受ける想定 |

## 利用前提

| 項目 | 内容 |
| --- | --- |
| 初期化必須 | `read_col_range` / `read_row_range` / `read_array` のいずれかを先に実行 |
| 添字 | 入力配列の添字を引き継ぐ |
| 未初期化時 | 参照系・加工系・集計系・出力系はエラー |
| データ形状 | 一次元配列のみ対応 |

## 典型的なユースケース

- シートの 1 列を読み込んで数値化する
- 列値から `eq` や `gt` で条件マスクを作る
- `sum` `mean` `unique` で列集計する
- `Table` から取り出した 1 列を単独処理する

## メソッド一覧

| 区分 | メソッド |
| --- | --- |
| 読込 | `read_col_range` `read_row_range` `read_array` |
| 参照 | `is_loaded` `data` `count` `item` `type_names` |
| 条件 | `eq` `ne` `gt` `ge` `lt` `le` `is_empty` |
| 加工 | `cast_to_double_safe` `cast_to_string_safe` `cast_to_date_safe` `fill_empty` `map` |
| 集計 | `unique` `sum` `mean` |
| 出力 | `to_range_vertical` `to_range_horizontal` `clear` |

## 読込系

### `read_col_range(ByVal rng As Range)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 1 列の `Range` を一次元配列として読み込む |
| 前提条件 | `rng` が `Nothing` でないこと、1 列であること |
| 入力 | `rng As Range` |
| 実行内容 | 単一セルなら 1 要素配列、複数セルなら縦方向データを配列化して保持 |
| 出力 | 内部状態 `VECTOR` を更新 |
| 注意点 | 2 列以上の `Range` はエラー |
| ユースケース | シート上の縦持ちデータの読込 |

ユースケース例: 売上列を読み込んで、このあと数値変換や集計に回す。

```vb
Dim vec As New Vector
vec.read_col_range Sheet1.Range("B2:B10")
```

### `read_row_range(ByVal rng As Range)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 1 行の `Range` を一次元配列として読み込む |
| 前提条件 | `rng` が `Nothing` でないこと、1 行であること |
| 入力 | `rng As Range` |
| 実行内容 | 横方向データを一次元配列として保持 |
| 出力 | 内部状態 `VECTOR` を更新 |
| 注意点 | 2 行以上の `Range` はエラー |
| ユースケース | 横持ちデータの一括処理 |

### `read_array(ByVal arr As Variant)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 一次元配列をそのまま取り込む |
| 前提条件 | `arr` が配列であり、一次元配列であること |
| 入力 | `arr As Variant` |
| 実行内容 | 内部配列として保持 |
| 出力 | 内部状態 `VECTOR` を更新 |
| 注意点 | 2 次元以上の配列はエラー |
| ユースケース | 他モジュールや関数から受け取った列データの取込 |

## 参照系

### `is_loaded() As Boolean`

| 項目 | 内容 |
| --- | --- |
| 役割 | 読込済みかを返す |
| 戻り値 | `Boolean` |
| ユースケース | 実行前ガード、デバッグ確認 |

### `data() As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 内部一次元配列のコピーを返す |
| 前提条件 | 読込済みであること |
| 戻り値 | `Variant` の一次元配列 |
| 注意点 | 返却後に変更しても内部配列は直接変わらない |
| ユースケース | 外部関数への受け渡し、ログ出力 |

### `count() As Long`

| 項目 | 内容 |
| --- | --- |
| 役割 | 要素数を返す |
| 前提条件 | 読込済みであること |
| 戻り値 | `Long` |
| ユースケース | ループ回数の決定、件数確認 |

### `item(ByVal index As Long) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定添字の値を返す |
| 前提条件 | 読込済みであること、`index` が範囲内であること |
| 入力 | `index As Long` |
| 戻り値 | 対象要素の `Variant` 値 |
| 注意点 | 範囲外添字はエラー |
| ユースケース | 特定レコード位置の確認 |

### `type_names() As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 各要素の型名を返す |
| 前提条件 | 読込済みであること |
| 戻り値 | `Variant` の一次元配列 |
| 注意点 | `Error` と `Empty` は固定文字列で返す |
| ユースケース | 型混在データの調査 |

## 条件マスク系

### `eq(ByVal matchValue As Variant) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 等値判定マスクを返す |
| 前提条件 | 読込済みであること |
| 入力 | `matchValue As Variant` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | `Null` / `Error` / 比較不能値は `False` |
| ユースケース | `Table.set_by_mask` や `Matrix.filter_rows` に渡す条件生成 |

ユースケース例: 区分列から `"対象"` の行だけを後段処理したい。

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_col_range Sheet1.Range("C2:C10")
mask = vec.eq("対象")
```

### `ne(ByVal matchValue As Variant) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 非等値判定マスクを返す |
| 前提条件 | 読込済みであること |
| 入力 | `matchValue As Variant` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | `eq` の反転版だが比較不能値は `False` |
| ユースケース | 特定値以外の抽出条件生成 |

### `gt(ByVal threshold As Variant) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | `>` 判定マスクを返す |
| 前提条件 | 読込済みであること |
| 入力 | `threshold As Variant` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | 比較不能値は `False` |
| ユースケース | 閾値超過データの抽出 |

ユースケース例: 売上が 100 を超えるレコードだけを抽出する条件を作る。

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_col_range Sheet1.Range("D2:D10")
vec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True
mask = vec.gt(100)
```

### `ge(ByVal threshold As Variant) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | `>=` 判定マスクを返す |
| 前提条件 | 読込済みであること |
| 入力 | `threshold As Variant` |
| 戻り値 | `Variant` のブール配列 |
| ユースケース | 下限値チェック |

### `lt(ByVal threshold As Variant) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | `<` 判定マスクを返す |
| 前提条件 | 読込済みであること |
| 入力 | `threshold As Variant` |
| 戻り値 | `Variant` のブール配列 |
| ユースケース | 上限制約の判定 |

### `le(ByVal threshold As Variant) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | `<=` 判定マスクを返す |
| 前提条件 | 読込済みであること |
| 入力 | `threshold As Variant` |
| 戻り値 | `Variant` のブール配列 |
| ユースケース | 上限値以下の抽出 |

### `is_empty(Optional ByVal treatBlankStringAsEmpty As Boolean = True) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 欠損判定マスクを返す |
| 前提条件 | 読込済みであること |
| 入力 | `treatBlankStringAsEmpty As Boolean` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | `True` の場合は空文字も欠損扱い |
| ユースケース | 欠損行抽出、補完前確認 |

## 型・加工系

### `cast_to_double_safe(Optional emptyAsZero As Boolean = False, Optional invalidAsZero As Boolean = False, Optional treatDateAsInvalid As Boolean = True)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 要素を `Double` に寄せて変換する |
| 前提条件 | 読込済みであること |
| 入力 | `emptyAsZero As Boolean` `invalidAsZero As Boolean` `treatDateAsInvalid As Boolean` |
| 実行内容 | 数値化できるものを `CDbl` 変換し、空や不正値を設定に応じて `0` または `Empty` にする |
| 出力 | 内部配列を上書き更新 |
| 注意点 | 日付は設定により無効値扱い |
| ユースケース | 集計前の数値整形 |

ユースケース例: 数値列を集計前に正規化して、不正値を 0 に寄せる。

```vb
Dim vec As New Vector

vec.read_col_range Sheet1.Range("B2:B10")
vec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True
```

### `cast_to_string_safe(Optional emptyAsBlank As Boolean = True, Optional errorAsBlank As Boolean = True)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 要素を文字列に変換する |
| 前提条件 | 読込済みであること |
| 入力 | `emptyAsBlank As Boolean` `errorAsBlank As Boolean` |
| 実行内容 | `CStr` ベースで文字列化し、空やエラー値は設定に応じて置換 |
| 出力 | 内部配列を上書き更新 |
| 注意点 | `errorAsBlank=False` かつエラー値を含むとエラー |
| ユースケース | 文字列比較前の整形、表示用整形 |

### `cast_to_date_safe(Optional ByVal invalidAsEmpty As Boolean = True)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 要素を日付に変換する |
| 前提条件 | 読込済みであること |
| 入力 | `invalidAsEmpty As Boolean` |
| 実行内容 | `IsDate` 判定可能な値を `CDate` 変換する |
| 出力 | 内部配列を上書き更新 |
| 注意点 | 日付化不能値は `Empty` またはエラー |
| ユースケース | 日付列の正規化 |

### `fill_empty(ByVal fillValue As Variant, Optional ByVal treatBlankStringAsEmpty As Boolean = True)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 欠損値を指定値で埋める |
| 前提条件 | 読込済みであること |
| 入力 | `fillValue As Variant` `treatBlankStringAsEmpty As Boolean` |
| 実行内容 | `Empty` と必要に応じて空文字を `fillValue` に置換 |
| 出力 | 内部配列を上書き更新 |
| ユースケース | 欠損補完、既定値設定 |

ユースケース例: 空欄を 0 で埋めてから集計したい。

```vb
Dim vec As New Vector

vec.read_col_range Sheet1.Range("E2:E10")
vec.fill_empty 0
```

### `map(ByVal functionName As String)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 公開関数を各要素へ適用する |
| 前提条件 | 読込済みであること、`functionName` が空でないこと、`Application.Run` で呼べる公開関数があること |
| 入力 | `functionName As String` |
| 実行内容 | 各要素に対して `(value, index)` を引数に関数実行し、戻り値で置換 |
| 出力 | 内部配列を上書き更新 |
| 注意点 | 関数名誤りや実行失敗時はエラー |
| ユースケース | 独自ルールの一括整形 |

## 集計系

### `unique() As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 重複除去済み配列を返す |
| 前提条件 | 読込済みであること |
| 戻り値 | `Variant` の一次元配列 |
| 注意点 | 最初に出現した値を残す |
| ユースケース | コード一覧、カテゴリ一覧の作成 |

ユースケース例: 担当者列から重複なし一覧を作る。

```vb
Dim vec As New Vector
Dim names As Variant

vec.read_col_range Sheet1.Range("A2:A20")
names = vec.unique()
```

### `sum() As Double`

| 項目 | 内容 |
| --- | --- |
| 役割 | 合計値を返す |
| 前提条件 | 読込済みであること、すべての要素が数値型であること |
| 戻り値 | `Double` |
| 注意点 | `Empty` `Null` `Error` を含むとエラー |
| ユースケース | 件数列や金額列の合計 |

ユースケース例: 売上列の合計を計算する。

```vb
Dim vec As New Vector

vec.read_col_range Sheet1.Range("F2:F10")
vec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True
Debug.Print vec.sum
```

### `mean() As Double`

| 項目 | 内容 |
| --- | --- |
| 役割 | 平均値を返す |
| 前提条件 | 読込済みであること、すべての要素が数値型であること |
| 戻り値 | `Double` |
| 注意点 | `sum / count` で計算する |
| ユースケース | 平均単価、平均点の計算 |

## 出力系

### `to_range_vertical(ByVal topLeft As Range)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 縦方向に出力する |
| 前提条件 | 読込済みであること、`topLeft` が `Nothing` でないこと |
| 入力 | `topLeft As Range` |
| 実行内容 | `topLeft` を左上として 1 列に書き込む |
| 出力 | ワークシート上のセル範囲 |
| ユースケース | 一次元配列の列方向出力 |

ユースケース例: 加工済みの列データを別列へ書き戻す。

```vb
Dim vec As New Vector

vec.read_col_range Sheet1.Range("B2:B10")
vec.to_range_vertical Sheet1.Range("H2")
```

### `to_range_horizontal(ByVal topLeft As Range)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 横方向に出力する |
| 前提条件 | 読込済みであること、`topLeft` が `Nothing` でないこと |
| 入力 | `topLeft As Range` |
| 実行内容 | `topLeft` を左上として 1 行に書き込む |
| 出力 | ワークシート上のセル範囲 |
| ユースケース | 一次元配列の行方向出力 |

### `clear()`

| 項目 | 内容 |
| --- | --- |
| 役割 | 内部状態を初期化する |
| 実行内容 | 保持配列を破棄し、未読込状態へ戻す |
| ユースケース | 再利用前の初期化 |

## 補足

- 条件系メソッドは `Variant` のブール配列を返します
- 加工系メソッドは内部配列をその場で更新します
- 集計前に `cast_to_double_safe` を通すと扱いやすくなります
