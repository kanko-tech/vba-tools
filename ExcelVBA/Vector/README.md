# クラスモジュール `Vector`

`Vector` は、一次元配列を扱うための VBA クラスモジュールです。  
1 列または 1 行のデータを読み込み、条件マスク生成、型変換、集計、セル出力を行います。

## まず何ができるか

- 1 列または 1 行の `Range` を一次元配列として保持する
- 一次元配列を直接読み込む
- `eq` `gt` `is_empty` などで条件マスクを作る
- `cast_to_double_safe` `fill_empty` などで列を整形する
- `sum` `mean` `unique` で列を集計する
- シートに縦方向・横方向で書き戻す

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

## 利用前提

| 項目 | 内容 |
| --- | --- |
| 初期化必須 | `read_col_range` / `read_row_range` / `read_array` のいずれかを先に実行 |
| データ形状 | 一次元配列のみ対応 |
| 添字 | 入力配列の添字を引き継ぐ |
| 未初期化時 | 参照系・加工系・集計系・出力系はエラー |

## メソッド早見表

| 区分 | メソッド | ひとこと |
| --- | --- | --- |
| 読込 | `read_col_range` | 1 列の `Range` を読む |
| 読込 | `read_row_range` | 1 行の `Range` を読む |
| 読込 | `read_array` | 一次元配列をそのまま読む |
| 参照 | `is_loaded` | 読込済みかを返す |
| 参照 | `data` | 内部配列のコピーを返す |
| 参照 | `count` | 要素数を返す |
| 参照 | `item` | 指定要素を返す |
| 参照 | `type_names` | 型名配列を返す |
| 条件 | `eq` | 等値判定マスクを返す |
| 条件 | `ne` | 非等値判定マスクを返す |
| 条件 | `gt` `ge` `lt` `le` | 比較演算マスクを返す |
| 条件 | `is_empty` | 欠損判定マスクを返す |
| 加工 | `cast_to_double_safe` | 数値変換する |
| 加工 | `cast_to_string_safe` | 文字列変換する |
| 加工 | `cast_to_date_safe` | 日付変換する |
| 加工 | `fill_empty` | 欠損を埋める |
| 加工 | `map` | 公開関数を各要素へ適用する |
| 集計 | `unique` | 重複なし配列を返す |
| 集計 | `sum` | 合計を返す |
| 集計 | `mean` | 平均を返す |
| 出力 | `to_range_vertical` | 縦に書き戻す |
| 出力 | `to_range_horizontal` | 横に書き戻す |
| 出力 | `clear` | 状態を初期化する |

## よく使うレシピ

### 空欄を 0 にして数値列として合計する

```vb
Dim vec As New Vector

vec.read_col_range Sheet1.Range("D2:D8")
vec.fill_empty 0
vec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True

Debug.Print vec.sum
```

### 特定の値に一致する行の条件マスクを作る

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_col_range Sheet1.Range("C2:C8")
mask = vec.eq("対象")
```

### `Table` から列を取り出して平均を計算する

```vb
Dim tbl As New Table
Dim scoreVec As Vector

tbl.read_range Sheet1.Range("A1:D10"), hasHeader:=True
Set scoreVec = tbl.col_vector("score")

scoreVec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True
Debug.Print scoreVec.mean
```

## 詳細仕様

<details open>
<summary><code>read_col_range(ByVal rng As Range)</code></summary>

1 列の `Range` を一次元配列として読み込みます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | `rng` が `Nothing` でないこと、1 列であること |
| 入力 | `rng As Range` |
| 出力 | 内部状態 `VECTOR` を更新 |
| 実行内容 | 単一セルなら 1 要素配列、複数セルなら縦方向の値を保持 |
| 注意点 | 2 列以上の `Range` はエラー |
| 代表ユースケース | シート上の縦持ち列を処理対象にする |

入力イメージ:

| B列 |
| --- |
| 120 |
| 80 |
| 150 |

```vb
Dim vec As New Vector
vec.read_col_range Sheet1.Range("B2:B4")
```

実行後イメージ:

| 確認項目 | 値 |
| --- | --- |
| `vec.count` | `3` |
| `vec.item(1)` | `120` |
| `vec.data` | `[120, 80, 150]` |

</details>

<details>
<summary><code>read_row_range(ByVal rng As Range)</code></summary>

1 行の `Range` を一次元配列として読み込みます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | `rng` が `Nothing` でないこと、1 行であること |
| 入力 | `rng As Range` |
| 出力 | 内部状態 `VECTOR` を更新 |
| 実行内容 | 横方向の値を一次元配列として保持 |
| 注意点 | 2 行以上の `Range` はエラー |
| 代表ユースケース | 横持ちデータの一括処理 |

入力イメージ:

| B | C | D |
| --- | --- | --- |
| A | B | C |

```vb
Dim vec As New Vector
vec.read_row_range Sheet1.Range("B2:D2")
```

実行後イメージ:

| 確認項目 | 値 |
| --- | --- |
| `vec.count` | `3` |
| `vec.item(2)` | `"B"` |
| `vec.data` | `["A", "B", "C"]` |

</details>

<details>
<summary><code>read_array(ByVal arr As Variant)</code></summary>

一次元配列をそのまま読み込みます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | `arr` が一次元配列であること |
| 入力 | `arr As Variant` |
| 出力 | 内部状態 `VECTOR` を更新 |
| 実行内容 | 渡された配列を内部データとして保持 |
| 注意点 | 2 次元以上の配列はエラー |
| 代表ユースケース | 他関数や他モジュールの戻り値を受ける |

```vb
Dim vec As New Vector
Dim arr(1 To 3) As Variant

arr(1) = 10
arr(2) = 20
arr(3) = 30

vec.read_array arr
```

実行後イメージ:

| 確認項目 | 値 |
| --- | --- |
| `vec.count` | `3` |
| `vec.data` | `[10, 20, 30]` |

</details>

<details>
<summary><code>data() As Variant</code></summary>

内部配列のコピーを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Variant` の一次元配列 |
| 実行内容 | 内部配列の複製を返す |
| 注意点 | 返却値を書き換えても内部配列は直接変わらない |
| 代表ユースケース | 他処理に安全に渡す |

```vb
Dim vec As New Vector
Dim values As Variant

vec.read_array Array(10, 20, 30)
values = vec.data
```

返却イメージ:

| 添字 | 値 |
| --- | --- |
| 0 | `10` |
| 1 | `20` |
| 2 | `30` |

</details>

<details>
<summary><code>eq(ByVal matchValue As Variant) As Variant</code></summary>

等値判定マスクを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `matchValue As Variant` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | `Null` / `Error` / 比較不能値は `False` |
| 代表ユースケース | 条件抽出用マスクを作る |

入力イメージ:

| 値 |
| --- |
| 対象 |
| 対象外 |
| 対象 |

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_col_range Sheet1.Range("C2:C4")
mask = vec.eq("対象")
```

出力イメージ:

| 行 | 判定結果 |
| --- | --- |
| 1 | `True` |
| 2 | `False` |
| 3 | `True` |

</details>

<details>
<summary><code>gt(ByVal threshold As Variant) As Variant</code></summary>

`>` 判定マスクを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `threshold As Variant` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | 比較不能値は `False` |
| 代表ユースケース | 閾値超過データの抽出 |

入力イメージ:

| 値 |
| --- |
| 80 |
| 120 |
| 150 |

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_col_range Sheet1.Range("D2:D4")
vec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True
mask = vec.gt(100)
```

出力イメージ:

| 行 | 判定結果 |
| --- | --- |
| 1 | `False` |
| 2 | `True` |
| 3 | `True` |

</details>

<details>
<summary><code>is_empty(Optional ByVal treatBlankStringAsEmpty As Boolean = True) As Variant</code></summary>

欠損判定マスクを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `treatBlankStringAsEmpty As Boolean` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | `True` の場合は空文字も欠損扱い |
| 代表ユースケース | 欠損行抽出、補完対象確認 |

入力イメージ:

| 値 |
| --- |
| 100 |
| `""` |
| Empty |

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_array Array(100, "", Empty)
mask = vec.is_empty(True)
```

出力イメージ:

| 行 | 判定結果 |
| --- | --- |
| 1 | `False` |
| 2 | `True` |
| 3 | `True` |

</details>

<details open>
<summary><code>cast_to_double_safe(Optional emptyAsZero As Boolean = False, Optional invalidAsZero As Boolean = False, Optional treatDateAsInvalid As Boolean = True)</code></summary>

要素を `Double` に寄せて変換します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `emptyAsZero As Boolean` `invalidAsZero As Boolean` `treatDateAsInvalid As Boolean` |
| 出力 | 内部配列を上書き更新 |
| 実行内容 | 数値化できる要素を `CDbl` 変換し、空・不正値を設定に応じて `0` または `Empty` にする |
| 注意点 | 日付は設定により無効値扱いになる |
| 代表ユースケース | 集計前の数値整形 |

入力イメージ:

| 値 |
| --- |
| `"120"` |
| `""` |
| `"abc"` |

```vb
Dim vec As New Vector

vec.read_array Array("120", "", "abc")
vec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True
```

実行後イメージ:

| 行 | 変換後 |
| --- | --- |
| 1 | `120` |
| 2 | `0` |
| 3 | `0` |

</details>

<details>
<summary><code>fill_empty(ByVal fillValue As Variant, Optional ByVal treatBlankStringAsEmpty As Boolean = True)</code></summary>

欠損値を指定値で埋めます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `fillValue As Variant` `treatBlankStringAsEmpty As Boolean` |
| 出力 | 内部配列を上書き更新 |
| 実行内容 | `Empty` と必要に応じて空文字を `fillValue` に置換 |
| 注意点 | 欠損以外の値は変更しない |
| 代表ユースケース | 欠損補完、既定値設定 |

入力イメージ:

| 値 |
| --- |
| 10 |
| `""` |
| Empty |

```vb
Dim vec As New Vector

vec.read_array Array(10, "", Empty)
vec.fill_empty 0
```

実行後イメージ:

| 行 | 変換後 |
| --- | --- |
| 1 | `10` |
| 2 | `0` |
| 3 | `0` |

</details>

<details>
<summary><code>unique() As Variant</code></summary>

重複除去済み配列を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Variant` の一次元配列 |
| 注意点 | 最初に出現した値を残す |
| 代表ユースケース | コード一覧、カテゴリ一覧の作成 |

入力イメージ:

| 値 |
| --- |
| A |
| B |
| A |
| C |

```vb
Dim vec As New Vector
Dim result As Variant

vec.read_array Array("A", "B", "A", "C")
result = vec.unique()
```

出力イメージ:

| 添字 | 値 |
| --- | --- |
| 1 | `A` |
| 2 | `B` |
| 3 | `C` |

</details>

<details open>
<summary><code>sum() As Double</code></summary>

合計値を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、すべての要素が数値型であること |
| 戻り値 | `Double` |
| 注意点 | `Empty` `Null` `Error` を含むとエラー |
| 代表ユースケース | 売上列、件数列の合計 |

入力イメージ:

| 値 |
| --- |
| 100 |
| 200 |
| 300 |

```vb
Dim vec As New Vector

vec.read_array Array(100, 200, 300)
Debug.Print vec.sum
```

出力イメージ:

| 項目 | 値 |
| --- | --- |
| `vec.sum` | `600` |

</details>

<details>
<summary><code>mean() As Double</code></summary>

平均値を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、すべての要素が数値型であること |
| 戻り値 | `Double` |
| 注意点 | `sum / count` で計算する |
| 代表ユースケース | 平均単価、平均点の計算 |

入力イメージ:

| 値 |
| --- |
| 100 |
| 200 |
| 300 |

```vb
Dim vec As New Vector

vec.read_array Array(100, 200, 300)
Debug.Print vec.mean
```

出力イメージ:

| 項目 | 値 |
| --- | --- |
| `vec.mean` | `200` |

</details>

<details>
<summary><code>to_range_vertical(ByVal topLeft As Range)</code></summary>

内部データを縦方向に書き戻します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`topLeft` が `Nothing` でないこと |
| 入力 | `topLeft As Range` |
| 出力 | ワークシート上のセル範囲 |
| 実行内容 | `topLeft` を左上として 1 列に書込み |
| 代表ユースケース | 加工済み列データをシートへ戻す |

入力イメージ:

| 内部配列 |
| --- |
| 10 |
| 20 |
| 30 |

```vb
Dim vec As New Vector

vec.read_array Array(10, 20, 30)
vec.to_range_vertical Sheet1.Range("H2")
```

出力イメージ:

| H列 |
| --- |
| 10 |
| 20 |
| 30 |

</details>

<details>
<summary><code>to_range_horizontal(ByVal topLeft As Range)</code></summary>

内部データを横方向に書き戻します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`topLeft` が `Nothing` でないこと |
| 入力 | `topLeft As Range` |
| 出力 | ワークシート上のセル範囲 |
| 実行内容 | `topLeft` を左上として 1 行に書込み |
| 代表ユースケース | 一次元配列を横持ちで出力する |

</details>

<details>
<summary><code>clear()</code></summary>

内部状態を初期化します。

| 項目 | 内容 |
| --- | --- |
| 実行内容 | 保持配列を破棄し、未読込状態へ戻す |
| 出力 | 内部状態を初期化 |
| 代表ユースケース | 再利用前のリセット |

</details>

## 補足

- 条件系メソッドは `Variant` のブール配列を返します
- 加工系メソッドは内部配列をその場で更新します
- 主要メソッドは開いた状態にしているので、まずそこから読むと全体像を掴みやすいです
