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
| 未初期化時 | 参照系・条件マスク系・変換・更新系・集計系・出力系はエラー |

## メソッド早見表

<table>
  <thead>
    <tr>
      <th>区分</th>
      <th>メソッド</th>
      <th>ひとこと</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td rowspan="3">読込</td>
      <td><code>read_col_range</code></td>
      <td>1 列の <code>Range</code> を読む</td>
    </tr>
    <tr>
      <td><code>read_row_range</code></td>
      <td>1 行の <code>Range</code> を読む</td>
    </tr>
    <tr>
      <td><code>read_array</code></td>
      <td>一次元配列をそのまま読む</td>
    </tr>
    <tr>
      <td rowspan="5">参照</td>
      <td><code>is_loaded</code></td>
      <td>読込済みかを返す</td>
    </tr>
    <tr>
      <td><code>data</code></td>
      <td>内部配列のコピーを返す</td>
    </tr>
    <tr>
      <td><code>count</code></td>
      <td>要素数を返す</td>
    </tr>
    <tr>
      <td><code>item</code></td>
      <td>指定要素を返す</td>
    </tr>
    <tr>
      <td><code>type_names</code></td>
      <td>型名配列を返す</td>
    </tr>
    <tr>
      <td rowspan="5">条件マスク</td>
      <td><code>eq</code></td>
      <td>等値判定マスクを返す</td>
    </tr>
    <tr>
      <td><code>ne</code></td>
      <td>非等値判定マスクを返す</td>
    </tr>
    <tr>
      <td><code>gt</code> <code>ge</code></td>
      <td>下限条件の比較マスクを返す</td>
    </tr>
    <tr>
      <td><code>lt</code> <code>le</code></td>
      <td>上限条件の比較マスクを返す</td>
    </tr>
    <tr>
      <td><code>is_empty</code></td>
      <td>欠損判定マスクを返す</td>
    </tr>
    <tr>
      <td rowspan="5">変換・更新</td>
      <td><code>cast_to_double_safe</code></td>
      <td>数値変換する</td>
    </tr>
    <tr>
      <td><code>cast_to_string_safe</code></td>
      <td>文字列変換する</td>
    </tr>
    <tr>
      <td><code>cast_to_date_safe</code></td>
      <td>日付変換する</td>
    </tr>
    <tr>
      <td><code>fill_empty</code></td>
      <td>欠損を埋める</td>
    </tr>
    <tr>
      <td><code>map</code></td>
      <td>関数名で指定した変換関数を適用する</td>
    </tr>
    <tr>
      <td rowspan="3">集計</td>
      <td><code>unique</code></td>
      <td>重複なし配列を返す</td>
    </tr>
    <tr>
      <td><code>sum</code></td>
      <td>合計を返す</td>
    </tr>
    <tr>
      <td><code>mean</code></td>
      <td>平均を返す</td>
    </tr>
    <tr>
      <td rowspan="2">出力</td>
      <td><code>to_range_vertical</code></td>
      <td>縦に書き戻す</td>
    </tr>
    <tr>
      <td><code>to_range_horizontal</code></td>
      <td>横に書き戻す</td>
    </tr>
    <tr>
      <td rowspan="1">状態管理</td>
      <td><code>clear</code></td>
      <td>状態を初期化する</td>
    </tr>
  </tbody>
</table>

## 区分ごとの補足

### 条件マスク系メソッドの補足

- `eq` `ne` `gt` `ge` `lt` `le` `is_empty` は、後続処理に渡しやすい `Variant` のブール配列を返します。
- `DataTable.set_by_mask` や `Matrix.filter_rows` と組み合わせる前提で使うと効果的です。

### 変換・更新系メソッドの補足

- 変換・更新系メソッドは内部配列をその場で更新します。

### 集計系メソッドの補足

- `sum` や `mean` の前に `cast_to_double_safe` を通すと扱いやすくなります。
- `unique` は最初に出現した値を残す仕様です。

### 状態管理系メソッドの補足

- `clear` はシート出力ではなく、内部状態を未読込状態へ戻すためのメソッドです。

## よく使うレシピ

### 空欄を 0 にして数値列として合計する

```vb
Dim vec As New Vector

vec.read_col_range Sheet1.Range("D2:D8")
vec.fill_empty 0
vec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True

Debug.Print vec.sum
```

### 条件マスクを作って別モジュールへ渡す

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_col_range Sheet1.Range("C2:C8")
mask = vec.eq("対象")
```

### `DataTable` から列を取り出して平均を計算する

```vb
Dim tbl As New DataTable
Dim scoreVec As Vector

tbl.read_range Sheet1.Range("A1:D10"), hasHeader:=True
Set scoreVec = tbl.col_vector("score")

scoreVec.cast_to_double_safe emptyAsZero:=True, invalidAsZero:=True
Debug.Print scoreVec.mean
```

## 全メソッド解説

### 読込系

<details>
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

### 参照系

<details>
<summary><code>is_loaded() As Boolean</code></summary>

読込済みかを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | なし |
| 戻り値 | `Boolean` |
| 実行内容 | 読込済みなら `True`、未読込なら `False` を返す |
| 代表ユースケース | 実行前チェック、デバッグ確認 |

```vb
Dim vec As New Vector

Debug.Print vec.is_loaded
vec.read_array Array(1, 2, 3)
Debug.Print vec.is_loaded
```

出力イメージ:

| 実行順 | 値 |
| --- | --- |
| 読込前 | `False` |
| 読込後 | `True` |

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
<summary><code>count() As Long</code></summary>

要素数を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Long` |
| 実行内容 | 内部配列の要素数を返す |
| 代表ユースケース | ループ回数の決定、件数確認 |

```vb
Dim vec As New Vector

vec.read_array Array("A", "B", "C")
Debug.Print vec.count
```

出力イメージ:

| 項目 | 値 |
| --- | --- |
| `vec.count` | `3` |

</details>

<details>
<summary><code>item(ByVal index As Long) As Variant</code></summary>

指定要素を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`index` が範囲内であること |
| 入力 | `index As Long` |
| 戻り値 | 指定位置の `Variant` 値 |
| 注意点 | 範囲外の添字はエラー |
| 代表ユースケース | 特定位置の値確認 |

```vb
Dim vec As New Vector

vec.read_array Array("A", "B", "C")
Debug.Print vec.item(1)
```

出力イメージ:

| 項目 | 値 |
| --- | --- |
| `vec.item(1)` | `"B"` |

</details>

<details>
<summary><code>type_names() As Variant</code></summary>

各要素の型名を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Variant` の一次元配列 |
| 実行内容 | 各要素に対応する `TypeName`、または `Error` / `Empty` を返す |
| 代表ユースケース | 型混在データの調査 |

```vb
Dim vec As New Vector
Dim result As Variant

vec.read_array Array(10, "A", Empty)
result = vec.type_names
```

出力イメージ:

| 行 | 型名 |
| --- | --- |
| 1 | `Integer` または `Long` |
| 2 | `String` |
| 3 | `Empty` |

</details>

### 条件マスク系

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
<summary><code>ne(ByVal matchValue As Variant) As Variant</code></summary>

非等値判定マスクを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `matchValue As Variant` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | `Null` / `Error` / 比較不能値は `False` |
| 代表ユースケース | 特定値以外の抽出条件作成 |

入力イメージ:

| 値 |
| --- |
| A |
| B |
| A |

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_array Array("A", "B", "A")
mask = vec.ne("A")
```

出力イメージ:

| 行 | 判定結果 |
| --- | --- |
| 1 | `False` |
| 2 | `True` |
| 3 | `False` |

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
<summary><code>ge(ByVal threshold As Variant) As Variant</code></summary>

`>=` 判定マスクを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `threshold As Variant` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | 比較不能値は `False` |
| 代表ユースケース | 下限値以上の抽出 |

入力イメージ:

| 値 |
| --- |
| 80 |
| 100 |
| 150 |

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_array Array(80, 100, 150)
mask = vec.ge(100)
```

出力イメージ:

| 行 | 判定結果 |
| --- | --- |
| 1 | `False` |
| 2 | `True` |
| 3 | `True` |

</details>

<details>
<summary><code>lt(ByVal threshold As Variant) As Variant</code></summary>

`<` 判定マスクを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `threshold As Variant` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | 比較不能値は `False` |
| 代表ユースケース | 上限値未満の抽出 |

入力イメージ:

| 値 |
| --- |
| 80 |
| 100 |
| 150 |

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_array Array(80, 100, 150)
mask = vec.lt(100)
```

出力イメージ:

| 行 | 判定結果 |
| --- | --- |
| 1 | `True` |
| 2 | `False` |
| 3 | `False` |

</details>

<details>
<summary><code>le(ByVal threshold As Variant) As Variant</code></summary>

`<=` 判定マスクを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `threshold As Variant` |
| 戻り値 | `Variant` のブール配列 |
| 注意点 | 比較不能値は `False` |
| 代表ユースケース | 上限値以下の抽出 |

入力イメージ:

| 値 |
| --- |
| 80 |
| 100 |
| 150 |

```vb
Dim vec As New Vector
Dim mask As Variant

vec.read_array Array(80, 100, 150)
mask = vec.le(100)
```

出力イメージ:

| 行 | 判定結果 |
| --- | --- |
| 1 | `True` |
| 2 | `True` |
| 3 | `False` |

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

### 変換・更新系

<details>
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
<summary><code>cast_to_string_safe(Optional emptyAsBlank As Boolean = True, Optional errorAsBlank As Boolean = True)</code></summary>

要素を文字列へ変換します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `emptyAsBlank As Boolean` `errorAsBlank As Boolean` |
| 出力 | 内部配列を上書き更新 |
| 実行内容 | `CStr` を使って文字列化し、空やエラー値は設定に応じて処理する |
| 注意点 | `errorAsBlank=False` でエラー値を含むとエラー |
| 代表ユースケース | 表示用整形、文字列比較前の変換 |

入力イメージ:

| 値 |
| --- |
| 100 |
| Empty |
| `ABC` |

```vb
Dim vec As New Vector

vec.read_array Array(100, Empty, "ABC")
vec.cast_to_string_safe emptyAsBlank:=True, errorAsBlank:=True
```

実行後イメージ:

| 行 | 変換後 |
| --- | --- |
| 1 | `"100"` |
| 2 | `""` |
| 3 | `"ABC"` |

</details>

<details>
<summary><code>cast_to_date_safe(Optional ByVal invalidAsEmpty As Boolean = True)</code></summary>

要素を日付へ変換します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `invalidAsEmpty As Boolean` |
| 出力 | 内部配列を上書き更新 |
| 実行内容 | `IsDate` 判定可能な値を `CDate` 変換する |
| 注意点 | 日付化不能値は `Empty` またはエラー |
| 代表ユースケース | 日付列の正規化 |

入力イメージ:

| 値 |
| --- |
| `2024/01/01` |
| `abc` |
| Empty |

```vb
Dim vec As New Vector

vec.read_array Array("2024/01/01", "abc", Empty)
vec.cast_to_date_safe True
```

実行後イメージ:

| 行 | 変換後 |
| --- | --- |
| 1 | `Date` 値 |
| 2 | `Empty` |
| 3 | `Empty` |

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
<summary><code>map(ByVal functionName As String)</code></summary>

関数名で指定した変換関数を各要素へ適用します。  
VBA では関数参照を直接渡しにくいため、実装上は `Application.Run` で呼べる公開関数名を受け取る高階関数風メソッドです。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`functionName` が空でないこと、`Application.Run` で呼べる公開関数があること |
| 入力 | `functionName As String` (`Public Function` 名) |
| 出力 | 内部配列を上書き更新 |
| 実行内容 | 各要素に対して `(value, index)` を渡して関数実行し、戻り値で置換する |
| 注意点 | コールバック関数は `value As Variant, index As Long` の 2 引数を受け取る形で定義する。関数名誤りや実行失敗時はエラー |
| 代表ユースケース | 独自整形ルールの一括適用 |

```vb
Public Function AddPrefix(ByVal value As Variant, ByVal index As Long) As Variant
    AddPrefix = "ID" & Format$(index, "00") & "-" & CStr(value)
End Function

Dim vec As New Vector

vec.read_array Array(10, 20, 30)
vec.map "AddPrefix"
```

実行後イメージ:

| 行 | 変換後 |
| --- | --- |
| 1 | `"ID01-10"` |
| 2 | `"ID02-20"` |
| 3 | `"ID03-30"` |

</details>

### 集計系

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

<details>
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

### 出力系

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

入力イメージ:

| 内部配列 |
| --- |
| A |
| B |
| C |

```vb
Dim vec As New Vector

vec.read_array Array("A", "B", "C")
vec.to_range_horizontal Sheet1.Range("H2")
```

出力イメージ:

| H | I | J |
| --- | --- | --- |
| A | B | C |

</details>

### 状態管理系

<details>
<summary><code>clear()</code></summary>

内部状態を初期化します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | なし |
| 入力 | なし |
| 出力 | 内部状態を未読込状態へ戻す |
| 実行内容 | `VECTOR` を空にし、`is_loaded` が `False` になる |
| 注意点 | データを保持したままには戻せないため、必要なら `data` を先に取得する |
| 代表ユースケース | 同じインスタンスを再利用する前に状態をリセットする |

```vb
Dim vec As New Vector

vec.read_array Array("A", "B", "C")
vec.clear

Debug.Print vec.is_loaded
```

実行後イメージ:

| 確認項目 | 値 |
| --- | --- |
| `vec.is_loaded` | `False` |

</details>
