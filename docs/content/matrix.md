# クラスモジュール `Matrix`

`Matrix` は、2 次元配列を扱うための VBA クラスモジュールです。  
列名を持たない表データを保持し、行抽出、列選択、更新、転置、シート出力を行います。

## まず何ができるか

- 矩形 `Range` や 2 次元配列を、そのまま表データとして保持する
- 行マスクで必要な行だけを抜き出す
- 必要な列だけを選んで新しい `Matrix` を作る
- 列の差し替え、行追加、列追加で表構造を更新する
- 転置した `Matrix` を作る
- シートへ表を一括で書き戻す

## クイックスタート

```vb
Sub Sample_Matrix_QuickStart()
    Dim mat As New Matrix
    Dim filtered As Matrix
    Dim picked As Matrix
    Dim mask(1 To 4) As Boolean

    mat.read_range Sheet1.Range("A2:C5")

    mask(1) = True
    mask(2) = False
    mask(3) = True
    mask(4) = False

    Set filtered = mat.filter_rows(mask)
    Set picked = filtered.select_columns(Array(1, 3))

    picked.to_range Sheet1.Range("F2")
End Sub
```

## 利用前提

| 項目 | 内容 |
| --- | --- |
| 初期化必須 | `read_range` / `read_array` / `read_empty` のいずれかを先に実行 |
| データ形状 | `read_array` は 2 次元配列専用 |
| 添字 | 内部では 1 始まりへ正規化 |
| 未初期化時 | 参照系・抽出選択系・更新追加系・変形系・出力系はエラー |
| 空表 | 行数 0 でも列数は保持可能 |

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
      <td><code>read_range</code></td>
      <td>矩形 <code>Range</code> を読む</td>
    </tr>
    <tr>
      <td><code>read_array</code></td>
      <td>2 次元配列を読む</td>
    </tr>
    <tr>
      <td><code>read_empty</code></td>
      <td>列数だけを持つ空表を作る</td>
    </tr>
    <tr>
      <td rowspan="7">参照</td>
      <td><code>is_loaded</code></td>
      <td>読込済みかを返す</td>
    </tr>
    <tr>
      <td><code>row_count</code></td>
      <td>行数を返す</td>
    </tr>
    <tr>
      <td><code>col_count</code></td>
      <td>列数を返す</td>
    </tr>
    <tr>
      <td><code>data</code></td>
      <td>内部配列のコピーを返す</td>
    </tr>
    <tr>
      <td><code>item</code></td>
      <td>指定位置の値を返す</td>
    </tr>
    <tr>
      <td><code>row_values</code></td>
      <td>指定行を一次元配列で返す</td>
    </tr>
    <tr>
      <td><code>col_values</code></td>
      <td>指定列を一次元配列で返す</td>
    </tr>
    <tr>
      <td rowspan="3">抽出・選択</td>
      <td><code>filter_rows</code></td>
      <td>行マスクで絞り込む</td>
    </tr>
    <tr>
      <td><code>slice_rows</code></td>
      <td>指定行番号だけを取り出す</td>
    </tr>
    <tr>
      <td><code>select_columns</code></td>
      <td>指定列だけを選ぶ</td>
    </tr>
    <tr>
      <td rowspan="3">更新・追加</td>
      <td><code>set_column_values</code></td>
      <td>列全体を差し替える</td>
    </tr>
    <tr>
      <td><code>append_row</code></td>
      <td>末尾に 1 行追加する</td>
    </tr>
    <tr>
      <td><code>append_column</code></td>
      <td>末尾に 1 列追加する</td>
    </tr>
    <tr>
      <td rowspan="1">変形</td>
      <td><code>transpose</code></td>
      <td>行列を入れ替えた新しい表を返す</td>
    </tr>
    <tr>
      <td rowspan="1">出力</td>
      <td><code>to_range</code></td>
      <td>シートへ書き戻す</td>
    </tr>
    <tr>
      <td rowspan="1">状態管理</td>
      <td><code>clear</code></td>
      <td>状態を初期化する</td>
    </tr>
  </tbody>
</table>

## 区分ごとの補足

### 抽出・選択系メソッドの補足

- `filter_rows` `slice_rows` `select_columns` は元の `Matrix` を変更せず、新しい `Matrix` を返します。
- `filter_rows` は 0 件になってもエラーではなく、列数を保った空 `Matrix` を返します。

### 更新・追加系メソッドの補足

- `set_column_values` `append_row` `append_column` は、保持中の表をその場で更新します。
- `append_column` は行数 0 の `Matrix` に対しても、列構造だけを増やせます。

### 変形系メソッドの補足

- `transpose` は元データを変更せず、転置済みの新しい `Matrix` を返します。
- 行数 0 の `Matrix` に対する転置は未サポートです。

### 出力系メソッドの補足

- `to_range` は行数 0 の `Matrix` を出力できません。
- 空表を出したい場合は、`DataTable` でヘッダー付き出力を行う方が向いています。

### 状態管理系メソッドの補足

- `clear` は出力ではなく、内部状態を未読込状態へ戻すためのメソッドです。

## よく使うレシピ

### 行マスクで絞り込んで必要列だけを抜き出す

```vb
Dim mat As New Matrix
Dim filtered As Matrix
Dim picked As Matrix
Dim mask(1 To 4) As Boolean

mat.read_range Sheet1.Range("A2:D5")

mask(1) = True
mask(2) = False
mask(3) = True
mask(4) = False

Set filtered = mat.filter_rows(mask)
Set picked = filtered.select_columns(Array(1, 4))
```

### 既存表に計算済みの列を追加する

```vb
Dim mat As New Matrix

mat.read_range Sheet1.Range("A2:C4")
mat.append_column Array("OK", "NG", "OK")
```

### `DataTable` の内部表現として使える形に整える

```vb
Dim mat As New Matrix
Dim tbl As New DataTable

mat.read_range Sheet1.Range("A2:C6")
tbl.read_matrix mat, Array("date", "item", "amount")
```

## 全メソッド解説

### 読込系

<details>
<summary><code>read_range(ByVal rng As Range)</code></summary>

矩形 `Range` を 2 次元配列として読み込みます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | `rng` が `Nothing` でないこと |
| 入力 | `rng As Range` |
| 出力 | 内部状態 `MATRIX` を更新 |
| 実行内容 | 単一セルなら `1 x 1`、複数セルなら矩形表として保持する |
| 注意点 | 入力の添字は内部で 1 始まりに正規化される |
| 代表ユースケース | シート上の表を一括で読み込む |

入力イメージ:

| A | B | C |
| --- | --- | --- |
| 1 | A | 100 |
| 2 | B | 200 |
| 3 | C | 300 |

```vb
Dim mat As New Matrix
mat.read_range Sheet1.Range("A2:C4")
```

実行後イメージ:

| 確認項目 | 値 |
| --- | --- |
| `mat.row_count` | `3` |
| `mat.col_count` | `3` |
| `mat.item(2, 3)` | `200` |

</details>

<details>
<summary><code>read_array(ByVal arr As Variant)</code></summary>

2 次元配列を読み込みます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | `arr` が配列であり、2 次元配列であること |
| 入力 | `arr As Variant` |
| 出力 | 内部状態 `MATRIX` を更新 |
| 実行内容 | 配列を 1 始まりの 2 次元配列へ正規化して保持する |
| 注意点 | 1 次元や 3 次元以上の配列はエラー |
| 代表ユースケース | 他処理で作った表データを受け取る |

```vb
Dim arr(5 To 6, 3 To 4) As Variant
Dim mat As New Matrix

arr(5, 3) = "A"
arr(5, 4) = 100
arr(6, 3) = "B"
arr(6, 4) = 200

mat.read_array arr
```

</details>

<details>
<summary><code>read_empty(ByVal colCount As Long)</code></summary>

列数だけを持つ空 `Matrix` を作ります。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | `colCount >= 1` |
| 入力 | `colCount As Long` |
| 出力 | 行数 0、列数 `colCount` の空表を生成 |
| 実行内容 | 将来の抽出結果や空表保持用の土台を作る |
| 注意点 | 行データはまだ存在しない |
| 代表ユースケース | 0 件結果を保持したいとき |

```vb
Dim mat As New Matrix

mat.read_empty 3
Debug.Print mat.row_count
Debug.Print mat.col_count
```

</details>

### 参照系

<details>
<summary><code>is_loaded() As Boolean</code></summary>

読込済みかどうかを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | なし |
| 戻り値 | `Boolean` |
| 代表ユースケース | 実行前確認、デバッグ確認 |

</details>

<details>
<summary><code>row_count() As Long</code></summary>

行数を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Long` |
| 代表ユースケース | 件数確認、ループ回数決定 |

</details>

<details>
<summary><code>col_count() As Long</code></summary>

列数を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Long` |
| 代表ユースケース | 列構造確認、列ループ |

</details>

<details>
<summary><code>data() As Variant</code></summary>

内部 2 次元配列のコピーを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Variant` の 2 次元配列 |
| 注意点 | 返却後に配列を変更しても内部状態へは直接反映されない |
| 代表ユースケース | 他モジュールへの受け渡し、デバッグ確認 |

</details>

<details>
<summary><code>item(ByVal rowIndex As Long, ByVal colIndex As Long) As Variant</code></summary>

指定位置の値を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`rowIndex` と `colIndex` が範囲内であること |
| 入力 | `rowIndex As Long`, `colIndex As Long` |
| 戻り値 | `Variant` |
| 代表ユースケース | 単一点の値確認 |

```vb
Dim mat As New Matrix

mat.read_range Sheet1.Range("A2:C4")
Debug.Print mat.item(2, 3)
```

</details>

<details>
<summary><code>row_values(ByVal rowIndex As Long) As Variant</code></summary>

指定行を一次元配列で返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`rowIndex` が範囲内であること |
| 入力 | `rowIndex As Long` |
| 戻り値 | `Variant` の一次元配列 |
| 代表ユースケース | 特定行の再利用、行単位処理 |

</details>

<details>
<summary><code>col_values(ByVal colIndex As Long) As Variant</code></summary>

指定列を一次元配列で返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`colIndex` が範囲内であること、行数が 1 以上であること |
| 入力 | `colIndex As Long` |
| 戻り値 | `Variant` の一次元配列 |
| 注意点 | 行数 0 の `Matrix` ではエラー |
| 代表ユースケース | 条件マスクや更新値配列の作成 |

</details>

### 抽出・選択系

<details>
<summary><code>filter_rows(ByVal mask As Variant) As Matrix</code></summary>

行マスクで必要な行だけを残します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`mask` が一次元配列で行数と一致すること |
| 入力 | `mask As Variant` |
| 戻り値 | 新しい `Matrix` |
| 注意点 | 0 件でも空 `Matrix` を返す |
| 代表ユースケース | 条件抽出、中間表の作成 |

入力イメージ:

| A | B |
| --- | --- |
| A | 100 |
| B | 200 |
| C | 300 |

```vb
Dim mat As New Matrix
Dim filtered As Matrix
Dim mask(1 To 3) As Boolean

mat.read_range Sheet1.Range("A2:B4")
mask(1) = True
mask(2) = False
mask(3) = True

Set filtered = mat.filter_rows(mask)
```

出力イメージ:

| A | B |
| --- | --- |
| A | 100 |
| C | 300 |

</details>

<details>
<summary><code>slice_rows(ByVal rowIndexes As Variant) As Matrix</code></summary>

指定した行番号だけを、指定順で取り出します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`rowIndexes` が一次元配列であること、各行番号が範囲内であること |
| 入力 | `rowIndexes As Variant` |
| 戻り値 | 新しい `Matrix` |
| 注意点 | 空配列相当なら空 `Matrix` を返す |
| 代表ユースケース | 任意行抜き出し、再並び替え |

</details>

<details>
<summary><code>select_columns(ByVal columnIndexes As Variant) As Matrix</code></summary>

指定列だけを選んだ新しい `Matrix` を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`columnIndexes` が一次元配列であること、各列番号が範囲内であること |
| 入力 | `columnIndexes As Variant` |
| 戻り値 | 新しい `Matrix` |
| 注意点 | 列指定 0 件はエラー |
| 代表ユースケース | 必要列だけの中間表を作る |

入力イメージ:

| A | B | C |
| --- | --- | --- |
| 1 | A | 100 |
| 2 | B | 200 |

```vb
Dim mat As New Matrix
Dim picked As Matrix

mat.read_range Sheet1.Range("A2:C3")
Set picked = mat.select_columns(Array(1, 3))
```

出力イメージ:

| A | C |
| --- | --- |
| 1 | 100 |
| 2 | 200 |

</details>

### 更新・追加系

<details>
<summary><code>set_column_values(ByVal colIndex As Long, ByVal values As Variant)</code></summary>

指定列全体を新しい値で差し替えます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`colIndex` が範囲内であること、`values` が一次元配列で行数と一致すること |
| 入力 | `colIndex As Long`, `values As Variant` |
| 出力 | 内部表を更新 |
| 実行内容 | 対象列の全行を上書きする |
| 代表ユースケース | 計算済み列の反映 |

</details>

<details>
<summary><code>append_row(ByVal values As Variant)</code></summary>

末尾に 1 行追加します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`values` が一次元配列で列数と一致すること |
| 入力 | `values As Variant` |
| 出力 | 内部表を更新 |
| 実行内容 | 既存表の末尾に新規行を追加する |
| 代表ユースケース | 追加入力行の付加 |

</details>

<details>
<summary><code>append_column(ByVal values As Variant)</code></summary>

末尾に 1 列追加します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 入力 | `values As Variant` |
| 出力 | 内部表を更新 |
| 実行内容 | 行がある場合は値配列を列として追加し、行が 0 件なら列構造だけを追加する |
| 注意点 | 行がある場合、`values` は一次元配列で行数と一致する必要がある |
| 代表ユースケース | フラグ列や計算列の追加 |

入力イメージ:

| A | B |
| --- | --- |
| A | 10 |
| B | 20 |

```vb
Dim mat As New Matrix

mat.read_range Sheet1.Range("A2:B3")
mat.append_column Array("OK", "NG")
```

出力イメージ:

| A | B | C |
| --- | --- | --- |
| A | 10 | OK |
| B | 20 | NG |

</details>

### 変形系

<details>
<summary><code>transpose() As Matrix</code></summary>

転置した新しい `Matrix` を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、行数 0 でないこと |
| 戻り値 | 新しい `Matrix` |
| 注意点 | 空 `Matrix` の転置は未サポート |
| 代表ユースケース | 行列方向の入れ替え |

```vb
Dim mat As New Matrix
Dim transposed As Matrix

mat.read_range Sheet1.Range("A2:C3")
Set transposed = mat.transpose()
```

</details>

### 出力系

<details>
<summary><code>to_range(ByVal topLeft As Range)</code></summary>

保持している表をシートへ書き戻します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`topLeft` が `Nothing` でないこと、行数 0 でないこと |
| 入力 | `topLeft As Range` |
| 出力 | ワークシート上のセル範囲 |
| 実行内容 | `topLeft` を左上として矩形範囲へ一括書込みする |
| 注意点 | 行数 0 の `Matrix` は出力できない |
| 代表ユースケース | 中間表や最終結果の確認 |

```vb
Dim mat As New Matrix

mat.read_range Sheet1.Range("A2:C4")
mat.to_range Sheet1.Range("F2")
```

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
| 実行内容 | 配列、行数、列数、読込状態を初期化する |
| 注意点 | 保持中データは失われる |
| 代表ユースケース | 同じインスタンスを再利用する前のリセット |

```vb
Dim mat As New Matrix

mat.read_range Sheet1.Range("A2:C4")
mat.clear

Debug.Print mat.is_loaded
```

</details>
