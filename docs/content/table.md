# クラスモジュール `Table`

`Table` は、列名を持つ表データを扱うための VBA クラスモジュールです。  
内部では `Matrix` を使いながら、列名ベースの抽出、更新、並べ替え、出力を行います。

## まず何ができるか

- ヘッダー付きの表を、そのまま列名つきデータとして読み込む
- 列名を使って条件抽出する
- 条件に合う行だけ別列を更新する
- 必要な列だけを選んで新しい `Table` を作る
- 列追加、列名変更、並べ替えで表を整える
- ヘッダー付きでシートへ書き戻す

## クイックスタート

```vb
Sub Sample_Table_QuickStart()
    Dim tbl As New Table
    Dim okRows As Table

    tbl.read_range Sheet1.Range("A1:D8"), hasHeader:=True
    tbl.set_by_equals "status", "NG", "score", 0
    Set okRows = tbl.filter_by_equals("status", "OK")

    okRows.to_range Sheet1.Range("G1"), includeHeader:=True
End Sub
```

## 利用前提

| 項目 | 内容 |
| --- | --- |
| 初期化必須 | `read_range` または `read_matrix` を先に実行 |
| 列名制約 | 空文字不可、重複不可 |
| 条件配列 | 行数と一致する必要がある |
| 未初期化時 | 参照系・抽出選択系・更新整形系・出力系はエラー |
| 空テーブル | データ行 0 件でも列構造は保持できる |

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
      <td rowspan="2">読込</td>
      <td><code>read_range</code></td>
      <td>シート上の表を読む</td>
    </tr>
    <tr>
      <td><code>read_matrix</code></td>
      <td><code>Matrix</code> に列名を付けて読む</td>
    </tr>
    <tr>
      <td rowspan="7">参照</td>
      <td><code>is_loaded</code></td>
      <td>読込済みかを返す</td>
    </tr>
    <tr>
      <td><code>row_count</code></td>
      <td>データ行数を返す</td>
    </tr>
    <tr>
      <td><code>col_count</code></td>
      <td>列数を返す</td>
    </tr>
    <tr>
      <td><code>column_names</code></td>
      <td>列名一覧を返す</td>
    </tr>
    <tr>
      <td><code>matrix</code></td>
      <td>内部データを <code>Matrix</code> として返す</td>
    </tr>
    <tr>
      <td><code>col</code></td>
      <td>指定列を一次元配列で返す</td>
    </tr>
    <tr>
      <td><code>col_vector</code></td>
      <td>指定列を <code>Vector</code> として返す</td>
    </tr>
    <tr>
      <td rowspan="7">抽出・選択</td>
      <td><code>filter_by_mask</code></td>
      <td>条件マスクで行を絞り込む</td>
    </tr>
    <tr>
      <td><code>filter_by_equals</code></td>
      <td>一致条件で抽出する</td>
    </tr>
    <tr>
      <td><code>filter_by_in</code></td>
      <td>複数候補のいずれかに一致する行を抽出する</td>
    </tr>
    <tr>
      <td><code>filter_by_contains</code></td>
      <td>部分一致で抽出する</td>
    </tr>
    <tr>
      <td><code>filter_by_all_equals</code></td>
      <td>複数条件を AND で抽出する</td>
    </tr>
    <tr>
      <td><code>filter_by_any_equals</code></td>
      <td>複数条件を OR で抽出する</td>
    </tr>
    <tr>
      <td><code>select_columns</code></td>
      <td>必要列だけを選ぶ</td>
    </tr>
    <tr>
      <td rowspan="6">更新・整形</td>
      <td><code>add_column</code></td>
      <td>新しい列を追加する</td>
    </tr>
    <tr>
      <td><code>rename_column</code></td>
      <td>列名を変更する</td>
    </tr>
    <tr>
      <td><code>sort_by</code></td>
      <td>指定列で並べ替える</td>
    </tr>
    <tr>
      <td><code>set_by_mask</code></td>
      <td>条件に合う行だけ値を更新する</td>
    </tr>
    <tr>
      <td><code>set_by_equals</code></td>
      <td>一致条件で別列を更新する</td>
    </tr>
    <tr>
      <td><code>set_column</code></td>
      <td>列全体を差し替える</td>
    </tr>
    <tr>
      <td rowspan="1">出力</td>
      <td><code>to_range</code></td>
      <td>ヘッダー付きで書き戻す</td>
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

- `filter_by_*` と `select_columns` は元の `Table` を変更せず、新しい `Table` を返します。
- `filter_by_mask` に渡す配列は、`Vector.eq` や `Vector.gt` で作ったブール配列とも組み合わせられます。

### 更新・整形系メソッドの補足

- `add_column` `rename_column` `sort_by` `set_by_*` `set_column` は内部の表をその場で更新します。
- `sort_by` は比較不能値を無理に並べ替えず、その行をそのまま残す方針です。

### 出力系メソッドの補足

- `to_range` は空テーブルでも `includeHeader=True` ならヘッダーだけを出力できます。
- データだけを出したい場合は `includeHeader:=False` を指定します。

### 状態管理系メソッドの補足

- `clear` はシート出力ではなく、列名、`Matrix`、列名マップを含む内部状態を未読込へ戻すためのメソッドです。

## よく使うレシピ

### `status="OK"` の行だけを抽出する

```vb
Dim tbl As New Table
Dim okRows As Table

tbl.read_range Sheet1.Range("A1:D20"), hasHeader:=True
Set okRows = tbl.filter_by_equals("status", "OK")
```

### `status="NG"` の行だけ `score` を 0 にする

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:D20"), hasHeader:=True
tbl.set_by_equals "status", "NG", "score", 0
```

### `Vector` を使って列条件を作り、`Table` へ戻す

```vb
Dim tbl As New Table
Dim amountVec As Vector
Dim mask As Variant
Dim highRows As Table

tbl.read_range Sheet1.Range("A1:D20"), hasHeader:=True
Set amountVec = tbl.col_vector("amount")
mask = amountVec.gt(1000)

Set highRows = tbl.filter_by_mask(mask)
```

## 全メソッド解説

### 読込系

<details>
<summary><code>read_range(ByVal rng As Range, Optional ByVal hasHeader As Boolean = True)</code></summary>

シート上の表を読み込みます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | `rng` が `Nothing` でないこと |
| 入力 | `rng As Range`, `hasHeader As Boolean` |
| 出力 | 内部 `Matrix` と列名定義を更新 |
| 実行内容 | `hasHeader=True` なら 1 行目を列名、残りをデータとして保持する |
| 注意点 | `hasHeader=True` かつ 1 行だけの場合はヘッダーのみの空テーブルになる |
| 代表ユースケース | Excel のヘッダー付き表をそのまま扱う |

入力イメージ:

| name | status | score |
| --- | --- | --- |
| A | OK | 80 |
| B | NG | 70 |
| C | OK | 90 |

```vb
Dim tbl As New Table
tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
```

実行後イメージ:

| 確認項目 | 値 |
| --- | --- |
| `tbl.row_count` | `3` |
| `tbl.col_count` | `3` |
| `tbl.column_names` | `["name", "status", "score"]` |

</details>

<details>
<summary><code>read_matrix(ByVal src As Matrix, ByVal columnNames As Variant)</code></summary>

`Matrix` と列名配列から `Table` を構築します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | `src` が `Nothing` でないこと、`columnNames` が一次元配列であること、列名数と列数が一致すること |
| 入力 | `src As Matrix`, `columnNames As Variant` |
| 出力 | 内部 `Matrix` と列名マップを更新 |
| 実行内容 | `Matrix` の内容を保持しつつ、列名付きテーブルへ変換する |
| 代表ユースケース | `Matrix` ベースの中間結果に列名を付けたいとき |

```vb
Dim mat As New Matrix
Dim tbl As New Table

mat.read_range Sheet1.Range("A2:C5")
tbl.read_matrix mat, Array("date", "item", "amount")
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

```vb
Dim tbl As New Table

Debug.Print tbl.is_loaded

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
Debug.Print tbl.is_loaded
```

</details>

<details>
<summary><code>row_count() As Long</code></summary>

データ行数を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Long` |
| 代表ユースケース | 件数確認、条件配列長の確認 |

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
Debug.Print tbl.row_count
```

</details>

<details>
<summary><code>col_count() As Long</code></summary>

列数を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Long` |
| 代表ユースケース | 列構造確認、出力列数確認 |

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
Debug.Print tbl.col_count
```

</details>

<details>
<summary><code>column_names() As Variant</code></summary>

列名一覧のコピーを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Variant` の一次元配列 |
| 注意点 | 返却後に配列を変更しても内部状態へは直接反映されない |
| 代表ユースケース | UI 表示、出力列確認 |

```vb
Dim tbl As New Table
Dim names As Variant

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
names = tbl.column_names
Debug.Print names(1)
```

</details>

<details>
<summary><code>matrix() As Matrix</code></summary>

内部データを `Matrix` として返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること |
| 戻り値 | `Matrix` |
| 注意点 | 返却される `Matrix` は別インスタンス |
| 代表ユースケース | 低レベルな行列処理へ渡す |

```vb
Dim tbl As New Table
Dim mat As Matrix

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
Set mat = tbl.matrix
Debug.Print mat.row_count
```

</details>

<details>
<summary><code>col(ByVal columnName As String) As Variant</code></summary>

指定列を一次元配列で返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、列名が存在すること、データ行が 1 件以上あること |
| 入力 | `columnName As String` |
| 戻り値 | `Variant` の一次元配列 |
| 注意点 | 空テーブルではエラー |
| 代表ユースケース | 条件配列や集計の元データ取得 |

```vb
Dim tbl As New Table
Dim scores As Variant

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
scores = tbl.col("score")
```

</details>

<details>
<summary><code>col_vector(ByVal columnName As String) As Vector</code></summary>

指定列を `Vector` として返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | `col` と同じ |
| 入力 | `columnName As String` |
| 戻り値 | `Vector` |
| 代表ユースケース | `Vector` の `eq` `gt` `sum` `mean` を使いたいとき |

```vb
Dim tbl As New Table
Dim scoreVec As Vector

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
Set scoreVec = tbl.col_vector("score")
Debug.Print scoreVec.mean
```

</details>

### 抽出・選択系

<details>
<summary><code>filter_by_mask(ByVal mask As Variant) As Table</code></summary>

条件マスクで行を絞り込みます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`mask` が一次元配列で行数と一致すること |
| 入力 | `mask As Variant` |
| 戻り値 | 新しい `Table` |
| 注意点 | 空テーブルでは空テーブルを返す |
| 代表ユースケース | 既に作成済みの条件配列で抽出する |

```vb
Dim tbl As New Table
Dim filtered As Table
Dim mask(1 To 3) As Boolean

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
mask(1) = True
mask(2) = False
mask(3) = True

Set filtered = tbl.filter_by_mask(mask)
```

</details>

<details>
<summary><code>filter_by_equals(ByVal columnName As String, ByVal matchValue As Variant) As Table</code></summary>

指定列が特定値と一致する行だけを返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、列名が存在すること |
| 入力 | `columnName As String`, `matchValue As Variant` |
| 戻り値 | 新しい `Table` |
| 注意点 | `Null` `Error` 比較不能値は不一致扱い |
| 代表ユースケース | ステータス一致、カテゴリ一致の抽出 |

入力イメージ:

| name | status | score |
| --- | --- | --- |
| A | OK | 80 |
| B | NG | 70 |
| C | OK | 90 |

```vb
Dim tbl As New Table
Dim okRows As Table

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
Set okRows = tbl.filter_by_equals("status", "OK")
```

出力イメージ:

| name | status | score |
| --- | --- | --- |
| A | OK | 80 |
| C | OK | 90 |

</details>

<details>
<summary><code>filter_by_in(ByVal columnName As String, ByVal matchValues As Variant) As Table</code></summary>

指定列が候補配列のいずれかに一致する行を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、列名が存在すること、`matchValues` が一次元配列であること |
| 入力 | `columnName As String`, `matchValues As Variant` |
| 戻り値 | 新しい `Table` |
| 代表ユースケース | 複数カテゴリ一括抽出 |

```vb
Dim tbl As New Table
Dim selectedRows As Table

tbl.read_range Sheet1.Range("A1:C6"), hasHeader:=True
Set selectedRows = tbl.filter_by_in("status", Array("OK", "PENDING"))
```

</details>

<details>
<summary><code>filter_by_contains(ByVal columnName As String, ByVal searchText As String, Optional ByVal caseSensitive As Boolean = False) As Table</code></summary>

指定列の文字列に検索文字列を含む行を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、列名が存在すること |
| 入力 | `columnName As String`, `searchText As String`, `caseSensitive As Boolean` |
| 戻り値 | 新しい `Table` |
| 注意点 | `Null` `Empty` `Error` は不一致扱い |
| 代表ユースケース | 部分一致検索、キーワード抽出 |

```vb
Dim tbl As New Table
Dim hitRows As Table

tbl.read_range Sheet1.Range("A1:C5"), hasHeader:=True
Set hitRows = tbl.filter_by_contains("product_name", "コーヒー")
```

</details>

<details>
<summary><code>filter_by_all_equals(ByVal columnNames As Variant, ByVal matchValues As Variant) As Table</code></summary>

複数条件を AND で結合して抽出します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`columnNames` と `matchValues` が一次元配列で長さ一致すること、各列名が存在すること |
| 入力 | `columnNames As Variant`, `matchValues As Variant` |
| 戻り値 | 新しい `Table` |
| 代表ユースケース | 複合条件の厳密抽出 |

```vb
Dim tbl As New Table
Dim narrowed As Table

tbl.read_range Sheet1.Range("A1:D8"), hasHeader:=True
Set narrowed = tbl.filter_by_all_equals(Array("status", "category"), Array("OK", "A"))
```

</details>

<details>
<summary><code>filter_by_any_equals(ByVal columnNames As Variant, ByVal matchValues As Variant) As Table</code></summary>

複数条件を OR で結合して抽出します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | `filter_by_all_equals` と同じ |
| 入力 | `columnNames As Variant`, `matchValues As Variant` |
| 戻り値 | 新しい `Table` |
| 代表ユースケース | いずれか条件に合う行の抽出 |

```vb
Dim tbl As New Table
Dim matched As Table

tbl.read_range Sheet1.Range("A1:D8"), hasHeader:=True
Set matched = tbl.filter_by_any_equals(Array("status", "category"), Array("NG", "B"))
```

</details>

<details>
<summary><code>select_columns(ByVal columnNames As Variant) As Table</code></summary>

必要な列だけを残した新しい `Table` を返します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`columnNames` が一次元配列であること、各列名が存在すること |
| 入力 | `columnNames As Variant` |
| 戻り値 | 新しい `Table` |
| 代表ユースケース | レポート用の列絞り込み |

入力イメージ:

| date | name | status | score |
| --- | --- | --- | --- |
| 4/1 | A | OK | 80 |
| 4/2 | B | NG | 70 |

```vb
Dim tbl As New Table
Dim reportTbl As Table

tbl.read_range Sheet1.Range("A1:D3"), hasHeader:=True
Set reportTbl = tbl.select_columns(Array("date", "score"))
```

出力イメージ:

| date | score |
| --- | --- |
| 4/1 | 80 |
| 4/2 | 70 |

</details>

### 更新・整形系

<details>
<summary><code>add_column(ByVal columnName As String, Optional ByVal values As Variant)</code></summary>

新しい列を末尾に追加します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`columnName` が空でなく重複しないこと |
| 入力 | `columnName As String`, `values As Variant` |
| 出力 | 内部テーブルを更新 |
| 実行内容 | 新しい列を追加し、列名マップを再構築する |
| 注意点 | `values` 省略時、データ行がある場合は `Empty` 列を追加する |
| 代表ユースケース | 計算列、フラグ列の追加 |

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
tbl.add_column "flag"
```

</details>

<details>
<summary><code>rename_column(ByVal oldName As String, ByVal newName As String)</code></summary>

列名を変更します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`oldName` が存在すること、`newName` が空でなく重複しないこと |
| 入力 | `oldName As String`, `newName As String` |
| 出力 | 内部列名定義を更新 |
| 実行内容 | 列名を差し替え、列名マップを再構築する |
| 代表ユースケース | 業務用ラベルへの変更、列名正規化 |

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
tbl.rename_column "score", "point"
```

</details>

<details>
<summary><code>sort_by(ByVal columnName As String, Optional ByVal ascending As Boolean = True)</code></summary>

指定列で並べ替えます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、列名が存在すること |
| 入力 | `columnName As String`, `ascending As Boolean` |
| 出力 | 内部テーブルを更新 |
| 実行内容 | 行単位で並べ替えを行う |
| 注意点 | 比較不能値は交換対象にしない |
| 代表ユースケース | 日付順、金額順、スコア順の並べ替え |

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:C5"), hasHeader:=True
tbl.sort_by "score", ascending:=False
```

</details>

<details>
<summary><code>set_by_mask(ByVal mask As Variant, ByVal columnName As String, ByVal newValue As Variant)</code></summary>

条件に合う行だけ、指定列の値を更新します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`mask` が一次元配列で行数と一致すること、列名が存在すること |
| 入力 | `mask As Variant`, `columnName As String`, `newValue As Variant` |
| 出力 | 内部テーブルを更新 |
| 実行内容 | `True` 行のみ対象列を `newValue` へ置換する |
| 代表ユースケース | 条件一致行だけの補正、フラグ更新 |

```vb
Dim tbl As New Table
Dim mask(1 To 3) As Boolean

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
mask(1) = False
mask(2) = True
mask(3) = True

tbl.set_by_mask mask, "score", 0
```

</details>

<details>
<summary><code>set_by_equals(ByVal conditionColumnName As String, ByVal matchValue As Variant, ByVal targetColumnName As String, ByVal newValue As Variant)</code></summary>

条件列が一致した行だけ、別列の値を更新します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、両列名が存在すること |
| 入力 | `conditionColumnName As String`, `matchValue As Variant`, `targetColumnName As String`, `newValue As Variant` |
| 出力 | 内部テーブルを更新 |
| 実行内容 | `conditionColumnName = matchValue` の行に対して `targetColumnName` を上書きする |
| 代表ユースケース | `status="NG"` 行だけ `score=0` にする処理 |

入力イメージ:

| name | status | score |
| --- | --- | --- |
| A | OK | 80 |
| B | NG | 70 |
| C | OK | 90 |

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
tbl.set_by_equals "status", "NG", "score", 0
```

出力イメージ:

| name | status | score |
| --- | --- | --- |
| A | OK | 80 |
| B | NG | 0 |
| C | OK | 90 |

</details>

<details>
<summary><code>set_column(ByVal columnName As String, ByVal values As Variant)</code></summary>

指定列全体を差し替えます。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、列名が存在すること、`values` が一次元配列で行数と一致すること |
| 入力 | `columnName As String`, `values As Variant` |
| 出力 | 内部テーブルを更新 |
| 実行内容 | 対象列を新しい値配列で全面更新する |
| 代表ユースケース | 計算済み列の一括反映 |

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
tbl.set_column "score", Array(100, 90, 80)
```

</details>

### 出力系

<details>
<summary><code>to_range(ByVal topLeft As Range, Optional ByVal includeHeader As Boolean = True)</code></summary>

表をシートへ出力します。

| 項目 | 内容 |
| --- | --- |
| 前提条件 | 読込済みであること、`topLeft` が `Nothing` でないこと |
| 入力 | `topLeft As Range`, `includeHeader As Boolean` |
| 出力 | ワークシート上のセル範囲 |
| 実行内容 | `includeHeader=True` ならヘッダー行を含めて出力する |
| 注意点 | 空テーブルでは `includeHeader=True` の場合のみヘッダーを出力する |
| 代表ユースケース | 抽出結果や更新結果をシートへ戻す |

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
tbl.to_range Sheet1.Range("G1"), includeHeader:=True
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
| 実行内容 | `Matrix`、列名、列名マップ、読込状態を初期化する |
| 注意点 | 保持中データは失われる |
| 代表ユースケース | 同じインスタンスを再利用する前のリセット |

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:C4"), hasHeader:=True
tbl.clear

Debug.Print tbl.is_loaded
```

</details>
