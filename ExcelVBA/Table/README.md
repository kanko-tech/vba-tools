# クラスモジュール `Table`

`Table` は、列名を持つ表データを扱うためのクラスモジュールです。  
内部では `Matrix` を利用しつつ、列名ベースの抽出・更新・並べ替えを提供します。

## 概要

| 項目 | 内容 |
| --- | --- |
| 主な用途 | ヘッダー付き表の読込、列名ベース操作 |
| 内部データ | `Matrix` + 列名配列 |
| 列名制約 | 空文字不可、重複不可 |
| 空テーブル | データ行 0 件でも列構造は保持可能 |
| 関連モジュール | 行列基盤は `Matrix`、列単位処理は `Vector` |

## 利用前提

| 項目 | 内容 |
| --- | --- |
| 初期化必須 | `read_range` または `read_matrix` を先に実行 |
| 条件配列 | 行数と一致する必要がある |
| 列名参照 | 存在しない列名はエラー |
| 出力 | 空テーブルでもヘッダーのみ出力可能 |

## 典型的なユースケース

- `status = "OK"` の行だけ抽出する
- 条件一致行の `score` を更新する
- 列名で必要列だけを選ぶ
- 新しい計算列を追加し、列名を付けて管理する

## メソッド一覧

| 区分 | メソッド |
| --- | --- |
| 読込 | `read_range` `read_matrix` |
| 参照 | `is_loaded` `row_count` `col_count` `column_names` `matrix` `col` `col_vector` |
| 抽出 | `filter_by_mask` `filter_by_equals` `filter_by_in` `filter_by_contains` `filter_by_all_equals` `filter_by_any_equals` |
| 更新・構造変更 | `select_columns` `add_column` `rename_column` `sort_by` `set_by_mask` `set_by_equals` `set_column` |
| 出力 | `to_range` `clear` |

## 読込系

### `read_range(ByVal rng As Range, Optional ByVal hasHeader As Boolean = True)`

| 項目 | 内容 |
| --- | --- |
| 役割 | シート上の表を読み込む |
| 前提条件 | `rng` が `Nothing` でないこと |
| 入力 | `rng As Range` `hasHeader As Boolean` |
| 実行内容 | `hasHeader=True` なら 1 行目を列名、残りをデータとして読込。`False` なら `col1` `col2` ... を自動採番 |
| 出力 | 内部 `Matrix` と列名定義を更新 |
| 注意点 | `hasHeader=True` かつ 1 行だけの場合はヘッダーのみの空テーブルになる |
| ユースケース | Excel の表をそのまま操作対象にしたいとき |

ユースケース例: ヘッダー付き売上表をそのまま読み込み、条件更新を行う。

```vb
Dim tbl As New Table
tbl.read_range Sheet1.Range("A1:D20"), hasHeader:=True
```

### `read_matrix(ByVal src As Matrix, ByVal columnNames As Variant)`

| 項目 | 内容 |
| --- | --- |
| 役割 | `Matrix` と列名配列から `Table` を構築する |
| 前提条件 | `src` が `Nothing` でないこと、`columnNames` が一次元配列であること、列名数と列数が一致すること |
| 入力 | `src As Matrix` `columnNames As Variant` |
| 実行内容 | 内部 `Matrix` と列名マップを構築 |
| 出力 | `Table` の内部状態を更新 |
| ユースケース | 中間 `Matrix` 結果に列名を付けて高水準操作へ渡す |

## 参照系

### `is_loaded() As Boolean`

| 項目 | 内容 |
| --- | --- |
| 役割 | 読込済みかを返す |
| 戻り値 | `Boolean` |
| ユースケース | 実行前確認、デバッグ用途 |

### `row_count() As Long`

| 項目 | 内容 |
| --- | --- |
| 役割 | データ行数を返す |
| 戻り値 | `Long` |
| ユースケース | 条件配列長の確認、件数表示 |

### `col_count() As Long`

| 項目 | 内容 |
| --- | --- |
| 役割 | 列数を返す |
| 戻り値 | `Long` |
| ユースケース | 構造確認、列ループ |

### `column_names() As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 列名一覧のコピーを返す |
| 前提条件 | 読込済みであること |
| 戻り値 | `Variant` の一次元配列 |
| ユースケース | UI 表示、出力列確認 |

### `matrix() As Matrix`

| 項目 | 内容 |
| --- | --- |
| 役割 | 内部データを `Matrix` として返す |
| 前提条件 | 読込済みであること |
| 戻り値 | `Matrix` |
| ユースケース | 低レベルな行列処理へ渡す |

### `col(ByVal columnName As String) As Variant`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定列を一次元配列で返す |
| 前提条件 | 読込済みであること、列名が存在すること、データ行が 1 件以上あること |
| 入力 | `columnName As String` |
| 戻り値 | `Variant` の一次元配列 |
| 注意点 | 空テーブルではエラー |
| ユースケース | 条件配列の生成、列集計の元データ取得 |

### `col_vector(ByVal columnName As String) As Vector`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定列を `Vector` として返す |
| 前提条件 | `col` と同じ |
| 入力 | `columnName As String` |
| 戻り値 | `Vector` |
| ユースケース | `eq` `gt` `sum` `mean` など `Vector` API を使う |

ユースケース例: `Table` から列を取り出して `Vector` の集計を使う。

```vb
Dim tbl As New Table
Dim scoreVec As Vector

tbl.read_range Sheet1.Range("A1:D10"), hasHeader:=True
Set scoreVec = tbl.col_vector("score")
Debug.Print scoreVec.mean
```

## 抽出系

### `filter_by_mask(ByVal mask As Variant) As Table`

| 項目 | 内容 |
| --- | --- |
| 役割 | 条件マスクで行を絞り込む |
| 前提条件 | 読込済みであること、`mask` が一次元配列で行数と一致すること |
| 入力 | `mask As Variant` |
| 戻り値 | 新しい `Table` |
| 注意点 | 空テーブルでは空テーブルを返す |
| ユースケース | 既に作った条件配列で抽出したいとき |

### `filter_by_equals(ByVal columnName As String, ByVal matchValue As Variant) As Table`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定列が特定値と一致する行だけを返す |
| 前提条件 | 読込済みであること、列名が存在すること |
| 入力 | `columnName As String` `matchValue As Variant` |
| 戻り値 | 新しい `Table` |
| 注意点 | `Null` / `Error` / 比較不能値は一致扱いしない |
| ユースケース | ステータス一致、カテゴリ一致抽出 |

ユースケース例: `status="OK"` の行だけ抜き出す。

```vb
Dim tbl As New Table
Dim okRows As Table

tbl.read_range Sheet1.Range("A1:D20"), hasHeader:=True
Set okRows = tbl.filter_by_equals("status", "OK")
```

### `filter_by_in(ByVal columnName As String, ByVal matchValues As Variant) As Table`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定列が候補配列のいずれかに一致する行を返す |
| 前提条件 | 読込済みであること、列名が存在すること、`matchValues` が一次元配列であること |
| 入力 | `columnName As String` `matchValues As Variant` |
| 戻り値 | 新しい `Table` |
| ユースケース | 複数カテゴリ一括抽出 |

### `filter_by_contains(ByVal columnName As String, ByVal searchText As String, Optional ByVal caseSensitive As Boolean = False) As Table`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定列の文字列に検索文字列を含む行を返す |
| 前提条件 | 読込済みであること、列名が存在すること |
| 入力 | `columnName As String` `searchText As String` `caseSensitive As Boolean` |
| 戻り値 | 新しい `Table` |
| 注意点 | `Null` / `Empty` / `Error` は不一致扱い |
| ユースケース | 部分一致検索、キーワード抽出 |

ユースケース例: 商品名列に `"コーヒー"` を含む行だけを抽出する。

```vb
Dim tbl As New Table
Dim hitRows As Table

tbl.read_range Sheet1.Range("A1:D20"), hasHeader:=True
Set hitRows = tbl.filter_by_contains("product_name", "コーヒー")
```

### `filter_by_all_equals(ByVal columnNames As Variant, ByVal matchValues As Variant) As Table`

| 項目 | 内容 |
| --- | --- |
| 役割 | 複数列条件を AND で結合して抽出する |
| 前提条件 | 読込済みであること、両引数が一次元配列で長さ一致すること、各列名が存在すること |
| 入力 | `columnNames As Variant` `matchValues As Variant` |
| 戻り値 | 新しい `Table` |
| ユースケース | 複合条件の厳密抽出 |

### `filter_by_any_equals(ByVal columnNames As Variant, ByVal matchValues As Variant) As Table`

| 項目 | 内容 |
| --- | --- |
| 役割 | 複数列条件を OR で結合して抽出する |
| 前提条件 | `filter_by_all_equals` と同じ |
| 入力 | `columnNames As Variant` `matchValues As Variant` |
| 戻り値 | 新しい `Table` |
| ユースケース | いずれか条件に合う行の抽出 |

## 更新・構造変更系

### `select_columns(ByVal columnNames As Variant) As Table`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定列だけを残した `Table` を返す |
| 前提条件 | 読込済みであること、`columnNames` が一次元配列であること、各列名が存在すること |
| 入力 | `columnNames As Variant` |
| 戻り値 | 新しい `Table` |
| ユースケース | 出力列の絞込み、派生表の作成 |

ユースケース例: 必要な列だけを抜き出してレポート用テーブルを作る。

```vb
Dim tbl As New Table
Dim reportTbl As Table

tbl.read_range Sheet1.Range("A1:E20"), hasHeader:=True
Set reportTbl = tbl.select_columns(Array("date", "name", "score"))
```

### `add_column(ByVal columnName As String, Optional ByVal values As Variant)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 新しい列を末尾に追加する |
| 前提条件 | 読込済みであること、`columnName` が空でなく重複しないこと |
| 入力 | `columnName As String` `values As Variant` |
| 実行内容 | 列追加し、列名マップを再構築 |
| 出力 | 内部テーブルを更新 |
| 注意点 | `values` を省略した場合、データ行があるときは `Empty` 列を追加する |
| ユースケース | 計算列、フラグ列の追加 |

ユースケース例: 集計前に判定用フラグ列を追加する。

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:D10"), hasHeader:=True
tbl.add_column "flag"
```

### `rename_column(ByVal oldName As String, ByVal newName As String)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 列名を変更する |
| 前提条件 | 読込済みであること、`oldName` が存在すること、`newName` が空でなく重複しないこと |
| 入力 | `oldName As String` `newName As String` |
| 実行内容 | 指定列名を更新し、列名マップを再構築 |
| 出力 | 内部列名定義を更新 |
| ユースケース | 業務用ラベルへの置換、列名正規化 |

### `sort_by(ByVal columnName As String, Optional ByVal ascending As Boolean = True)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 指定列で並べ替える |
| 前提条件 | 読込済みであること、列名が存在すること |
| 入力 | `columnName As String` `ascending As Boolean` |
| 実行内容 | 行単位で並べ替えを実行 |
| 出力 | 内部テーブルを更新 |
| 注意点 | 比較不能値は交換対象にしない |
| ユースケース | 日付順、金額順、コード順の並べ替え |

ユースケース例: 得点列の降順で並べ替える。

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:D10"), hasHeader:=True
tbl.sort_by "score", ascending:=False
```

### `set_by_mask(ByVal mask As Variant, ByVal columnName As String, ByVal newValue As Variant)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 条件に合う行だけ対象列を更新する |
| 前提条件 | 読込済みであること、`mask` が一次元配列で行数と一致すること、列名が存在すること |
| 入力 | `mask As Variant` `columnName As String` `newValue As Variant` |
| 実行内容 | `True` 行のみ対象列を `newValue` で置換 |
| 出力 | 内部テーブルを更新 |
| ユースケース | 条件一致行のフラグ更新、値補正 |

### `set_by_equals(ByVal conditionColumnName As String, ByVal matchValue As Variant, ByVal targetColumnName As String, ByVal newValue As Variant)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 条件列一致時だけ別列を書き換える |
| 前提条件 | 読込済みであること、両列名が存在すること |
| 入力 | `conditionColumnName As String` `matchValue As Variant` `targetColumnName As String` `newValue As Variant` |
| 実行内容 | `conditionColumnName = matchValue` の行に対して `targetColumnName` を更新 |
| 出力 | 内部テーブルを更新 |
| ユースケース | `status="NG"` 行だけ `score=0` にする処理 |

ユースケース例: `status` が `NG` の行だけ `score` を 0 にする。

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:D20"), hasHeader:=True
tbl.set_by_equals "status", "NG", "score", 0
```

### `set_column(ByVal columnName As String, ByVal values As Variant)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 列全体を差し替える |
| 前提条件 | 読込済みであること、列名が存在すること、`values` が一次元配列で行数と一致すること |
| 入力 | `columnName As String` `values As Variant` |
| 実行内容 | 対象列を新しい値で全面更新 |
| 出力 | 内部テーブルを更新 |
| ユースケース | 計算済み列の一括反映 |

## 出力系

### `to_range(ByVal topLeft As Range, Optional ByVal includeHeader As Boolean = True)`

| 項目 | 内容 |
| --- | --- |
| 役割 | 表をシートへ出力する |
| 前提条件 | 読込済みであること、`topLeft` が `Nothing` でないこと |
| 入力 | `topLeft As Range` `includeHeader As Boolean` |
| 実行内容 | `includeHeader=True` ならヘッダー行を含めて出力 |
| 出力 | ワークシート上のセル範囲 |
| 注意点 | 空テーブルでは `includeHeader=True` の場合のみヘッダーを出力 |
| ユースケース | 抽出結果、更新結果、最終表の書戻し |

ユースケース例: 抽出・更新後の表を別領域へ出力する。

```vb
Dim tbl As New Table

tbl.read_range Sheet1.Range("A1:D20"), hasHeader:=True
tbl.to_range Sheet1.Range("G1"), includeHeader:=True
```

### `clear()`

| 項目 | 内容 |
| --- | --- |
| 役割 | 内部状態を初期化する |
| 実行内容 | `Matrix`、列名、列名マップ、読込状態を初期化 |
| ユースケース | 再利用前のリセット |

## 補足

- 高水準なテーブル操作は `Table`、列単位処理は `Vector`、行列基盤は `Matrix` が担当します
- 空テーブルでも列構造は保持されます
- GitHub 上では各メソッド見出しと表をセットで見ると仕様を追いやすくなります
