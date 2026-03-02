# ExcelVBA 拡張計画メモ（1次元配列→2次元配列→DataFrame）

## 命名案（OneDim / TwoDim / DataFrame の改善）

現状名は意味が通る一方で、責務（ベクトル/行列/表）を名前から直感しづらい。以下の命名を推奨。

- 1次元配列クラス: `Vector1D`（代替: `Series`）
- 2次元配列クラス: `Matrix2D`（代替: `Grid2D`）
- データフレーム風クラス: `TableFrame`（代替: `DataTable`）

推奨セット（統一案）:

- `Vector1D` / `Matrix2D` / `TableFrame`

理由:

- `OneDim` / `TwoDim` よりもデータ構造の意味が明確。
- `DataFrame` は他言語ライブラリ名と同一で期待値が高くなりやすいため、VBA向け軽量実装では `TableFrame` の方がスコープを表現しやすい。
- 将来 `Series` 的な列演算を追加する場合でも、`Vector1D` を列基盤として再利用しやすい。

移行方針（互換性維持）:

1. 先に新クラス名で新規作成（`Vector1D.cls` など）。
2. 既存 `OneDim.cls` は当面残し、READMEで「移行先」を案内。
3. 互換期間後に `OneDim` を非推奨化（段階的廃止）。

## 現状の棚卸し

- 現在の主要コンポーネントは `OneDim.cls` のみ。
- `OneDim` は以下の責務を持つ。
  - 読み込み（列Range、行Range、1次元配列）
  - 型変換（Double/String/Date）
  - 出力（縦・横）
- API は「読み込み済みフラグ（`IsLoaded`）」を前提にしており、状態管理つきクラスとして整理されている。

## 次フェーズの設計方針

### 1. `TwoDim` クラスを新設

想定責務:

- `Read_Range` / `Read_Array`（2次元配列のみ受理）
- `RowCount` / `ColCount` / `Item(row, col)`
- `GetRow(index)` は `OneDim` を返す
- `GetCol(index)` は `OneDim` を返す
- `ToRange(topLeft)`
- `TransposeInPlace`（必要なら）

設計メモ:

- Excel の Range.Value は 1-based 2次元配列として扱う。
- 内部も 1-based を基本にすると Range との相互変換が単純。
- 型変換系は `OneDim` と同等の命名を揃え、利用者の学習コストを下げる。

### 2. DataFrame 風クラス（仮: `DataFrame.cls`）

ユーザー案（連想配列: キー=列名, 値=1次元配列）は VBA で実装しやすく、初期版として妥当。

推奨内部構造（初期版）:

- `mColumns As Scripting.Dictionary`（Key: String, Item: OneDim）
- `mColOrder As Collection`（列順維持）
- `mRowCount As Long`

この構造のメリット:

- 列追加・列参照が高速（辞書アクセス）
- 列順を保持できる（Collection 併用）
- 列ごとの型変換を `OneDim` に委譲可能

最低限 API 案:

- `AddColumn(colName As String, col As OneDim)`
- `Column(colName As String) As OneDim`
- `Columns() As Variant`（列名配列）
- `RowCount() As Long` / `ColCount() As Long`
- `FromRange(rng As Range, hasHeader As Boolean)`
- `ToRange(topLeft As Range, includeHeader As Boolean)`

### 3. 将来拡張のためのルール

- エラーソース名は実クラス名に一致させる（例: `OneDim.Read_Array`）。
- 例外番号帯をクラスごとに分離する。
  - `OneDim`: 1000番台
  - `TwoDim`: 2000番台
  - `DataFrame`: 3000番台
- `Read_*` 後の不変条件（行数、列数、ロード済み）を明文化する。

## 実装順（推奨）

1. `TwoDim.cls` を追加し、Range/配列の入出力を固める。
2. `TwoDim` から `GetCol` で `OneDim` を返す導線を作る。
3. `DataFrame.cls` を列指向で実装し、`FromRange` / `ToRange` を接続する。
4. 型変換・フィルタ・ソートなど高次機能を段階的に追加する。

## 先に直しておくと安全な点

- `OneDim.cls` 内の `Err.Raise` の source が `Araya1.*` になっている箇所は `OneDim.*` に揃える。
- コメントの文字コードが環境によって文字化けする場合は UTF-8 で管理する。
