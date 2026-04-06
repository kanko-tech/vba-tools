# ExcelVBA

`src/ExcelVBA` ディレクトリには、Excel VBA で再利用しやすいクラスモジュールを整理して配置しています。  
現在は、一次元配列を扱う `Vector`、二次元配列を扱う `Matrix`、列名付きテーブルを扱う `Table` を中心に構成しています。

## まず見るページ

- ルート README: `README.md`
- GitHub Pages トップ: `docs/index.html`
- `Vector` 詳細: `docs/vector.html`
- `Matrix` 詳細: `docs/matrix.html`
- `Table` 詳細: `docs/table.html`

## モジュールの使い分け

### `Vector`

単一列や単一行の値をまとめて処理したいときに使います。

- `Range` から 1 次元配列を読み込む
- 比較マスクを作る
- 型変換や欠損補完を行う
- 合計や平均を計算する

### `Matrix`

列名を持たない表データ全体を扱いたいときに使います。

- 矩形 `Range` を 2 次元配列として読む
- 行抽出や列選択を行う
- 行追加、列追加、列更新を行う
- シートへ一括出力する

### `Table`

列名つきの実務データを読みやすく操作したいときに使います。

- ヘッダー付きの表を読む
- 列名ベースで抽出や更新を行う
- 列追加や並べ替えを行う
- `Vector` や `Matrix` と連携する

## 推奨する読み進め方

1. まずこの README で全体像をつかむ
2. `docs/index.html` でディレクトリ構成と使い分けを確認する
3. 必要なモジュールの詳細ページを見る
4. 実装ファイルと同梱 README で個別メソッドを確認する

## ディレクトリ一覧

- `src/ExcelVBA/Vector`
- `src/ExcelVBA/Matrix`
- `src/ExcelVBA/Table`

## 詳細ドキュメントの位置づけ

この README は入口に絞り、詳しい説明や使い方は GitHub Pages 側へ段階的に移しています。  
実装に密着したメソッド仕様は、各ディレクトリの README もあわせて参照してください。
