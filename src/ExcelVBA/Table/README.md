# クラスモジュール `Table`

`Table` は、列名を持つ表データを扱うための VBA クラスモジュールです。  
内部では `Matrix` を使いながら、列名ベースの抽出、更新、並べ替え、出力を行います。

## この README の役割

この README は入口に絞っています。  
メソッド早見表、補足、レシピ、全メソッド解説は GitHub Pages 側に移しています。

## 詳細ドキュメント

- Pages ガイド: [Table Guide](https://kanko-tech.github.io/VBA_tools/table.html)
- ソース配置: `src/ExcelVBA/Table`

## まず何ができるか

- ヘッダー付きの表を、そのまま列名つきデータとして読み込む
- 列名を使って条件抽出する
- 条件に合う行だけ別列を更新する
- 必要な列だけを選んで新しい `Table` を作る
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
