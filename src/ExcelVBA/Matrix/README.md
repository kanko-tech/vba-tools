# クラスモジュール `Matrix`

`Matrix` は、2 次元配列を扱うための VBA クラスモジュールです。  
列名を持たない表データを保持し、行抽出、列選択、更新、転置、シート出力を行います。

## この README の役割

この README は入口に絞っています。  
メソッド早見表、補足、レシピ、全メソッド解説は GitHub Pages 側に移しています。

## 詳細ドキュメント

- Pages ガイド: [Matrix Guide](https://kanko-tech.github.io/VBA_tools/matrix.html)
- ソース配置: `src/ExcelVBA/Matrix`

## まず何ができるか

- 矩形 `Range` や 2 次元配列を、そのまま表データとして保持する
- 行マスクで必要な行だけを抜き出す
- 必要な列だけを選んで新しい `Matrix` を作る
- 列の差し替え、行追加、列追加で表構造を更新する
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
