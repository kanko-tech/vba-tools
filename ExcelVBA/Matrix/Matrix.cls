VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------
' プロパティ
'----------------------------
Private mData As Variant
Private mIsLoaded As Boolean
Private mRowCount As Long
Private mColCount As Long

'----------------------------
' 読込系
'----------------------------
Public Sub read_range(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 2001, "Matrix.read_range", "引数rngがNothingです。"

    Dim arr As Variant
    If rng.Cells.Count = 1 Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = rng.Value
        load_matrix_internal arr, 1, 1
    Else
        arr = rng.Value
        load_matrix_internal normalize_matrix_data(arr), rng.Rows.Count, rng.Columns.Count
    End If
End Sub

Public Sub read_array(ByVal arr As Variant)
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 2002, "Matrix.read_array", "配列ではありません。"
    End If

    ensure_is_two_dimensional arr
    load_matrix_internal normalize_matrix_data(arr), UBound(arr, 1) - LBound(arr, 1) + 1, UBound(arr, 2) - LBound(arr, 2) + 1
End Sub

Public Sub read_empty(ByVal colCount As Long)
    If colCount < 1 Then
        Err.Raise vbObjectError + 2005, "Matrix.read_empty", "colCount は 1 以上を指定してください。"
    End If

    load_matrix_internal build_empty_matrix_placeholder(colCount), 0, colCount
End Sub

'----------------------------
' ガード系
'----------------------------
Private Sub ensure_loaded(Optional ByVal caller As String = "")
    If Not mIsLoaded Then
        Dim msg As String
        msg = "データが読み込まれていません。read_range / read_array / read_empty を先に呼び出してください。"

        If caller <> "" Then
            msg = msg & vbCrLf & "呼び出し元: " & caller
        End If

        Err.Raise vbObjectError + 2099, "Matrix.ensure_loaded", msg
    End If
End Sub

Private Sub ensure_is_two_dimensional(ByVal arr As Variant)
    On Error Resume Next

    Dim lb1 As Long
    Dim lb2 As Long
    Dim lb3 As Long

    lb1 = LBound(arr, 1)
    lb2 = LBound(arr, 2)

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 2003, "Matrix.ensure_is_two_dimensional", "2次元配列を指定してください。"
    End If

    lb3 = LBound(arr, 3)
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 2004, "Matrix.ensure_is_two_dimensional", "3次元以上の配列は受け付けません。"
    End If

    Err.Clear
    On Error GoTo 0
End Sub

'----------------------------
' 内部補助
'----------------------------
Private Sub load_matrix_internal(ByVal src As Variant, ByVal rowCount As Long, ByVal colCount As Long)
    mData = src
    mRowCount = rowCount
    mColCount = colCount
    mIsLoaded = True
End Sub

Private Function normalize_matrix_data(ByVal src As Variant) As Variant
    Dim rows As Long, cols As Long
    Dim r As Long, c As Long
    Dim r1 As Long, c1 As Long
    Dim dst As Variant

    rows = UBound(src, 1) - LBound(src, 1) + 1
    cols = UBound(src, 2) - LBound(src, 2) + 1
    r1 = LBound(src, 1)
    c1 = LBound(src, 2)

    ReDim dst(1 To rows, 1 To cols)
    For r = 1 To rows
        For c = 1 To cols
            dst(r, c) = src(r1 + r - 1, c1 + c - 1)
        Next c
    Next r

    normalize_matrix_data = dst
End Function

Private Function clone_matrix_data(ByVal src As Variant, ByVal rowCount As Long, ByVal colCount As Long) As Variant
    Dim r As Long, c As Long
    Dim dst As Variant

    If rowCount = 0 Then
        clone_matrix_data = build_empty_matrix_placeholder(colCount)
        Exit Function
    End If

    ReDim dst(1 To rowCount, 1 To colCount)
    For r = 1 To rowCount
        For c = 1 To colCount
            dst(r, c) = src(r, c)
        Next c
    Next r

    clone_matrix_data = dst
End Function

Private Function build_empty_matrix_placeholder(ByVal colCount As Long) As Variant
    Dim dst As Variant
    Dim c As Long

    ReDim dst(1 To 1, 1 To colCount)
    For c = 1 To colCount
        dst(1, c) = Empty
    Next c

    build_empty_matrix_placeholder = dst
End Function

Private Sub ensure_valid_row_index(ByVal rowIndex As Long, ByVal caller As String)
    If rowIndex < 1 Or rowIndex > mRowCount Then
        Err.Raise vbObjectError + 2010, caller, "rowIndex が範囲外です。"
    End If
End Sub

Private Sub ensure_valid_col_index(ByVal colIndex As Long, ByVal caller As String)
    If colIndex < 1 Or colIndex > mColCount Then
        Err.Raise vbObjectError + 2011, caller, "colIndex が範囲外です。"
    End If
End Sub

'----------------------------
' 参照系
'----------------------------
Public Property Get is_loaded() As Boolean
    is_loaded = mIsLoaded
End Property

Public Property Get row_count() As Long
    ensure_loaded
    row_count = mRowCount
End Property

Public Property Get col_count() As Long
    ensure_loaded
    col_count = mColCount
End Property

Public Property Get data() As Variant
    ensure_loaded
    data = clone_matrix_data(mData, mRowCount, mColCount)
End Property

Public Property Get item(ByVal rowIndex As Long, ByVal colIndex As Long) As Variant
    ensure_loaded
    ensure_valid_row_index rowIndex, "Matrix.item"
    ensure_valid_col_index colIndex, "Matrix.item"
    item = mData(rowIndex, colIndex)
End Property

Public Function row_values(ByVal rowIndex As Long) As Variant
    ensure_loaded
    ensure_valid_row_index rowIndex, "Matrix.row_values"

    Dim c As Long
    Dim res As Variant

    ReDim res(1 To mColCount)
    For c = 1 To mColCount
        res(c) = mData(rowIndex, c)
    Next c

    row_values = res
End Function

Public Function col_values(ByVal colIndex As Long) As Variant
    ensure_loaded
    ensure_valid_col_index colIndex, "Matrix.col_values"

    If mRowCount = 0 Then
        Err.Raise vbObjectError + 2014, "Matrix.col_values", "行が0件のため列データを返せません。"
    End If

    Dim r As Long
    Dim res As Variant

    ReDim res(1 To mRowCount)
    For r = 1 To mRowCount
        res(r) = mData(r, colIndex)
    Next r

    col_values = res
End Function

'----------------------------
' 抽出・選択系
'----------------------------
Public Function filter_rows(ByVal mask As Variant) As Matrix
    ensure_loaded

    If Not IsArray(mask) Then
        Err.Raise vbObjectError + 2020, "Matrix.filter_rows", "mask は一次元配列で指定してください。"
    End If

    Dim maskLb As Long, maskUb As Long
    Dim r As Long, c As Long, keepCount As Long, outRow As Long
    Dim dst As Variant
    Dim result As Matrix

    maskLb = LBound(mask)
    maskUb = UBound(mask)
    If (maskUb - maskLb + 1) <> mRowCount Then
        Err.Raise vbObjectError + 2021, "Matrix.filter_rows", "mask の長さが行数と一致しません。"
    End If

    keepCount = 0
    For r = 1 To mRowCount
        If CBool(mask(maskLb + r - 1)) Then keepCount = keepCount + 1
    Next r

    Set result = New Matrix
    If keepCount = 0 Then
        result.read_empty mColCount
        Set filter_rows = result
        Exit Function
    End If

    ReDim dst(1 To keepCount, 1 To mColCount)
    outRow = 1
    For r = 1 To mRowCount
        If CBool(mask(maskLb + r - 1)) Then
            For c = 1 To mColCount
                dst(outRow, c) = mData(r, c)
            Next c
            outRow = outRow + 1
        End If
    Next r

    result.read_array dst
    Set filter_rows = result
End Function

Public Function slice_rows(ByVal rowIndexes As Variant) As Matrix
    ensure_loaded

    If Not IsArray(rowIndexes) Then
        Err.Raise vbObjectError + 2028, "Matrix.slice_rows", "rowIndexes は一次元配列で指定してください。"
    End If

    Dim lb As Long, ub As Long, i As Long, c As Long
    Dim dst As Variant
    Dim result As Matrix
    Dim rowIndex As Long

    lb = LBound(rowIndexes)
    ub = UBound(rowIndexes)
    Set result = New Matrix

    If ub < lb Then
        result.read_empty mColCount
        Set slice_rows = result
        Exit Function
    End If

    ReDim dst(1 To ub - lb + 1, 1 To mColCount)
    For i = lb To ub
        rowIndex = CLng(rowIndexes(i))
        ensure_valid_row_index rowIndex, "Matrix.slice_rows"
        For c = 1 To mColCount
            dst(1 + (i - lb), c) = mData(rowIndex, c)
        Next c
    Next i

    result.read_array dst
    Set slice_rows = result
End Function

Public Function select_columns(ByVal columnIndexes As Variant) As Matrix
    ensure_loaded

    If Not IsArray(columnIndexes) Then
        Err.Raise vbObjectError + 2022, "Matrix.select_columns", "columnIndexes は一次元配列で指定してください。"
    End If

    Dim lb As Long, ub As Long, i As Long, r As Long, srcCol As Long
    Dim dst As Variant
    Dim result As Matrix

    lb = LBound(columnIndexes)
    ub = UBound(columnIndexes)
    If ub < lb Then
        Err.Raise vbObjectError + 2023, "Matrix.select_columns", "columnIndexes は1件以上指定してください。"
    End If

    Set result = New Matrix
    If mRowCount = 0 Then
        result.read_empty ub - lb + 1
        Set select_columns = result
        Exit Function
    End If

    ReDim dst(1 To mRowCount, 1 To ub - lb + 1)
    For i = lb To ub
        srcCol = CLng(columnIndexes(i))
        ensure_valid_col_index srcCol, "Matrix.select_columns"
        For r = 1 To mRowCount
            dst(r, 1 + (i - lb)) = mData(r, srcCol)
        Next r
    Next i

    result.read_array dst
    Set select_columns = result
End Function

'----------------------------
' 更新・追加系
'----------------------------
Public Sub set_column_values(ByVal colIndex As Long, ByVal values As Variant)
    ensure_loaded
    ensure_valid_col_index colIndex, "Matrix.set_column_values"

    If Not IsArray(values) Then
        Err.Raise vbObjectError + 2025, "Matrix.set_column_values", "values は一次元配列で指定してください。"
    End If

    If mRowCount = 0 Then Exit Sub

    Dim lb As Long, ub As Long, r As Long
    lb = LBound(values)
    ub = UBound(values)
    If (ub - lb + 1) <> mRowCount Then
        Err.Raise vbObjectError + 2026, "Matrix.set_column_values", "values の長さが行数と一致しません。"
    End If

    For r = 1 To mRowCount
        mData(r, colIndex) = values(lb + r - 1)
    Next r
End Sub

Public Sub append_row(ByVal values As Variant)
    ensure_loaded

    If Not IsArray(values) Then
        Err.Raise vbObjectError + 2029, "Matrix.append_row", "values は一次元配列で指定してください。"
    End If

    Dim lb As Long, ub As Long, r As Long, c As Long
    Dim dst As Variant

    lb = LBound(values)
    ub = UBound(values)
    If (ub - lb + 1) <> mColCount Then
        Err.Raise vbObjectError + 2031, "Matrix.append_row", "values の長さが列数と一致しません。"
    End If

    ReDim dst(1 To mRowCount + 1, 1 To mColCount)
    For r = 1 To mRowCount
        For c = 1 To mColCount
            dst(r, c) = mData(r, c)
        Next c
    Next r
    For c = 1 To mColCount
        dst(mRowCount + 1, c) = values(lb + c - 1)
    Next c

    load_matrix_internal dst, mRowCount + 1, mColCount
End Sub

Public Sub append_column(ByVal values As Variant)
    ensure_loaded

    Dim r As Long, c As Long
    Dim dst As Variant

    If mRowCount > 0 Then
        If Not IsArray(values) Then
            Err.Raise vbObjectError + 2032, "Matrix.append_column", "values は一次元配列で指定してください。"
        End If

        Dim lb As Long, ub As Long
        lb = LBound(values)
        ub = UBound(values)
        If (ub - lb + 1) <> mRowCount Then
            Err.Raise vbObjectError + 2033, "Matrix.append_column", "values の長さが行数と一致しません。"
        End If

        ReDim dst(1 To mRowCount, 1 To mColCount + 1)
        For r = 1 To mRowCount
            For c = 1 To mColCount
                dst(r, c) = mData(r, c)
            Next c
            dst(r, mColCount + 1) = values(lb + r - 1)
        Next r
        load_matrix_internal dst, mRowCount, mColCount + 1
    Else
        load_matrix_internal build_empty_matrix_placeholder(mColCount + 1), 0, mColCount + 1
    End If
End Sub

'----------------------------
' 変形系
'----------------------------
Public Function transpose() As Matrix
    ensure_loaded

    Dim dst As Variant
    Dim r As Long, c As Long
    Dim result As Matrix

    Set result = New Matrix

    If mRowCount = 0 Then
        Err.Raise vbObjectError + 2035, "Matrix.transpose", "行が0件のMatrixの転置はサポートしていません。"
        Set transpose = result
        Exit Function
    End If

    ReDim dst(1 To mColCount, 1 To mRowCount)
    For r = 1 To mRowCount
        For c = 1 To mColCount
            dst(c, r) = mData(r, c)
        Next c
    Next r

    result.read_array dst
    Set transpose = result
End Function

'----------------------------
' 出力系
'----------------------------
Public Sub to_range(ByVal topLeft As Range)
    ensure_loaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 2030, "Matrix.to_range", "topLeft が Nothing です。"
    If mRowCount = 0 Then Err.Raise vbObjectError + 2034, "Matrix.to_range", "行が0件のためシートへ出力できません。"

    topLeft.Resize(mRowCount, mColCount).Value = mData
End Sub

'----------------------------
' 状態管理系
'----------------------------
Public Sub clear()
    mData = Empty
    mIsLoaded = False
    mRowCount = 0
    mColCount = 0
End Sub
