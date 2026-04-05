'----------------------------
'クラスモジュールのメタ属性
'----------------------------
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
Private MATRIX As Variant
Private MATRIX_IS_LOADED As Boolean

'----------------------------
' 読込系
'----------------------------
Public Sub read_range(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 2001, "Matrix.read_range", "引数rngがNothingです。"

    If rng.Cells.Count = 1 Then
        Dim tmp1(1 To 1, 1 To 1) As Variant
        tmp1(1, 1) = rng.Value
        MATRIX = tmp1
    Else
        MATRIX = rng.Value
    End If

    MATRIX_IS_LOADED = True
End Sub

Public Sub read_array(ByVal arr As Variant)
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 2002, "Matrix.read_array", "配列ではありません。"
    End If

    ensure_is_two_dimensional arr

    MATRIX = arr
    MATRIX_IS_LOADED = True
End Sub

'----------------------------
' ガード系
'----------------------------
Private Sub ensure_loaded(Optional ByVal caller As String = "")
    If Not MATRIX_IS_LOADED Then
        Dim msg As String
        msg = "データが読み込まれていません。read_range / read_array を先に呼び出してください。"

        If caller <> "" Then
            msg = msg & vbCrLf & "呼び出し元: " & caller
        End If

        Err.Raise vbObjectError + 2099, "Matrix.ensure_loaded", msg
    End If
End Sub

Private Sub ensure_is_two_dimensional(ByVal arr As Variant)
    On Error Resume Next

    Dim lb1 As Long
    Dim ub1 As Long
    Dim lb2 As Long
    Dim ub2 As Long
    Dim lb3 As Long

    lb1 = LBound(arr, 1)
    ub1 = UBound(arr, 1)
    lb2 = LBound(arr, 2)
    ub2 = UBound(arr, 2)

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
Private Function clone_matrix_data(ByVal src As Variant) As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim r As Long, c As Long
    Dim dst As Variant

    r1 = LBound(src, 1)
    r2 = UBound(src, 1)
    c1 = LBound(src, 2)
    c2 = UBound(src, 2)

    ReDim dst(r1 To r2, c1 To c2)

    For r = r1 To r2
        For c = c1 To c2
            dst(r, c) = src(r, c)
        Next c
    Next r

    clone_matrix_data = dst
End Function

'----------------------------
' 参照系
'----------------------------
Public Property Get is_loaded() As Boolean
    is_loaded = MATRIX_IS_LOADED
End Property

Public Property Get row_count() As Long
    ensure_loaded
    row_count = UBound(MATRIX, 1) - LBound(MATRIX, 1) + 1
End Property

Public Property Get col_count() As Long
    ensure_loaded
    col_count = UBound(MATRIX, 2) - LBound(MATRIX, 2) + 1
End Property

Public Property Get data() As Variant
    ensure_loaded
    data = clone_matrix_data(MATRIX)
End Property

Public Property Get item(ByVal rowIndex As Long, ByVal colIndex As Long) As Variant
    ensure_loaded

    If rowIndex < LBound(MATRIX, 1) Or rowIndex > UBound(MATRIX, 1) Then
        Err.Raise vbObjectError + 2010, "Matrix.item", "rowIndex が範囲外です。"
    End If
    If colIndex < LBound(MATRIX, 2) Or colIndex > UBound(MATRIX, 2) Then
        Err.Raise vbObjectError + 2011, "Matrix.item", "colIndex が範囲外です。"
    End If

    item = MATRIX(rowIndex, colIndex)
End Property

Public Function row_values(ByVal rowIndex As Long) As Variant
    ensure_loaded

    If rowIndex < LBound(MATRIX, 1) Or rowIndex > UBound(MATRIX, 1) Then
        Err.Raise vbObjectError + 2012, "Matrix.row_values", "rowIndex が範囲外です。"
    End If

    Dim c1 As Long, c2 As Long, c As Long
    Dim res As Variant

    c1 = LBound(MATRIX, 2)
    c2 = UBound(MATRIX, 2)
    ReDim res(c1 To c2)

    For c = c1 To c2
        res(c) = MATRIX(rowIndex, c)
    Next c

    row_values = res
End Function

Public Function col_values(ByVal colIndex As Long) As Variant
    ensure_loaded

    If colIndex < LBound(MATRIX, 2) Or colIndex > UBound(MATRIX, 2) Then
        Err.Raise vbObjectError + 2013, "Matrix.col_values", "colIndex が範囲外です。"
    End If

    Dim r1 As Long, r2 As Long, r As Long
    Dim res As Variant

    r1 = LBound(MATRIX, 1)
    r2 = UBound(MATRIX, 1)
    ReDim res(r1 To r2)

    For r = r1 To r2
        res(r) = MATRIX(r, colIndex)
    Next r

    col_values = res
End Function

'----------------------------
' 行列操作系
'----------------------------
Public Function filter_rows(ByVal mask As Variant) As Matrix
    ensure_loaded

    Dim maskLb As Long, maskUb As Long
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim r As Long, c As Long, outRow As Long
    Dim keepCount As Long
    Dim dst As Variant
    Dim result As Matrix

    If Not IsArray(mask) Then
        Err.Raise vbObjectError + 2020, "Matrix.filter_rows", "mask は一次元配列で指定してください。"
    End If

    maskLb = LBound(mask)
    maskUb = UBound(mask)
    r1 = LBound(MATRIX, 1)
    r2 = UBound(MATRIX, 1)
    c1 = LBound(MATRIX, 2)
    c2 = UBound(MATRIX, 2)

    If (maskUb - maskLb) <> (r2 - r1) Then
        Err.Raise vbObjectError + 2021, "Matrix.filter_rows", "mask の長さが行数と一致しません。"
    End If

    keepCount = 0
    For r = r1 To r2
        If CBool(mask(maskLb + (r - r1))) Then keepCount = keepCount + 1
    Next r

    If keepCount = 0 Then
        Err.Raise vbObjectError + 2027, "Matrix.filter_rows", "条件に一致する行がありません。"
    End If

    ReDim dst(1 To keepCount, c1 To c2)
    outRow = 1
    For r = r1 To r2
        If CBool(mask(maskLb + (r - r1))) Then
            For c = c1 To c2
                dst(outRow, c) = MATRIX(r, c)
            Next c
            outRow = outRow + 1
        End If
    Next r

    Set result = New Matrix
    result.read_array dst
    Set filter_rows = result
End Function

Public Function select_columns(ByVal columnIndexes As Variant) As Matrix
    ensure_loaded

    If Not IsArray(columnIndexes) Then
        Err.Raise vbObjectError + 2022, "Matrix.select_columns", "columnIndexes は一次元配列で指定してください。"
    End If

    Dim idxLb As Long, idxUb As Long
    Dim r1 As Long, r2 As Long
    Dim outCol As Long, r As Long, i As Long, srcCol As Long
    Dim dst As Variant
    Dim result As Matrix

    idxLb = LBound(columnIndexes)
    idxUb = UBound(columnIndexes)
    r1 = LBound(MATRIX, 1)
    r2 = UBound(MATRIX, 1)

    ReDim dst(r1 To r2, 1 To idxUb - idxLb + 1)

    outCol = 1
    For i = idxLb To idxUb
        srcCol = CLng(columnIndexes(i))
        If srcCol < LBound(MATRIX, 2) Or srcCol > UBound(MATRIX, 2) Then
            Err.Raise vbObjectError + 2023, "Matrix.select_columns", "columnIndexes に範囲外の列番号が含まれています。"
        End If

        For r = r1 To r2
            dst(r, outCol) = MATRIX(r, srcCol)
        Next r
        outCol = outCol + 1
    Next i

    Set result = New Matrix
    result.read_array dst
    Set select_columns = result
End Function

Public Sub set_column_values(ByVal colIndex As Long, ByVal values As Variant)
    ensure_loaded

    If colIndex < LBound(MATRIX, 2) Or colIndex > UBound(MATRIX, 2) Then
        Err.Raise vbObjectError + 2024, "Matrix.set_column_values", "colIndex が範囲外です。"
    End If

    If Not IsArray(values) Then
        Err.Raise vbObjectError + 2025, "Matrix.set_column_values", "values は一次元配列で指定してください。"
    End If

    Dim vLb As Long, vUb As Long, r1 As Long, r2 As Long, r As Long
    vLb = LBound(values)
    vUb = UBound(values)
    r1 = LBound(MATRIX, 1)
    r2 = UBound(MATRIX, 1)

    If (vUb - vLb) <> (r2 - r1) Then
        Err.Raise vbObjectError + 2026, "Matrix.set_column_values", "values の長さが行数と一致しません。"
    End If

    For r = r1 To r2
        MATRIX(r, colIndex) = values(vLb + (r - r1))
    Next r
End Sub

'----------------------------
' 出力系
'----------------------------
Public Sub to_range(ByVal topLeft As Range)
    ensure_loaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 2030, "Matrix.to_range", "topLeft が Nothing です。"

    topLeft.Resize(Me.row_count, Me.col_count).Value = MATRIX
End Sub

Public Sub clear()
    MATRIX = Empty
    MATRIX_IS_LOADED = False
End Sub
