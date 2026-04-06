VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------
' プロパティ
'----------------------------
Private DATA As Matrix
Private COLUMN_NAMES As Variant
Private COLUMN_INDEX_MAP As Object
Private TABLE_IS_LOADED As Boolean

'----------------------------
' 読込系
'----------------------------
Public Sub read_range(ByVal rng As Range, Optional ByVal hasHeader As Boolean = True)
    If rng Is Nothing Then Err.Raise vbObjectError + 3001, "Table.read_range", "引数rngがNothingです。"

    Dim raw As Variant
    Dim rows As Long, cols As Long
    Dim r As Long, c As Long
    Dim body As Variant

    If rng.Cells.Count = 1 Then
        ReDim raw(1 To 1, 1 To 1)
        raw(1, 1) = rng.Value
        rows = 1
        cols = 1
    Else
        raw = rng.Value
        rows = UBound(raw, 1) - LBound(raw, 1) + 1
        cols = UBound(raw, 2) - LBound(raw, 2) + 1
    End If

    ReDim COLUMN_NAMES(1 To cols)

    If hasHeader Then
        For c = 1 To cols
            COLUMN_NAMES(c) = CStr(raw(1, c))
        Next c

        If rows = 1 Then
            Set DATA = New Matrix
            DATA.read_empty cols
            rebuild_column_index_map
            TABLE_IS_LOADED = True
            Exit Sub
        End If

        ReDim body(1 To rows - 1, 1 To cols)
        For r = 2 To rows
            For c = 1 To cols
                body(r - 1, c) = raw(r, c)
            Next c
        Next r
    Else
        For c = 1 To cols
            COLUMN_NAMES(c) = "col" & CStr(c)
        Next c

        ReDim body(1 To rows, 1 To cols)
        For r = 1 To rows
            For c = 1 To cols
                body(r, c) = raw(r, c)
            Next c
        Next r
    End If

    load_internal body, COLUMN_NAMES
End Sub

Public Sub read_matrix(ByVal src As Matrix, ByVal columnNames As Variant)
    If src Is Nothing Then
        Err.Raise vbObjectError + 3003, "Table.read_matrix", "引数srcがNothingです。"
    End If
    If Not IsArray(columnNames) Then
        Err.Raise vbObjectError + 3004, "Table.read_matrix", "columnNames は一次元配列で指定してください。"
    End If

    load_from_matrix src, columnNames
End Sub

'----------------------------
' ガード系
'----------------------------
Private Sub ensure_loaded(Optional ByVal caller As String = "")
    If Not TABLE_IS_LOADED Then
        Dim msg As String
        msg = "データが読み込まれていません。read_range / read_matrix を先に呼び出してください。"

        If caller <> "" Then
            msg = msg & vbCrLf & "呼び出し元: " & caller
        End If

        Err.Raise vbObjectError + 3099, "Table.ensure_loaded", msg
    End If
End Sub

'----------------------------
' 内部補助
'----------------------------
Private Sub load_internal(ByVal dataArr As Variant, ByVal columnNames As Variant)
    If Not IsArray(columnNames) Then
        Err.Raise vbObjectError + 3005, "Table.load_internal", "columnNames は一次元配列で指定してください。"
    End If

    Dim normalizedNames As Variant
    normalizedNames = normalize_column_names(columnNames)

    If (UBound(normalizedNames) - LBound(normalizedNames) + 1) <> (UBound(dataArr, 2) - LBound(dataArr, 2) + 1) Then
        Err.Raise vbObjectError + 3006, "Table.load_internal", "列名数と列数が一致しません。"
    End If

    Set DATA = New Matrix
    DATA.read_array dataArr

    COLUMN_NAMES = normalizedNames
    rebuild_column_index_map
    TABLE_IS_LOADED = True
End Sub

Private Sub load_from_matrix(ByVal src As Matrix, ByVal columnNames As Variant)
    Dim normalizedNames As Variant
    normalizedNames = normalize_column_names(columnNames)

    If (UBound(normalizedNames) - LBound(normalizedNames) + 1) <> src.col_count Then
        Err.Raise vbObjectError + 3006, "Table.load_from_matrix", "列名数と列数が一致しません。"
    End If

    Set DATA = New Matrix
    If src.row_count = 0 Then
        DATA.read_empty src.col_count
    Else
        DATA.read_array src.data
    End If

    COLUMN_NAMES = normalizedNames
    rebuild_column_index_map
    TABLE_IS_LOADED = True
End Sub

Private Function normalize_column_names(ByVal columnNames As Variant) As Variant
    Dim srcLb As Long, srcUb As Long, i As Long
    Dim names As Variant

    srcLb = LBound(columnNames)
    srcUb = UBound(columnNames)
    ReDim names(1 To srcUb - srcLb + 1)

    For i = srcLb To srcUb
        names(1 + (i - srcLb)) = CStr(columnNames(i))
    Next i

    normalize_column_names = names
End Function

Private Function clone_column_names() As Variant
    Dim i As Long
    Dim names As Variant

    ReDim names(1 To UBound(COLUMN_NAMES))
    For i = 1 To UBound(COLUMN_NAMES)
        names(i) = COLUMN_NAMES(i)
    Next i

    clone_column_names = names
End Function

Private Sub rebuild_column_index_map()
    Dim i As Long

    Set COLUMN_INDEX_MAP = CreateObject("Scripting.Dictionary")
    COLUMN_INDEX_MAP.CompareMode = 1

    For i = LBound(COLUMN_NAMES) To UBound(COLUMN_NAMES)
        If Trim$(CStr(COLUMN_NAMES(i))) = "" Then
            Err.Raise vbObjectError + 3009, "Table.rebuild_column_index_map", "空の列名は使用できません。"
        End If
        If COLUMN_INDEX_MAP.Exists(CStr(COLUMN_NAMES(i))) Then
            Err.Raise vbObjectError + 3007, "Table.rebuild_column_index_map", "列名が重複しています: " & CStr(COLUMN_NAMES(i))
        End If
        COLUMN_INDEX_MAP.Add CStr(COLUMN_NAMES(i)), i
    Next i
End Sub

Private Function resolve_column_index(ByVal columnName As String, ByVal caller As String) As Long
    ensure_loaded caller

    If Not COLUMN_INDEX_MAP.Exists(columnName) Then
        Err.Raise vbObjectError + 3008, caller, "指定された列名が存在しません: " & columnName
    End If

    resolve_column_index = CLng(COLUMN_INDEX_MAP(columnName))
End Function

Private Function build_mask_for_equals(ByVal colIndex As Long, ByVal matchValue As Variant) As Variant
    Dim values As Variant
    Dim i As Long
    Dim mask As Variant

    values = DATA.col_values(colIndex)
    ReDim mask(1 To DATA.row_count)

    For i = 1 To DATA.row_count
        If compare_equals_safe(values(i), matchValue) Then
            mask(i) = True
        Else
            mask(i) = False
        End If
    Next i

    build_mask_for_equals = mask
End Function

Private Function build_mask_for_in(ByVal colIndex As Long, ByVal matchValues As Variant) As Variant
    Dim values As Variant
    Dim i As Long, j As Long
    Dim mask As Variant
    Dim lb As Long, ub As Long

    If Not IsArray(matchValues) Then
        Err.Raise vbObjectError + 3013, "Table.build_mask_for_in", "matchValues は一次元配列で指定してください。"
    End If

    values = DATA.col_values(colIndex)
    lb = LBound(matchValues)
    ub = UBound(matchValues)
    ReDim mask(1 To DATA.row_count)

    For i = 1 To DATA.row_count
        mask(i) = False
        For j = lb To ub
            If compare_equals_safe(values(i), matchValues(j)) Then
                mask(i) = True
                Exit For
            End If
        Next j
    Next i

    build_mask_for_in = mask
End Function

Private Function build_mask_for_contains(ByVal colIndex As Long, ByVal searchText As String, ByVal caseSensitive As Boolean) As Variant
    Dim values As Variant
    Dim mask As Variant
    Dim i As Long
    Dim haystack As String
    Dim needle As String

    values = DATA.col_values(colIndex)
    ReDim mask(1 To DATA.row_count)

    needle = searchText
    If Not caseSensitive Then needle = LCase$(needle)

    For i = 1 To DATA.row_count
        If IsError(values(i)) Or IsEmpty(values(i)) Or IsNull(values(i)) Then
            mask(i) = False
        Else
            haystack = CStr(values(i))
            If Not caseSensitive Then haystack = LCase$(haystack)
            mask(i) = (InStr(1, haystack, needle, vbBinaryCompare) > 0)
        End If
    Next i

    build_mask_for_contains = mask
End Function

Private Function combine_masks(ByVal leftMask As Variant, ByVal rightMask As Variant, ByVal useAnd As Boolean) As Variant
    Dim i As Long
    Dim mask As Variant

    ReDim mask(1 To DATA.row_count)
    For i = 1 To DATA.row_count
        If useAnd Then
            mask(i) = CBool(leftMask(i)) And CBool(rightMask(i))
        Else
            mask(i) = CBool(leftMask(i)) Or CBool(rightMask(i))
        End If
    Next i

    combine_masks = mask
End Function

Private Function subset_column_names(ByVal selectedIndexes As Variant) As Variant
    Dim lb As Long, ub As Long, i As Long
    Dim names As Variant

    lb = LBound(selectedIndexes)
    ub = UBound(selectedIndexes)
    ReDim names(1 To ub - lb + 1)

    For i = lb To ub
        names(1 + (i - lb)) = COLUMN_NAMES(CLng(selectedIndexes(i)))
    Next i

    subset_column_names = names
End Function

Private Sub swap_rows(ByRef arr As Variant, ByVal rowA As Long, ByVal rowB As Long)
    Dim c As Long
    Dim tmp As Variant

    For c = 1 To UBound(arr, 2)
        tmp = arr(rowA, c)
        arr(rowA, c) = arr(rowB, c)
        arr(rowB, c) = tmp
    Next c
End Sub

Private Function compare_equals_safe(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If IsError(leftValue) Or IsError(rightValue) Or IsNull(leftValue) Or IsNull(rightValue) Then
        compare_equals_safe = False
        Exit Function
    End If

    On Error Resume Next
    compare_equals_safe = (leftValue = rightValue)
    If Err.Number <> 0 Then
        Err.Clear
        compare_equals_safe = False
    End If
    On Error GoTo 0
End Function

Private Function compare_order_safe(ByVal leftValue As Variant, ByVal rightValue As Variant, ByVal ascending As Boolean) As Boolean
    If IsError(leftValue) Or IsError(rightValue) Or IsNull(leftValue) Or IsNull(rightValue) Then
        compare_order_safe = False
        Exit Function
    End If

    On Error Resume Next
    If ascending Then
        compare_order_safe = (leftValue > rightValue)
    Else
        compare_order_safe = (leftValue < rightValue)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        compare_order_safe = False
    End If
    On Error GoTo 0
End Function

Private Function create_empty_column_values() As Variant
    Dim values As Variant
    Dim i As Long

    ReDim values(1 To DATA.row_count)
    For i = 1 To DATA.row_count
        values(i) = Empty
    Next i

    create_empty_column_values = values
End Function

Private Function clone_empty_table() As Table
    Dim result As Table
    Dim emptyMatrix As Matrix

    Set result = New Table
    Set emptyMatrix = New Matrix
    emptyMatrix.read_empty DATA.col_count
    result.read_matrix emptyMatrix, clone_column_names()
    Set clone_empty_table = result
End Function

'----------------------------
' 参照系
'----------------------------
Public Property Get is_loaded() As Boolean
    is_loaded = TABLE_IS_LOADED
End Property

Public Property Get row_count() As Long
    ensure_loaded
    row_count = DATA.row_count
End Property

Public Property Get col_count() As Long
    ensure_loaded
    col_count = DATA.col_count
End Property

Public Property Get column_names() As Variant
    ensure_loaded
    column_names = clone_column_names()
End Property

Public Property Get matrix() As Matrix
    ensure_loaded

    Dim result As Matrix
    Set result = New Matrix
    If DATA.row_count = 0 Then
        result.read_empty DATA.col_count
    Else
        result.read_array DATA.data
    End If
    Set matrix = result
End Property

Public Function col(ByVal columnName As String) As Variant
    col = DATA.col_values(resolve_column_index(columnName, "Table.col"))
End Function

Public Function col_vector(ByVal columnName As String) As Vector
    Dim vec As Vector
    Set vec = New Vector
    vec.read_array Me.col(columnName)
    Set col_vector = vec
End Function

'----------------------------
' 抽出・選択系
'----------------------------
Public Function filter_by_mask(ByVal mask As Variant) As Table
    ensure_loaded

    If DATA.row_count = 0 Then
        Set filter_by_mask = clone_empty_table()
        Exit Function
    End If

    Dim result As Table
    Set result = New Table
    result.read_matrix DATA.filter_rows(mask), clone_column_names()
    Set filter_by_mask = result
End Function

Public Function filter_by_equals(ByVal columnName As String, ByVal matchValue As Variant) As Table
    ensure_loaded
    If DATA.row_count = 0 Then
        Set filter_by_equals = clone_empty_table()
        Exit Function
    End If
    Set filter_by_equals = Me.filter_by_mask(build_mask_for_equals(resolve_column_index(columnName, "Table.filter_by_equals"), matchValue))
End Function

Public Function filter_by_in(ByVal columnName As String, ByVal matchValues As Variant) As Table
    ensure_loaded
    If DATA.row_count = 0 Then
        Set filter_by_in = clone_empty_table()
        Exit Function
    End If
    Set filter_by_in = Me.filter_by_mask(build_mask_for_in(resolve_column_index(columnName, "Table.filter_by_in"), matchValues))
End Function

Public Function filter_by_contains(ByVal columnName As String, ByVal searchText As String, Optional ByVal caseSensitive As Boolean = False) As Table
    ensure_loaded
    If DATA.row_count = 0 Then
        Set filter_by_contains = clone_empty_table()
        Exit Function
    End If
    Set filter_by_contains = Me.filter_by_mask(build_mask_for_contains(resolve_column_index(columnName, "Table.filter_by_contains"), searchText, caseSensitive))
End Function

Public Function filter_by_all_equals(ByVal columnNames As Variant, ByVal matchValues As Variant) As Table
    ensure_loaded

    If DATA.row_count = 0 Then
        Set filter_by_all_equals = clone_empty_table()
        Exit Function
    End If

    If Not IsArray(columnNames) Or Not IsArray(matchValues) Then
        Err.Raise vbObjectError + 3014, "Table.filter_by_all_equals", "columnNames と matchValues は一次元配列で指定してください。"
    End If

    Dim lb As Long, ub As Long, i As Long
    Dim mask As Variant

    lb = LBound(columnNames)
    ub = UBound(columnNames)
    If (UBound(matchValues) - LBound(matchValues)) <> (ub - lb) Then
        Err.Raise vbObjectError + 3015, "Table.filter_by_all_equals", "columnNames と matchValues の長さが一致しません。"
    End If

    mask = build_mask_for_equals(resolve_column_index(CStr(columnNames(lb)), "Table.filter_by_all_equals"), matchValues(LBound(matchValues)))
    For i = lb + 1 To ub
        mask = combine_masks(mask, build_mask_for_equals(resolve_column_index(CStr(columnNames(i)), "Table.filter_by_all_equals"), matchValues(LBound(matchValues) + (i - lb))), True)
    Next i

    Set filter_by_all_equals = Me.filter_by_mask(mask)
End Function

Public Function filter_by_any_equals(ByVal columnNames As Variant, ByVal matchValues As Variant) As Table
    ensure_loaded

    If DATA.row_count = 0 Then
        Set filter_by_any_equals = clone_empty_table()
        Exit Function
    End If

    If Not IsArray(columnNames) Or Not IsArray(matchValues) Then
        Err.Raise vbObjectError + 3016, "Table.filter_by_any_equals", "columnNames と matchValues は一次元配列で指定してください。"
    End If

    Dim lb As Long, ub As Long, i As Long
    Dim mask As Variant

    lb = LBound(columnNames)
    ub = UBound(columnNames)
    If (UBound(matchValues) - LBound(matchValues)) <> (ub - lb) Then
        Err.Raise vbObjectError + 3017, "Table.filter_by_any_equals", "columnNames と matchValues の長さが一致しません。"
    End If

    mask = build_mask_for_equals(resolve_column_index(CStr(columnNames(lb)), "Table.filter_by_any_equals"), matchValues(LBound(matchValues)))
    For i = lb + 1 To ub
        mask = combine_masks(mask, build_mask_for_equals(resolve_column_index(CStr(columnNames(i)), "Table.filter_by_any_equals"), matchValues(LBound(matchValues) + (i - lb))), False)
    Next i

    Set filter_by_any_equals = Me.filter_by_mask(mask)
End Function

Public Function select_columns(ByVal columnNames As Variant) As Table
    ensure_loaded

    If Not IsArray(columnNames) Then
        Err.Raise vbObjectError + 3010, "Table.select_columns", "columnNames は一次元配列で指定してください。"
    End If

    Dim lb As Long, ub As Long, i As Long
    Dim indexes As Variant
    Dim result As Table

    lb = LBound(columnNames)
    ub = UBound(columnNames)
    ReDim indexes(lb To ub)

    For i = lb To ub
        indexes(i) = resolve_column_index(CStr(columnNames(i)), "Table.select_columns")
    Next i

    Set result = New Table
    result.read_matrix DATA.select_columns(indexes), subset_column_names(indexes)
    Set select_columns = result
End Function

'----------------------------
' 更新・整形系
'----------------------------
Public Sub add_column(ByVal columnName As String, Optional ByVal values As Variant)
    ensure_loaded

    If Trim$(columnName) = "" Then
        Err.Raise vbObjectError + 3018, "Table.add_column", "columnName は空文字にできません。"
    End If
    If COLUMN_INDEX_MAP.Exists(columnName) Then
        Err.Raise vbObjectError + 3019, "Table.add_column", "既に存在する列名です: " & columnName
    End If

    If DATA.row_count > 0 And IsMissing(values) Then
        values = create_empty_column_values()
    End If

    DATA.append_column values

    Dim newNames As Variant
    Dim i As Long
    ReDim newNames(1 To UBound(COLUMN_NAMES) + 1)
    For i = 1 To UBound(COLUMN_NAMES)
        newNames(i) = COLUMN_NAMES(i)
    Next i
    newNames(UBound(newNames)) = columnName

    COLUMN_NAMES = newNames
    rebuild_column_index_map
End Sub

Public Sub rename_column(ByVal oldName As String, ByVal newName As String)
    ensure_loaded

    Dim colIndex As Long
    colIndex = resolve_column_index(oldName, "Table.rename_column")

    If Trim$(newName) = "" Then
        Err.Raise vbObjectError + 3021, "Table.rename_column", "newName は空文字にできません。"
    End If
    If StrComp(oldName, newName, vbTextCompare) <> 0 And COLUMN_INDEX_MAP.Exists(newName) Then
        Err.Raise vbObjectError + 3022, "Table.rename_column", "既に存在する列名です: " & newName
    End If

    COLUMN_NAMES(colIndex) = newName
    rebuild_column_index_map
End Sub

Public Sub sort_by(ByVal columnName As String, Optional ByVal ascending As Boolean = True)
    ensure_loaded

    If DATA.row_count <= 1 Then Exit Sub

    Dim colIndex As Long
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim shouldSwap As Boolean

    colIndex = resolve_column_index(columnName, "Table.sort_by")
    arr = DATA.data

    For i = 1 To DATA.row_count - 1
        For j = i + 1 To DATA.row_count
            shouldSwap = compare_order_safe(arr(i, colIndex), arr(j, colIndex), ascending)
            If shouldSwap Then swap_rows arr, i, j
        Next j
    Next i

    DATA.read_array arr
End Sub

Public Sub set_by_mask(ByVal mask As Variant, ByVal columnName As String, ByVal newValue As Variant)
    ensure_loaded

    If Not IsArray(mask) Then
        Err.Raise vbObjectError + 3011, "Table.set_by_mask", "mask は一次元配列で指定してください。"
    End If

    If DATA.row_count = 0 Then Exit Sub

    Dim colIndex As Long
    Dim values As Variant
    Dim i As Long

    colIndex = resolve_column_index(columnName, "Table.set_by_mask")
    values = DATA.col_values(colIndex)

    If (UBound(mask) - LBound(mask) + 1) <> DATA.row_count Then
        Err.Raise vbObjectError + 3012, "Table.set_by_mask", "mask の長さが行数と一致しません。"
    End If

    For i = 1 To DATA.row_count
        If CBool(mask(LBound(mask) + i - 1)) Then
            values(i) = newValue
        End If
    Next i

    DATA.set_column_values colIndex, values
End Sub

Public Sub set_by_equals(ByVal conditionColumnName As String, ByVal matchValue As Variant, ByVal targetColumnName As String, ByVal newValue As Variant)
    Me.set_by_mask build_mask_for_equals(resolve_column_index(conditionColumnName, "Table.set_by_equals"), matchValue), targetColumnName, newValue
End Sub

Public Sub set_column(ByVal columnName As String, ByVal values As Variant)
    ensure_loaded
    DATA.set_column_values resolve_column_index(columnName, "Table.set_column"), values
End Sub

'----------------------------
' 出力系
'----------------------------
Public Sub to_range(ByVal topLeft As Range, Optional ByVal includeHeader As Boolean = True)
    ensure_loaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 3020, "Table.to_range", "topLeft が Nothing です。"

    Dim outputArr As Variant
    Dim body As Variant
    Dim r As Long, c As Long

    If DATA.row_count = 0 Then
        If includeHeader Then
            ReDim outputArr(1 To 1, 1 To DATA.col_count)
            For c = 1 To DATA.col_count
                outputArr(1, c) = COLUMN_NAMES(c)
            Next c
            topLeft.Resize(1, DATA.col_count).Value = outputArr
        End If
        Exit Sub
    End If

    body = DATA.data

    If includeHeader Then
        ReDim outputArr(1 To DATA.row_count + 1, 1 To DATA.col_count)
        For c = 1 To DATA.col_count
            outputArr(1, c) = COLUMN_NAMES(c)
        Next c
        For r = 1 To DATA.row_count
            For c = 1 To DATA.col_count
                outputArr(r + 1, c) = body(r, c)
            Next c
        Next r
        topLeft.Resize(DATA.row_count + 1, DATA.col_count).Value = outputArr
    Else
        topLeft.Resize(DATA.row_count, DATA.col_count).Value = body
    End If
End Sub

'----------------------------
' 状態管理系
'----------------------------
Public Sub clear()
    Set DATA = Nothing
    COLUMN_NAMES = Empty
    Set COLUMN_INDEX_MAP = Nothing
    TABLE_IS_LOADED = False
End Sub
