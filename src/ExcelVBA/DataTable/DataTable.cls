VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "DataTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------
' �v���p�e�B
'----------------------------
Private mData As Matrix
Private mColumnNames As Variant
Private mColumnIndexMap As Object
Private mIsLoaded As Boolean

'----------------------------
' �Ǎ��n
'----------------------------
Public Sub read_range(ByVal rng As Range, Optional ByVal hasHeader As Boolean = True)
    If rng Is Nothing Then Err.Raise vbObjectError + 3001, "DataTable.read_range", "����rng��Nothing�ł��B"

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

    ReDim mColumnNames(1 To cols)

    If hasHeader Then
        For c = 1 To cols
            mColumnNames(c) = CStr(raw(1, c))
        Next c

        If rows = 1 Then
            Set mData = New Matrix
            mData.read_empty cols
            rebuild_column_index_map
            mIsLoaded = True
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
            mColumnNames(c) = "col" & CStr(c)
        Next c

        ReDim body(1 To rows, 1 To cols)
        For r = 1 To rows
            For c = 1 To cols
                body(r, c) = raw(r, c)
            Next c
        Next r
    End If

    load_internal body, mColumnNames
End Sub

Public Sub read_matrix(ByVal src As Matrix, ByVal columnNames As Variant)
    If src Is Nothing Then
        Err.Raise vbObjectError + 3003, "DataTable.read_matrix", "����src��Nothing�ł��B"
    End If
    If Not IsArray(columnNames) Then
        Err.Raise vbObjectError + 3004, "DataTable.read_matrix", "columnNames �͈ꎟ���z��Ŏw�肵�Ă��������B"
    End If

    load_from_matrix src, columnNames
End Sub

'----------------------------
' �K�[�h�n
'----------------------------
Private Sub ensure_loaded(Optional ByVal caller As String = "")
    If Not mIsLoaded Then
        Dim msg As String
        msg = "�f�[�^���ǂݍ��܂�Ă��܂���Bread_range / read_matrix ���ɌĂяo���Ă��������B"

        If caller <> "" Then
            msg = msg & vbCrLf & "�Ăяo����: " & caller
        End If

        Err.Raise vbObjectError + 3099, "DataTable.ensure_loaded", msg
    End If
End Sub

'----------------------------
' ����⏕
'----------------------------
Private Sub load_internal(ByVal dataArr As Variant, ByVal columnNames As Variant)
    If Not IsArray(columnNames) Then
        Err.Raise vbObjectError + 3005, "DataTable.load_internal", "columnNames �͈ꎟ���z��Ŏw�肵�Ă��������B"
    End If

    Dim normalizedNames As Variant
    normalizedNames = normalize_column_names(columnNames)

    If (UBound(normalizedNames) - LBound(normalizedNames) + 1) <> (UBound(dataArr, 2) - LBound(dataArr, 2) + 1) Then
        Err.Raise vbObjectError + 3006, "DataTable.load_internal", "�񖼐��Ɨ񐔂���v���܂���B"
    End If

    Set mData = New Matrix
    mData.read_array dataArr

    mColumnNames = normalizedNames
    rebuild_column_index_map
    mIsLoaded = True
End Sub

Private Sub load_from_matrix(ByVal src As Matrix, ByVal columnNames As Variant)
    Dim normalizedNames As Variant
    normalizedNames = normalize_column_names(columnNames)

    If (UBound(normalizedNames) - LBound(normalizedNames) + 1) <> src.col_count Then
        Err.Raise vbObjectError + 3006, "DataTable.load_from_matrix", "�񖼐��Ɨ񐔂���v���܂���B"
    End If

    Set mData = New Matrix
    If src.row_count = 0 Then
        mData.read_empty src.col_count
    Else
        mData.read_array src.data
    End If

    mColumnNames = normalizedNames
    rebuild_column_index_map
    mIsLoaded = True
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

    ReDim names(1 To UBound(mColumnNames))
    For i = 1 To UBound(mColumnNames)
        names(i) = mColumnNames(i)
    Next i

    clone_column_names = names
End Function

Private Sub rebuild_column_index_map()
    Dim i As Long

    Set mColumnIndexMap = CreateObject("Scripting.Dictionary")
    mColumnIndexMap.CompareMode = 1

    For i = LBound(mColumnNames) To UBound(mColumnNames)
        If Trim$(CStr(mColumnNames(i))) = "" Then
            Err.Raise vbObjectError + 3009, "DataTable.rebuild_column_index_map", "��̗񖼂͎g�p�ł��܂���B"
        End If
        If mColumnIndexMap.Exists(CStr(mColumnNames(i))) Then
            Err.Raise vbObjectError + 3007, "DataTable.rebuild_column_index_map", "�񖼂��d�����Ă��܂�: " & CStr(mColumnNames(i))
        End If
        mColumnIndexMap.Add CStr(mColumnNames(i)), i
    Next i
End Sub

Private Function resolve_column_index(ByVal columnName As String, ByVal caller As String) As Long
    ensure_loaded caller

    If Not mColumnIndexMap.Exists(columnName) Then
        Err.Raise vbObjectError + 3008, caller, "�w�肳�ꂽ�񖼂����݂��܂���: " & columnName
    End If

    resolve_column_index = CLng(mColumnIndexMap(columnName))
End Function

Private Function build_mask_for_equals(ByVal colIndex As Long, ByVal matchValue As Variant) As Variant
    Dim values As Variant
    Dim i As Long
    Dim mask As Variant

    values = mData.col_values(colIndex)
    ReDim mask(1 To mData.row_count)

    For i = 1 To mData.row_count
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
        Err.Raise vbObjectError + 3013, "DataTable.build_mask_for_in", "matchValues �͈ꎟ���z��Ŏw�肵�Ă��������B"
    End If

    values = mData.col_values(colIndex)
    lb = LBound(matchValues)
    ub = UBound(matchValues)
    ReDim mask(1 To mData.row_count)

    For i = 1 To mData.row_count
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

    values = mData.col_values(colIndex)
    ReDim mask(1 To mData.row_count)

    needle = searchText
    If Not caseSensitive Then needle = LCase$(needle)

    For i = 1 To mData.row_count
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

    ReDim mask(1 To mData.row_count)
    For i = 1 To mData.row_count
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
        names(1 + (i - lb)) = mColumnNames(CLng(selectedIndexes(i)))
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

    ReDim values(1 To mData.row_count)
    For i = 1 To mData.row_count
        values(i) = Empty
    Next i

    create_empty_column_values = values
End Function

Private Function clone_empty_datatable() As DataTable
    Dim result As DataTable
    Dim emptyMatrix As Matrix

    Set result = New DataTable
    Set emptyMatrix = New Matrix
    emptyMatrix.read_empty mData.col_count
    result.read_matrix emptyMatrix, clone_column_names()
    Set clone_empty_datatable = result
End Function

'----------------------------
' �Q�ƌn
'----------------------------
Public Property Get is_loaded() As Boolean
    is_loaded = mIsLoaded
End Property

Public Property Get row_count() As Long
    ensure_loaded
    row_count = mData.row_count
End Property

Public Property Get col_count() As Long
    ensure_loaded
    col_count = mData.col_count
End Property

Public Property Get column_names() As Variant
    ensure_loaded
    column_names = clone_column_names()
End Property

Public Property Get matrix() As Matrix
    ensure_loaded

    Dim result As Matrix
    Set result = New Matrix
    If mData.row_count = 0 Then
        result.read_empty mData.col_count
    Else
        result.read_array mData.data
    End If
    Set matrix = result
End Property

Public Function col(ByVal columnName As String) As Variant
    col = mData.col_values(resolve_column_index(columnName, "DataTable.col"))
End Function

Public Function col_vector(ByVal columnName As String) As Vector
    Dim vec As Vector
    Set vec = New Vector
    vec.read_array Me.col(columnName)
    Set col_vector = vec
End Function

'----------------------------
' ���o�E�I��n
'----------------------------
Public Function filter_by_mask(ByVal mask As Variant) As DataTable
    ensure_loaded

    If mData.row_count = 0 Then
        Set filter_by_mask = clone_empty_datatable()
        Exit Function
    End If

    Dim result As DataTable
    Set result = New DataTable
    result.read_matrix mData.filter_rows(mask), clone_column_names()
    Set filter_by_mask = result
End Function

Public Function filter_by_equals(ByVal columnName As String, ByVal matchValue As Variant) As DataTable
    ensure_loaded
    If mData.row_count = 0 Then
        Set filter_by_equals = clone_empty_datatable()
        Exit Function
    End If
    Set filter_by_equals = Me.filter_by_mask(build_mask_for_equals(resolve_column_index(columnName, "DataTable.filter_by_equals"), matchValue))
End Function

Public Function filter_by_in(ByVal columnName As String, ByVal matchValues As Variant) As DataTable
    ensure_loaded
    If mData.row_count = 0 Then
        Set filter_by_in = clone_empty_datatable()
        Exit Function
    End If
    Set filter_by_in = Me.filter_by_mask(build_mask_for_in(resolve_column_index(columnName, "DataTable.filter_by_in"), matchValues))
End Function

Public Function filter_by_contains(ByVal columnName As String, ByVal searchText As String, Optional ByVal caseSensitive As Boolean = False) As DataTable
    ensure_loaded
    If mData.row_count = 0 Then
        Set filter_by_contains = clone_empty_datatable()
        Exit Function
    End If
    Set filter_by_contains = Me.filter_by_mask(build_mask_for_contains(resolve_column_index(columnName, "DataTable.filter_by_contains"), searchText, caseSensitive))
End Function

Public Function filter_by_all_equals(ByVal columnNames As Variant, ByVal matchValues As Variant) As DataTable
    ensure_loaded

    If mData.row_count = 0 Then
        Set filter_by_all_equals = clone_empty_datatable()
        Exit Function
    End If

    If Not IsArray(columnNames) Or Not IsArray(matchValues) Then
        Err.Raise vbObjectError + 3014, "DataTable.filter_by_all_equals", "columnNames �� matchValues �͈ꎟ���z��Ŏw�肵�Ă��������B"
    End If

    Dim lb As Long, ub As Long, i As Long
    Dim mask As Variant

    lb = LBound(columnNames)
    ub = UBound(columnNames)
    If (UBound(matchValues) - LBound(matchValues)) <> (ub - lb) Then
        Err.Raise vbObjectError + 3015, "DataTable.filter_by_all_equals", "columnNames �� matchValues �̒�������v���܂���B"
    End If

    mask = build_mask_for_equals(resolve_column_index(CStr(columnNames(lb)), "DataTable.filter_by_all_equals"), matchValues(LBound(matchValues)))
    For i = lb + 1 To ub
        mask = combine_masks(mask, build_mask_for_equals(resolve_column_index(CStr(columnNames(i)), "DataTable.filter_by_all_equals"), matchValues(LBound(matchValues) + (i - lb))), True)
    Next i

    Set filter_by_all_equals = Me.filter_by_mask(mask)
End Function

Public Function filter_by_any_equals(ByVal columnNames As Variant, ByVal matchValues As Variant) As DataTable
    ensure_loaded

    If mData.row_count = 0 Then
        Set filter_by_any_equals = clone_empty_datatable()
        Exit Function
    End If

    If Not IsArray(columnNames) Or Not IsArray(matchValues) Then
        Err.Raise vbObjectError + 3016, "DataTable.filter_by_any_equals", "columnNames �� matchValues �͈ꎟ���z��Ŏw�肵�Ă��������B"
    End If

    Dim lb As Long, ub As Long, i As Long
    Dim mask As Variant

    lb = LBound(columnNames)
    ub = UBound(columnNames)
    If (UBound(matchValues) - LBound(matchValues)) <> (ub - lb) Then
        Err.Raise vbObjectError + 3017, "DataTable.filter_by_any_equals", "columnNames �� matchValues �̒�������v���܂���B"
    End If

    mask = build_mask_for_equals(resolve_column_index(CStr(columnNames(lb)), "DataTable.filter_by_any_equals"), matchValues(LBound(matchValues)))
    For i = lb + 1 To ub
        mask = combine_masks(mask, build_mask_for_equals(resolve_column_index(CStr(columnNames(i)), "DataTable.filter_by_any_equals"), matchValues(LBound(matchValues) + (i - lb))), False)
    Next i

    Set filter_by_any_equals = Me.filter_by_mask(mask)
End Function

Public Function select_columns(ByVal columnNames As Variant) As DataTable
    ensure_loaded

    If Not IsArray(columnNames) Then
        Err.Raise vbObjectError + 3010, "DataTable.select_columns", "columnNames �͈ꎟ���z��Ŏw�肵�Ă��������B"
    End If

    Dim lb As Long, ub As Long, i As Long
    Dim indexes As Variant
    Dim result As DataTable

    lb = LBound(columnNames)
    ub = UBound(columnNames)
    ReDim indexes(lb To ub)

    For i = lb To ub
        indexes(i) = resolve_column_index(CStr(columnNames(i)), "DataTable.select_columns")
    Next i

    Set result = New DataTable
    result.read_matrix mData.select_columns(indexes), subset_column_names(indexes)
    Set select_columns = result
End Function

'----------------------------
' �X�V�E���`�n
'----------------------------
Public Sub add_column(ByVal columnName As String, Optional ByVal values As Variant)
    ensure_loaded

    If Trim$(columnName) = "" Then
        Err.Raise vbObjectError + 3018, "DataTable.add_column", "columnName �͋󕶎��ɂł��܂���B"
    End If
    If mColumnIndexMap.Exists(columnName) Then
        Err.Raise vbObjectError + 3019, "DataTable.add_column", "���ɑ��݂���񖼂ł�: " & columnName
    End If

    If mData.row_count > 0 And IsMissing(values) Then
        values = create_empty_column_values()
    End If

    mData.append_column values

    Dim newNames As Variant
    Dim i As Long
    ReDim newNames(1 To UBound(mColumnNames) + 1)
    For i = 1 To UBound(mColumnNames)
        newNames(i) = mColumnNames(i)
    Next i
    newNames(UBound(newNames)) = columnName

    mColumnNames = newNames
    rebuild_column_index_map
End Sub

Public Sub rename_column(ByVal oldName As String, ByVal newName As String)
    ensure_loaded

    Dim colIndex As Long
    colIndex = resolve_column_index(oldName, "DataTable.rename_column")

    If Trim$(newName) = "" Then
        Err.Raise vbObjectError + 3021, "DataTable.rename_column", "newName �͋󕶎��ɂł��܂���B"
    End If
    If StrComp(oldName, newName, vbTextCompare) <> 0 And mColumnIndexMap.Exists(newName) Then
        Err.Raise vbObjectError + 3022, "DataTable.rename_column", "���ɑ��݂���񖼂ł�: " & newName
    End If

    mColumnNames(colIndex) = newName
    rebuild_column_index_map
End Sub

Public Sub sort_by(ByVal columnName As String, Optional ByVal ascending As Boolean = True)
    ensure_loaded

    If mData.row_count <= 1 Then Exit Sub

    Dim colIndex As Long
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim shouldSwap As Boolean

    colIndex = resolve_column_index(columnName, "DataTable.sort_by")
    arr = mData.data

    For i = 1 To mData.row_count - 1
        For j = i + 1 To mData.row_count
            shouldSwap = compare_order_safe(arr(i, colIndex), arr(j, colIndex), ascending)
            If shouldSwap Then swap_rows arr, i, j
        Next j
    Next i

    mData.read_array arr
End Sub

Public Sub set_by_mask(ByVal mask As Variant, ByVal columnName As String, ByVal newValue As Variant)
    ensure_loaded

    If Not IsArray(mask) Then
        Err.Raise vbObjectError + 3011, "DataTable.set_by_mask", "mask �͈ꎟ���z��Ŏw�肵�Ă��������B"
    End If

    If mData.row_count = 0 Then Exit Sub

    Dim colIndex As Long
    Dim values As Variant
    Dim i As Long

    colIndex = resolve_column_index(columnName, "DataTable.set_by_mask")
    values = mData.col_values(colIndex)

    If (UBound(mask) - LBound(mask) + 1) <> mData.row_count Then
        Err.Raise vbObjectError + 3012, "DataTable.set_by_mask", "mask �̒������s���ƈ�v���܂���B"
    End If

    For i = 1 To mData.row_count
        If CBool(mask(LBound(mask) + i - 1)) Then
            values(i) = newValue
        End If
    Next i

    mData.set_column_values colIndex, values
End Sub

Public Sub set_by_equals(ByVal conditionColumnName As String, ByVal matchValue As Variant, ByVal targetColumnName As String, ByVal newValue As Variant)
    Me.set_by_mask build_mask_for_equals(resolve_column_index(conditionColumnName, "DataTable.set_by_equals"), matchValue), targetColumnName, newValue
End Sub

Public Sub set_column(ByVal columnName As String, ByVal values As Variant)
    ensure_loaded
    mData.set_column_values resolve_column_index(columnName, "DataTable.set_column"), values
End Sub

'----------------------------
' �o�͌n
'----------------------------
Public Sub to_range(ByVal topLeft As Range, Optional ByVal includeHeader As Boolean = True)
    ensure_loaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 3020, "DataTable.to_range", "topLeft �� Nothing �ł��B"

    Dim outputArr As Variant
    Dim body As Variant
    Dim r As Long, c As Long

    If mData.row_count = 0 Then
        If includeHeader Then
            ReDim outputArr(1 To 1, 1 To mData.col_count)
            For c = 1 To mData.col_count
                outputArr(1, c) = mColumnNames(c)
            Next c
            topLeft.Resize(1, mData.col_count).Value = outputArr
        End If
        Exit Sub
    End If

    body = mData.data

    If includeHeader Then
        ReDim outputArr(1 To mData.row_count + 1, 1 To mData.col_count)
        For c = 1 To mData.col_count
            outputArr(1, c) = mColumnNames(c)
        Next c
        For r = 1 To mData.row_count
            For c = 1 To mData.col_count
                outputArr(r + 1, c) = body(r, c)
            Next c
        Next r
        topLeft.Resize(mData.row_count + 1, mData.col_count).Value = outputArr
    Else
        topLeft.Resize(mData.row_count, mData.col_count).Value = body
    End If
End Sub

'----------------------------
' ��ԊǗ��n
'----------------------------
Public Sub clear()
    Set mData = Nothing
    mColumnNames = Empty
    Set mColumnIndexMap = Nothing
    mIsLoaded = False
End Sub
