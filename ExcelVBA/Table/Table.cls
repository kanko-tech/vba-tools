'----------------------------
'クラスモジュールのメタ属性
'----------------------------
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

    If hasHeader Then
        If rows < 2 Then
            Err.Raise vbObjectError + 3002, "Table.read_range", "ヘッダー付きで読み込む場合は2行以上必要です。"
        End If

        ReDim COLUMN_NAMES(1 To cols)
        For c = 1 To cols
            COLUMN_NAMES(c) = CStr(raw(1, c))
        Next c

        ReDim body(1 To rows - 1, 1 To cols)
        For r = 2 To rows
            For c = 1 To cols
                body(r - 1, c) = raw(r, c)
            Next c
        Next r
    Else
        ReDim COLUMN_NAMES(1 To cols)
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

    load_internal src.data, columnNames
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

Private Sub load_internal(ByVal dataArr As Variant, ByVal columnNames As Variant)
    If Not IsArray(columnNames) Then
        Err.Raise vbObjectError + 3005, "Table.load_internal", "columnNames は一次元配列で指定してください。"
    End If

    Dim colLb As Long, colUb As Long, i As Long
    colLb = LBound(columnNames)
    colUb = UBound(columnNames)

    If (colUb - colLb + 1) <> (UBound(dataArr, 2) - LBound(dataArr, 2) + 1) Then
        Err.Raise vbObjectError + 3006, "Table.load_internal", "列名数と列数が一致しません。"
    End If

    Set DATA = New Matrix
    DATA.read_array dataArr

    COLUMN_NAMES = normalize_column_names(columnNames)
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

Private Sub rebuild_column_index_map()
    Dim i As Long

    Set COLUMN_INDEX_MAP = CreateObject("Scripting.Dictionary")
    COLUMN_INDEX_MAP.CompareMode = 1

    For i = LBound(COLUMN_NAMES) To UBound(COLUMN_NAMES)
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
    Dim lb As Long, ub As Long, i As Long
    Dim mask As Variant

    values = DATA.col_values(colIndex)
    lb = LBound(values)
    ub = UBound(values)
    ReDim mask(lb To ub)

    For i = lb To ub
        mask(i) = (values(i) = matchValue)
    Next i

    build_mask_for_equals = mask
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
    column_names = COLUMN_NAMES
End Property

Public Property Get matrix() As Matrix
    ensure_loaded

    Dim result As Matrix
    Set result = New Matrix
    result.read_array DATA.data
    Set matrix = result
End Property

Public Function col(ByVal columnName As String) As Variant
    Dim colIndex As Long
    colIndex = resolve_column_index(columnName, "Table.col")
    col = DATA.col_values(colIndex)
End Function

'----------------------------
' pandas 風操作系
'----------------------------
Public Function filter_by_mask(ByVal mask As Variant) As Table
    ensure_loaded

    Dim result As Table
    Dim filtered As Matrix

    Set filtered = DATA.filter_rows(mask)
    Set result = New Table
    result.read_matrix filtered, COLUMN_NAMES
    Set filter_by_mask = result
End Function

Public Function filter_by_equals(ByVal columnName As String, ByVal matchValue As Variant) As Table
    ensure_loaded

    Dim colIndex As Long
    Dim mask As Variant

    colIndex = resolve_column_index(columnName, "Table.filter_by_equals")
    mask = build_mask_for_equals(colIndex, matchValue)

    Set filter_by_equals = Me.filter_by_mask(mask)
End Function

Public Function select_columns(ByVal columnNames As Variant) As Table
    ensure_loaded

    If Not IsArray(columnNames) Then
        Err.Raise vbObjectError + 3010, "Table.select_columns", "columnNames は一次元配列で指定してください。"
    End If

    Dim lb As Long, ub As Long, i As Long
    Dim indexes As Variant
    Dim result As Table
    Dim selected As Matrix

    lb = LBound(columnNames)
    ub = UBound(columnNames)
    ReDim indexes(lb To ub)

    For i = lb To ub
        indexes(i) = resolve_column_index(CStr(columnNames(i)), "Table.select_columns")
    Next i

    Set selected = DATA.select_columns(indexes)
    Set result = New Table
    result.read_matrix selected, subset_column_names(indexes)
    Set select_columns = result
End Function

Public Sub set_by_mask(ByVal mask As Variant, ByVal columnName As String, ByVal newValue As Variant)
    ensure_loaded

    If Not IsArray(mask) Then
        Err.Raise vbObjectError + 3011, "Table.set_by_mask", "mask は一次元配列で指定してください。"
    End If

    Dim colIndex As Long
    Dim values As Variant
    Dim lb As Long, ub As Long, i As Long

    colIndex = resolve_column_index(columnName, "Table.set_by_mask")
    values = DATA.col_values(colIndex)

    lb = LBound(values)
    ub = UBound(values)
    If (UBound(mask) - LBound(mask)) <> (ub - lb) Then
        Err.Raise vbObjectError + 3012, "Table.set_by_mask", "mask の長さが行数と一致しません。"
    End If

    For i = lb To ub
        If CBool(mask(LBound(mask) + (i - lb))) Then
            values(i) = newValue
        End If
    Next i

    DATA.set_column_values colIndex, values
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
    Dim r As Long, c As Long
    Dim rows As Long, cols As Long
    Dim body As Variant

    body = DATA.data
    rows = DATA.row_count
    cols = DATA.col_count

    If includeHeader Then
        ReDim outputArr(1 To rows + 1, 1 To cols)
        For c = 1 To cols
            outputArr(1, c) = COLUMN_NAMES(c)
        Next c
        For r = 1 To rows
            For c = 1 To cols
                outputArr(r + 1, c) = body(r, c)
            Next c
        Next r
        topLeft.Resize(rows + 1, cols).Value = outputArr
    Else
        topLeft.Resize(rows, cols).Value = body
    End If
End Sub

Public Sub clear()
    Set DATA = Nothing
    COLUMN_NAMES = Empty
    Set COLUMN_INDEX_MAP = Nothing
    TABLE_IS_LOADED = False
End Sub
