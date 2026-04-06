VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'----------------------------
' プロパティ
'----------------------------
Private VECTOR As Variant
Private VECTOR_IS_LOADED As Boolean

'----------------------------
' 読込系
'----------------------------
Public Sub read_col_range(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 1001, "Vector.read_col_range", "引数rngがNothingです。"

    If rng.Columns.Count <> 1 Then
        Err.Raise vbObjectError + 1002, "Vector.read_col_range", "1列のRangeのみ指定してください。"
    End If

    If rng.Cells.Count = 1 Then
        Dim tmp1(1 To 1) As Variant
        tmp1(1) = rng.Value
        VECTOR = tmp1
    Else
        VECTOR = Application.Index(rng.Value, 0, 1)
    End If

    VECTOR_IS_LOADED = True
End Sub

Public Sub read_row_range(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 1010, "Vector.read_row_range", "引数rngがNothingです。"

    If rng.Rows.Count <> 1 Then
        Err.Raise vbObjectError + 1012, "Vector.read_row_range", "1行のRangeのみ指定してください。"
    End If

    If rng.Cells.Count = 1 Then
        Dim tmp2(1 To 1) As Variant
        tmp2(1) = rng.Value
        VECTOR = tmp2
    Else
        VECTOR = Application.Index(rng.Value, 1, 0)
    End If

    VECTOR_IS_LOADED = True
End Sub

Public Sub read_array(ByVal arr As Variant)
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 1030, "Vector.read_array", "配列ではありません。"
    End If

    ensure_is_one_dimensional arr

    VECTOR = arr
    VECTOR_IS_LOADED = True
End Sub

'----------------------------
' ガード（内部前提）系
'----------------------------
Private Sub ensure_loaded(Optional ByVal caller As String = "")
    If Not VECTOR_IS_LOADED Then
        Dim msg As String
        msg = "データが読み込まれていません。read_col_range / read_row_range / read_array を先に呼び出してください。"

        If caller <> "" Then
            msg = msg & vbCrLf & "呼び出し元: " & caller
        End If
        Err.Raise vbObjectError + 1099, "Vector.ensure_loaded", msg
    End If
End Sub

Private Sub ensure_is_one_dimensional(ByVal arr As Variant)
    On Error Resume Next

    Dim lb1 As Long
    Dim lb2 As Long

    lb1 = LBound(arr, 1)

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 1032, "Vector.ensure_is_one_dimensional", "一次元配列を指定してください。"
    End If

    lb2 = LBound(arr, 2)
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 1031, "Vector.ensure_is_one_dimensional", "2次元以上の配列は受け付けません。"
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Private Sub ensure_strict_numeric_vector()
    Dim i As Long
    Dim vt As VbVarType
    Dim msg As String

    For i = LBound(VECTOR) To UBound(VECTOR)
        If IsEmpty(VECTOR(i)) Then
            Err.Raise vbObjectError + 1111, "Vector.ensure_strict_numeric_vector", "VECTOR に Empty が含まれています。添字: " & CStr(i)
        End If
        If IsError(VECTOR(i)) Then
            Err.Raise vbObjectError + 1112, "Vector.ensure_strict_numeric_vector", "VECTOR に Excel のエラー値が含まれています。添字: " & CStr(i)
        End If
        If IsNull(VECTOR(i)) Then
            Err.Raise vbObjectError + 1113, "Vector.ensure_strict_numeric_vector", "VECTOR に Null が含まれています。添字: " & CStr(i)
        End If

        vt = VarType(VECTOR(i))

        Select Case vt
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                ' 許可
            Case Else
                msg = "VECTOR に数値型でない要素が含まれています。" & _
                      " 添字: " & CStr(i) & _
                      " / 値: " & CStr(VECTOR(i)) & _
                      " / VarType: " & CStr(vt)
                Err.Raise vbObjectError + 1114, "Vector.ensure_strict_numeric_vector", msg
        End Select
    Next i
End Sub

'----------------------------
' 内部補助
'----------------------------
Private Function clone_vector_data(ByVal src As Variant) As Variant
    Dim lb As Long, ub As Long, i As Long
    Dim dst As Variant

    lb = LBound(src)
    ub = UBound(src)
    ReDim dst(lb To ub)

    For i = lb To ub
        dst(i) = src(i)
    Next i

    clone_vector_data = dst
End Function

Private Function compare_relational(ByVal leftValue As Variant, ByVal rightValue As Variant, ByVal operatorName As String) As Boolean
    On Error Resume Next

    Select Case operatorName
        Case "gt"
            compare_relational = (leftValue > rightValue)
        Case "ge"
            compare_relational = (leftValue >= rightValue)
        Case "lt"
            compare_relational = (leftValue < rightValue)
        Case "le"
            compare_relational = (leftValue <= rightValue)
        Case Else
            compare_relational = False
    End Select

    If Err.Number <> 0 Then
        Err.Clear
        compare_relational = False
    End If

    On Error GoTo 0
End Function

Private Function build_boolean_mask_equals(ByVal matchValue As Variant, ByVal negate As Boolean) As Variant
    Dim lb As Long, ub As Long, i As Long
    Dim mask As Variant

    ensure_loaded

    lb = LBound(VECTOR)
    ub = UBound(VECTOR)
    ReDim mask(lb To ub)

    For i = lb To ub
        If IsError(VECTOR(i)) Or IsError(matchValue) Or IsNull(VECTOR(i)) Or IsNull(matchValue) Then
            mask(i) = False
        Else
            On Error Resume Next
            mask(i) = (VECTOR(i) = matchValue)
            If Err.Number <> 0 Then
                Err.Clear
                mask(i) = False
            End If
            On Error GoTo 0
            If negate Then mask(i) = Not CBool(mask(i))
        End If
    Next i

    build_boolean_mask_equals = mask
End Function

Private Function build_boolean_mask_relational(ByVal threshold As Variant, ByVal operatorName As String) As Variant
    Dim lb As Long, ub As Long, i As Long
    Dim mask As Variant

    ensure_loaded

    lb = LBound(VECTOR)
    ub = UBound(VECTOR)
    ReDim mask(lb To ub)

    For i = lb To ub
        If IsError(VECTOR(i)) Or IsEmpty(VECTOR(i)) Or IsNull(VECTOR(i)) Then
            mask(i) = False
        Else
            mask(i) = compare_relational(VECTOR(i), threshold, operatorName)
        End If
    Next i

    build_boolean_mask_relational = mask
End Function

Private Function build_variant_key(ByVal value As Variant) As String
    If IsError(value) Then
        build_variant_key = "Error:" & CStr(CLng(value))
    ElseIf IsEmpty(value) Then
        build_variant_key = "Empty:"
    ElseIf IsNull(value) Then
        build_variant_key = "Null:"
    Else
        build_variant_key = TypeName(value) & ":" & CStr(value)
    End If
End Function

'----------------------------
' 参照系
'----------------------------
Public Property Get is_loaded() As Boolean
    is_loaded = VECTOR_IS_LOADED
End Property

Public Property Get data() As Variant
    ensure_loaded
    data = clone_vector_data(VECTOR)
End Property

Public Property Get count() As Long
    ensure_loaded
    count = UBound(VECTOR) - LBound(VECTOR) + 1
End Property

Public Property Get item(ByVal index As Long) As Variant
    ensure_loaded
    If index < LBound(VECTOR) Or index > UBound(VECTOR) Then
        Err.Raise vbObjectError + 1005, "Vector.item", "index が範囲外です。"
    End If
    item = VECTOR(index)
End Property

Public Function type_names() As Variant
    ensure_loaded

    Dim lb As Long, ub As Long, i As Long
    Dim res As Variant

    lb = LBound(VECTOR)
    ub = UBound(VECTOR)
    ReDim res(lb To ub)

    For i = lb To ub
        If IsError(VECTOR(i)) Then
            res(i) = "Error"
        ElseIf IsEmpty(VECTOR(i)) Then
            res(i) = "Empty"
        Else
            res(i) = TypeName(VECTOR(i))
        End If
    Next i

    type_names = res
End Function

'----------------------------
' 条件マスク系
'----------------------------
Public Function eq(ByVal matchValue As Variant) As Variant
    eq = build_boolean_mask_equals(matchValue, False)
End Function

Public Function ne(ByVal matchValue As Variant) As Variant
    ne = build_boolean_mask_equals(matchValue, True)
End Function

Public Function gt(ByVal threshold As Variant) As Variant
    gt = build_boolean_mask_relational(threshold, "gt")
End Function

Public Function ge(ByVal threshold As Variant) As Variant
    ge = build_boolean_mask_relational(threshold, "ge")
End Function

Public Function lt(ByVal threshold As Variant) As Variant
    lt = build_boolean_mask_relational(threshold, "lt")
End Function

Public Function le(ByVal threshold As Variant) As Variant
    le = build_boolean_mask_relational(threshold, "le")
End Function

Public Function is_empty(Optional ByVal treatBlankStringAsEmpty As Boolean = True) As Variant
    ensure_loaded

    Dim lb As Long, ub As Long, i As Long
    Dim mask As Variant

    lb = LBound(VECTOR)
    ub = UBound(VECTOR)
    ReDim mask(lb To ub)

    For i = lb To ub
        mask(i) = IsEmpty(VECTOR(i))
        If treatBlankStringAsEmpty And VarType(VECTOR(i)) = vbString Then
            mask(i) = CBool(mask(i)) Or (Trim$(CStr(VECTOR(i))) = "")
        End If
    Next i

    is_empty = mask
End Function

'----------------------------
' 変換・更新系
'----------------------------
Public Sub cast_to_double_safe(Optional ByVal emptyAsZero As Boolean = False, _
                               Optional ByVal invalidAsZero As Boolean = False, _
                               Optional ByVal treatDateAsInvalid As Boolean = True)
    ensure_loaded

    Dim i As Long
    For i = LBound(VECTOR) To UBound(VECTOR)
        If IsError(VECTOR(i)) Or IsEmpty(VECTOR(i)) Then
            VECTOR(i) = IIf(emptyAsZero, 0#, Empty)
        ElseIf VarType(VECTOR(i)) = vbString And Trim$(CStr(VECTOR(i))) = "" Then
            VECTOR(i) = IIf(emptyAsZero, 0#, Empty)
        ElseIf IsDate(VECTOR(i)) And treatDateAsInvalid Then
            VECTOR(i) = IIf(invalidAsZero, 0#, Empty)
        ElseIf IsNumeric(VECTOR(i)) Then
            VECTOR(i) = CDbl(VECTOR(i))
        Else
            VECTOR(i) = IIf(invalidAsZero, 0#, Empty)
        End If
    Next i
End Sub

Public Sub cast_to_string_safe(Optional ByVal emptyAsBlank As Boolean = True, _
                               Optional ByVal errorAsBlank As Boolean = True)
    ensure_loaded

    Dim i As Long
    For i = LBound(VECTOR) To UBound(VECTOR)
        If IsError(VECTOR(i)) Then
            If errorAsBlank Then
                VECTOR(i) = ""
            Else
                Err.Raise vbObjectError + 1201, "Vector.cast_to_string_safe", "Error値が含まれています。"
            End If
        ElseIf IsEmpty(VECTOR(i)) Then
            VECTOR(i) = IIf(emptyAsBlank, "", Empty)
        Else
            VECTOR(i) = CStr(VECTOR(i))
        End If
    Next i
End Sub

Public Sub cast_to_date_safe(Optional ByVal invalidAsEmpty As Boolean = True)
    ensure_loaded

    Dim i As Long
    For i = LBound(VECTOR) To UBound(VECTOR)
        If IsError(VECTOR(i)) Or IsEmpty(VECTOR(i)) Then
            VECTOR(i) = Empty
        ElseIf VarType(VECTOR(i)) = vbString And Trim$(CStr(VECTOR(i))) = "" Then
            VECTOR(i) = Empty
        ElseIf IsDate(VECTOR(i)) Then
            VECTOR(i) = CDate(VECTOR(i))
        Else
            If invalidAsEmpty Then
                VECTOR(i) = Empty
            Else
                Err.Raise vbObjectError + 1202, "Vector.cast_to_date_safe", "日付に変換できない値が含まれています。"
            End If
        End If
    Next i
End Sub

Public Sub fill_empty(ByVal fillValue As Variant, Optional ByVal treatBlankStringAsEmpty As Boolean = True)
    ensure_loaded

    Dim i As Long
    For i = LBound(VECTOR) To UBound(VECTOR)
        If IsEmpty(VECTOR(i)) Then
            VECTOR(i) = fillValue
        ElseIf treatBlankStringAsEmpty And VarType(VECTOR(i)) = vbString And Trim$(CStr(VECTOR(i))) = "" Then
            VECTOR(i) = fillValue
        End If
    Next i
End Sub

Public Sub map(ByVal functionName As String)
    ensure_loaded

    If Trim$(functionName) = "" Then
        Err.Raise vbObjectError + 1203, "Vector.map", "functionName は空文字にできません。"
    End If

    Dim i As Long
    On Error GoTo MapFailed
    For i = LBound(VECTOR) To UBound(VECTOR)
        VECTOR(i) = Application.Run(functionName, VECTOR(i), i)
    Next i
    On Error GoTo 0
    Exit Sub

MapFailed:
    Err.Raise vbObjectError + 1204, "Vector.map", "指定した関数の実行に失敗しました: " & functionName
End Sub

'----------------------------
' 集計系
'----------------------------
Public Function unique() As Variant
    ensure_loaded

    Dim dict As Object
    Dim ordered As Collection
    Dim i As Long
    Dim key As String
    Dim res As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    Set ordered = New Collection

    For i = LBound(VECTOR) To UBound(VECTOR)
        key = build_variant_key(VECTOR(i))
        If Not dict.Exists(key) Then
            dict.Add key, True
            ordered.Add VECTOR(i)
        End If
    Next i

    ReDim res(1 To ordered.Count)
    For i = 1 To ordered.Count
        res(i) = ordered(i)
    Next i

    unique = res
End Function

Public Function sum() As Double
    ensure_loaded
    ensure_strict_numeric_vector

    Dim i As Long
    Dim total As Double

    total = 0#
    For i = LBound(VECTOR) To UBound(VECTOR)
        total = total + CDbl(VECTOR(i))
    Next i

    sum = total
End Function

Public Function mean() As Double
    ensure_loaded
    ensure_strict_numeric_vector
    mean = Me.sum / CDbl(Me.count)
End Function

'----------------------------
' 出力系
'----------------------------
Public Sub to_range_vertical(ByVal topLeft As Range)
    ensure_loaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 1006, "Vector.to_range_vertical", "topLeft が Nothing です。"

    Dim n As Long
    n = Me.count
    topLeft.Resize(n, 1).Value = Application.WorksheetFunction.Transpose(VECTOR)
End Sub

Public Sub to_range_horizontal(ByVal topLeft As Range)
    ensure_loaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 1007, "Vector.to_range_horizontal", "topLeft が Nothing です。"

    Dim n As Long
    n = Me.count
    topLeft.Resize(1, n).Value = VECTOR
End Sub

'----------------------------
' 状態管理系
'----------------------------
Public Sub clear()
    VECTOR = Empty
    VECTOR_IS_LOADED = False
End Sub
