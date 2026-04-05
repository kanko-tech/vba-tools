'----------------------------
'クラスモジュールのメタ属性
'----------------------------
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1                         'クラスのインスタンスを複数生成できるか（-1：True（可能）、0:False（不可））
END
Attribute VB_Name = "Vector"            'クラス名定義
Attribute VB_GlobalNameSpace = False    'グローバル名前空間に公開するか（通常はFalse）
Attribute VB_Creatable = False          '外部（他アプリ）から CreateObject で生成できるか（False：外部から生成不可（通常）、True：COMとして公開する場合のみ使用）  
Attribute VB_PredeclaredId = False      'Newせずとも使用可とするか（Trueにするとシングルトン的挙動に。通常はFalse）
Attribute VB_Exposed = False            'クラスを外部に公開するか（通常はFalse）
Option Explicit                         '型宣言を強制

'----------------------------
' プロパティ
'----------------------------
Private VECTOR As Variant               '
Private VECTOR_IS_LOADED As Boolean     '

'----------------------------
' 読込系
'----------------------------
Public Sub read_col_range(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 1001, "Vector.read_col_range", "引数rngがNothingです。"
    
    If rng.Columns.count <> 1 Then
        Err.Raise vbObjectError + 1002, "Vector.read_col_range", "1列のRangeのみ指定してください。"
    End If

    If rng.Cells.count = 1 Then
        Dim tmp1(1 To 1) As Variant
        tmp1(1) = rng.value
        VECTOR = tmp1
    Else
        VECTOR = Application.index(rng.value, 0, 1)
    End If
    
    VECTOR_IS_LOADED = True
End Sub

Public Sub read_row_range(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 1010, "Vector.read_row_range", "引数rngがNothingです。"
    
    If rng.Rows.count <> 1 Then
        Err.Raise vbObjectError + 1012, "Vector.read_row_range", "1行のRangeのみ指定してください。"
    End If

    If rng.Cells.count = 1 Then
        Dim tmp1(1 To 1) As Variant
        tmp1(1) = rng.value
        VECTOR = tmp1
    Else
        VECTOR = Application.index(rng.value, 1, 0)
    End If
    
    VECTOR_IS_LOADED = True
End Sub

Public Sub read_array(ByVal arr As Variant)
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 1030, "Vector.read_array", "配列ではありません。"
    End If
    
    ' 1次元配列チェック
    On Error Resume Next    '以後、エラーが起きてもその場で停止せず、次の行へ進む
    Dim t As Long
    t = UBound(arr, 2)      'UBound(arr, 2) が失敗するかどうかを使って、次元数を判定（2次元以上なら UBound(arr, 2) が通ってしまう）
    If Err.Number = 0 Then  '成功(0)なら2次元以上のためエラー
        Err.clear
        On Error GoTo 0
        Err.Raise vbObjectError + 1031, "Vector.read_array", "2次元以上の配列は受け付けません。"
    End If
    Err.clear
    On Error GoTo 0
    
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
        Err.Raise vbObjectError + 1099, _
                  "Vector.ensure_loaded", _
                  msg
    End If
End Sub

Private Sub ensure_strict_numeric_vector()
    Dim i As Long
    Dim vt As VbVarType
    Dim msg As String
    
    For i = LBound(VECTOR) To UBound(VECTOR)
        If IsEmpty(VECTOR(i)) Then
            Err.Raise vbObjectError + 1111, _
                      "Vector.ensure_strict_numeric_vector", _
                      "VECTOR に Empty が含まれています。添字: " & CStr(i)
        End If
        If IsError(VECTOR(i)) Then
            Err.Raise vbObjectError + 1112, _
                      "Vector.ensure_strict_numeric_vector", _
                      "VECTOR に Excel のエラー値が含まれています。添字: " & CStr(i)
        End If
        If IsNull(VECTOR(i)) Then
            Err.Raise vbObjectError + 1113, _
                      "Vector.ensure_strict_numeric_vector", _
                      "VECTOR に Null が含まれています。添字: " & CStr(i)
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
                
                Err.Raise vbObjectError + 1114, _
                          "Vector.ensure_strict_numeric_vector", _
                          msg
        End Select
    Next i
End Sub

'----------------------------
' 参照系
'----------------------------
Public Property Get is_loaded() As Boolean
    is_loaded = VECTOR_IS_LOADED
End Property

Public Property Get data() As Variant
    ensure_loaded
    data = VECTOR
End Property

Public Property Get count() As Long
    ensure_loaded
    count = UBound(VECTOR) - LBound(VECTOR) + 1
End Property

Public Property Get item(ByVal index As Long) As Variant
    ensure_loaded
    If index < LBound(VECTOR) Or index > UBound(VECTOR) Then
        Err.Raise vbObjectError + 1005, "Vector.item", "index ���͈͊O�ł��B"
    End If
    item = VECTOR(index)
End Property

Public Function type_names() As Variant
    ensure_loaded
    
    Dim lb As Long, ub As Long, i As Long
    lb = LBound(VECTOR)
    ub = UBound(VECTOR)
    
    Dim res() As Variant
    ReDim res(lb To ub)
    
    For i = lb To ub
        If IsError(VECTOR(i)) Then
            res(i) = "Error"
        ElseIf IsEmpty(VECTOR(i)) Then
            res(i) = "Empty"
        Else
            res(i) = TypeName(VECTOR(i))  ' "Double" "String" "Date" �Ȃ�
        End If
    Next
    
    type_names = res
End Function

'----------------------------
' 型系
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
                Err.Raise vbObjectError + 1201, "Vector.cast_to_string_safe", "Error�l���܂܂�Ă��܂��B"
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
                Err.Raise vbObjectError + 1202, "Vector.cast_to_date_safe", "���t�ɕϊ��ł��Ȃ��l���܂܂�Ă��܂��B"
            End If
        End If
        
    Next i
End Sub

'----------------------------
' 出力系
'----------------------------
Public Sub to_range_vertical(ByVal topLeft As Range)
    ensure_loaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 1006, "Vector.to_range_vertical", "topLeft �� Nothing �ł��B"
    
    Dim n As Long: n = Me.count
    topLeft.Resize(n, 1).value = Application.WorksheetFunction.Transpose(VECTOR)
End Sub

Public Sub to_range_horizontal(ByVal topLeft As Range)
    ensure_loaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 1007, "Vector.to_range_horizontal", "topLeft �� Nothing �ł��B"
    
    Dim n As Long: n = Me.count
    topLeft.Resize(1, n).value = VECTOR
End Sub

Public Sub clear()
    VECTOR = Empty
    VECTOR_IS_LOADED = False
End Sub
