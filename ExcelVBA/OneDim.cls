VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneDim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mOneDim As Variant  '一次元配列本体
Private mLoaded As Boolean    'mOneDimが正常に読み込まれたフラグ

'----------------------------
' 読み込み系
'----------------------------
Public Sub Read_ColRange(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 1001, "Araya1.Read_ColRange", "rng が Nothing です。"
    
    If rng.Cells.Count = 1 Then
        Dim tmp1(1 To 1) As Variant
        tmp1(1) = rng.value
        mOneDim = tmp1
    Else
        ' 縦Range → 1次元配列
        mOneDim = Application.index(rng.value, 0, 1)
    End If
    
    mLoaded = True
End Sub

Public Sub Read_RowRange(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 1002, "Araya1.Read_RowRange", "rng が Nothing です。"
    
    If rng.Cells.Count = 1 Then
        Dim tmp1(1 To 1) As Variant
        tmp1(1) = rng.value
        mOneDim = tmp1
    Else
        ' 横Range → 1次元配列
        mOneDim = Application.index(rng.value, 1, 0)
    End If
    
    mLoaded = True
End Sub

Public Sub Read_Array(ByVal arr As Variant)
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 1003, "Araya1.Read_Array", "配列ではありません。"
    End If
    
    ' 1次元かチェック（2次元だと UBound(arr,2) が取れてしまう）
    On Error Resume Next
    Dim t As Long
    t = UBound(arr, 2)
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 1004, "Araya1.Read_Array", "2次元配列は受け付けません（1次元のみ）。"
    End If
    Err.Clear
    On Error GoTo 0
    
    mOneDim = arr
    mLoaded = True
End Sub

'----------------------------
' 参照系（おすすめ）
'----------------------------
Public Property Get IsLoaded() As Boolean
    IsLoaded = mLoaded
End Property

Public Property Get Data() As Variant
    EnsureLoaded
    Data = mOneDim
End Property

Public Property Get Count() As Long
    EnsureLoaded
    Count = UBound(mOneDim) - LBound(mOneDim) + 1
End Property

Public Property Get Item(ByVal index As Long) As Variant
    EnsureLoaded
    If index < LBound(mOneDim) Or index > UBound(mOneDim) Then
        Err.Raise vbObjectError + 1005, "Araya1.Item", "index が範囲外です。"
    End If
    Item = mOneDim(index)
End Property

Public Function TypeNames() As Variant
    EnsureLoaded
    
    Dim lb As Long, ub As Long, i As Long
    lb = LBound(mOneDim)
    ub = UBound(mOneDim)
    
    Dim res() As Variant
    ReDim res(lb To ub)
    
    For i = lb To ub
        If IsError(mOneDim(i)) Then
            res(i) = "Error"
        ElseIf IsEmpty(mOneDim(i)) Then
            res(i) = "Empty"
        Else
            res(i) = TypeName(mOneDim(i))  ' "Double" "String" "Date" など
        End If
    Next
    
    TypeNames = res
End Function

'----------------------------
' 型統一（Transform）
'----------------------------
Public Sub CastToDoubleSafe(Optional ByVal emptyAsZero As Boolean = False, _
                            Optional ByVal invalidAsZero As Boolean = False, _
                            Optional ByVal treatDateAsInvalid As Boolean = True)
    EnsureLoaded
    
    Dim i As Long
    For i = LBound(mOneDim) To UBound(mOneDim)
        
        If IsError(mOneDim(i)) Or IsEmpty(mOneDim(i)) Then
            mOneDim(i) = IIf(emptyAsZero, 0#, Empty)
        
        ElseIf VarType(mOneDim(i)) = vbString And Trim$(CStr(mOneDim(i))) = "" Then
            mOneDim(i) = IIf(emptyAsZero, 0#, Empty)
        
        ElseIf IsDate(mOneDim(i)) And treatDateAsInvalid Then
            mOneDim(i) = IIf(invalidAsZero, 0#, Empty)
        
        ElseIf IsNumeric(mOneDim(i)) Then
            mOneDim(i) = CDbl(mOneDim(i))
        
        Else
            mOneDim(i) = IIf(invalidAsZero, 0#, Empty)
        End If
        
    Next i
End Sub

Public Sub CastToStringSafe(Optional ByVal emptyAsBlank As Boolean = True, _
                            Optional ByVal errorAsBlank As Boolean = True)
    EnsureLoaded
    
    Dim i As Long
    For i = LBound(mOneDim) To UBound(mOneDim)
        
        If IsError(mOneDim(i)) Then
            If errorAsBlank Then
                mOneDim(i) = ""
            Else
                Err.Raise vbObjectError + 1201, "Araya1.CastToStringSafe", "Error値が含まれています。"
            End If
        
        ElseIf IsEmpty(mOneDim(i)) Then
            mOneDim(i) = IIf(emptyAsBlank, "", Empty)
        
        Else
            mOneDim(i) = CStr(mOneDim(i))
        End If
        
    Next i
End Sub

Public Sub CastToDateSafe(Optional ByVal invalidAsEmpty As Boolean = True)
    EnsureLoaded
    
    Dim i As Long
    For i = LBound(mOneDim) To UBound(mOneDim)
        
        If IsError(mOneDim(i)) Or IsEmpty(mOneDim(i)) Then
            mOneDim(i) = Empty
        
        ElseIf VarType(mOneDim(i)) = vbString And Trim$(CStr(mOneDim(i))) = "" Then
            mOneDim(i) = Empty
        
        ElseIf IsDate(mOneDim(i)) Then
            mOneDim(i) = CDate(mOneDim(i))
        
        Else
            If invalidAsEmpty Then
                mOneDim(i) = Empty
            Else
                Err.Raise vbObjectError + 1202, "Araya1.CastToDateSafe", "日付に変換できない値が含まれています。"
            End If
        End If
        
    Next i
End Sub

'----------------------------
' 出力系
'----------------------------
Public Sub ToRangeVertical(ByVal topLeft As Range)
    EnsureLoaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 1006, "Araya1.ToRangeVertical", "topLeft が Nothing です。"
    
    Dim n As Long: n = Me.Count
    topLeft.Resize(n, 1).value = Application.WorksheetFunction.Transpose(mOneDim)
End Sub

Public Sub ToRangeHorizontal(ByVal topLeft As Range)
    EnsureLoaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 1007, "Araya1.ToRangeHorizontal", "topLeft が Nothing です。"
    
    Dim n As Long: n = Me.Count
    topLeft.Resize(1, n).value = mOneDim
End Sub

Public Sub Clear()
    mOneDim = Empty
    mLoaded = False
End Sub

'----------------------------
' 内部
'----------------------------
Private Sub EnsureLoaded()
    If Not mLoaded Then
        Err.Raise vbObjectError + 1099, "Araya1", "データが未読み込みです。Read_ColRange / Read_RowRange / Read_Array を先に呼んでください。"
    End If
End Sub
