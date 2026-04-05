'----------------------------
'クラスモジュールの属性
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
Option Explicit                         '変数の明示宣言を強制

Private VECTOR As Variant               '
Private VECTOR_IS_LOADED As Boolean     '

'----------------------------
' 読込系
'----------------------------
Public Sub Read_ColRange(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 1001, "Vector.Read_ColRange", "rng  Nothing "
    
    If rng.Cells.Count = 1 Then
        Dim tmp1(1 To 1) As Variant
        tmp1(1) = rng.value
        mOneDim = tmp1
    Else
        ' �cRange �� 1�����z��
        mOneDim = Application.index(rng.value, 0, 1)
    End If
    
    mLoaded = True
End Sub

Public Sub Read_RowRange(ByVal rng As Range)
    If rng Is Nothing Then Err.Raise vbObjectError + 1002, "Vector.Read_RowRange", "rng �� Nothing �ł��B"
    
    If rng.Cells.Count = 1 Then
        Dim tmp1(1 To 1) As Variant
        tmp1(1) = rng.value
        mOneDim = tmp1
    Else
        ' ��Range �� 1�����z��
        mOneDim = Application.index(rng.value, 1, 0)
    End If
    
    mLoaded = True
End Sub

Public Sub Read_Array(ByVal arr As Variant)
    If Not IsArray(arr) Then
        Err.Raise vbObjectError + 1003, "Vector.Read_Array", "�z��ł͂���܂���B"
    End If
    
    ' 1�������`�F�b�N�i2�������� UBound(arr,2) �����Ă��܂��j
    On Error Resume Next
    Dim t As Long
    t = UBound(arr, 2)
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise vbObjectError + 1004, "Vector.Read_Array", "2�����z��͎󂯕t���܂���i1�����̂݁j�B"
    End If
    Err.Clear
    On Error GoTo 0
    
    mOneDim = arr
    mLoaded = True
End Sub

'----------------------------
' �Q�ƌn�i�������߁j
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
        Err.Raise vbObjectError + 1005, "Vector.Item", "index ���͈͊O�ł��B"
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
            res(i) = TypeName(mOneDim(i))  ' "Double" "String" "Date" �Ȃ�
        End If
    Next
    
    TypeNames = res
End Function

'----------------------------
' �^����iTransform�j
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
                Err.Raise vbObjectError + 1201, "Vector.CastToStringSafe", "Error�l���܂܂�Ă��܂��B"
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
                Err.Raise vbObjectError + 1202, "Vector.CastToDateSafe", "���t�ɕϊ��ł��Ȃ��l���܂܂�Ă��܂��B"
            End If
        End If
        
    Next i
End Sub

'----------------------------
' �o�͌n
'----------------------------
Public Sub ToRangeVertical(ByVal topLeft As Range)
    EnsureLoaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 1006, "Vector.ToRangeVertical", "topLeft �� Nothing �ł��B"
    
    Dim n As Long: n = Me.Count
    topLeft.Resize(n, 1).value = Application.WorksheetFunction.Transpose(mOneDim)
End Sub

Public Sub ToRangeHorizontal(ByVal topLeft As Range)
    EnsureLoaded
    If topLeft Is Nothing Then Err.Raise vbObjectError + 1007, "Vector.ToRangeHorizontal", "topLeft �� Nothing �ł��B"
    
    Dim n As Long: n = Me.Count
    topLeft.Resize(1, n).value = mOneDim
End Sub

Public Sub Clear()
    mOneDim = Empty
    mLoaded = False
End Sub

'----------------------------
' ����
'----------------------------
Private Sub EnsureLoaded()
    If Not mLoaded Then
        Err.Raise vbObjectError + 1099, "Vector", "�f�[�^�����ǂݍ��݂ł��BRead_ColRange / Read_RowRange / Read_Array ���ɌĂ�ł��������B"
    End If
End Sub
