VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Arr() As Variant

Private Sub Class_Initialize()
    Clear
End Sub

Private Sub Class_Terminate()
    Clear
End Sub

Public Property Get At(ByVal i As Long) As Variant
    If IsEmpty Then
        At = Null
        Exit Property
    End If
    
    If i < 0 Then
        i = i + Size
    End If
    
    At = Arr(i)
End Property

Public Property Let At(ByVal i As Long, ByVal NewValue As Variant)
    If i < 0 Then
        i = i + Size
    End If
    
    Arr(i) = NewValue
End Property

Public Function Clear()
    Erase Arr
End Function

Public Property Get Clone() As ArrayEx
    Dim NewInstance As ArrayEx: Set NewInstance = New ArrayEx
    NewInstance.Resize (Me.Size)
    
    Dim i As Long
    For i = LowerBound To UpperBound
        NewInstance.At(i) = At(i)
    Next i
    
    Set Clone = NewInstance
End Property

Public Function Concat(ByVal Arg As ArrayEx) As ArrayEx
    Dim Ret As ArrayEx
    Set Ret = Me.Clone
    Ret.Resize (Ret.Size + Arg.Size)
    
    Dim i As Long
    For i = Arg.LowerBound To Arg.UpperBound
        Ret.At(i + Me.Size) = Arg.At(i)
    Next i
    
    Concat = Ret
End Function

Public Function Contains(ByVal Value As Variant) As Boolean
    Dim i As Long
    
    For i = LowerBound To UpperBound
        If At(i) = Value Then
            Contains = True
            Exit Function
        End If
    Next i
    
    Contains = False
End Function

Public Property Get Entity() As Variant()
    Dim Ret() As Variant
    ReDim Ret(UpperBound)
    
    Dim i As Long
    For i = LBound(Ret) To UBound(Ret)
        Ret(i) = At(i)
    Next i
    
    Entity = Ret
End Property

Public Property Let Entity(ByRef NewArray() As Variant)
    Resize (UBound(NewArray) - LBound(NewArray) + 1)
    Dim i As Long
    Dim j As Long
    j = LBound(NewArray)
    
    For i = LowerBound To UpperBound
        At(i) = NewArray(j)
        j = j + 1
    Next i
End Property

Public Property Get First() As Variant
    If IsEmpty Then
        First = Null
    Else
        First = At(LowerBound)
    End If
End Property

Public Property Let First(ByVal NewValue As Variant)
    If Not IsEmpty Then
        At(LowerBound) = NewValue
    End If
End Property

Public Function IndexOf(ByVal Element As Variant) As Long
    Dim i As Long
    
    For i = LowerBound To UpperBound
        If At(i) = Element Then
            IndexOf = i
            Exit Function
        End If
    Next i

    IndexOf = -1
End Function

Public Function IsEmpty() As Boolean
    On Error GoTo ErrorHandle
    
    If UBound(Arr) >= 0 Then
        IsEmpty = False
    Else
        IsEmpty = True
    End If
    
    Exit Function
    
ErrorHandle:
    IsEmpty = True
End Function

Public Function Join(ByVal Joiner As String) As String
    Dim Ret As String
    Dim i As Long
    
    If Me.IsEmpty Then
        Join = ""
        Exit Function
    End If
    
    Ret = First
    
    For i = LowerBound + 1 To UpperBound Step 1
        Ret = Ret & Joiner & At(i)
    Next i
    
    Join = Ret
End Function

Public Property Get Last() As Variant
    If IsEmpty Then
        Last = Null
    Else
        Last = At(UpperBound)
    End If
End Property

Public Property Let Last(ByVal NewValue As Variant)
    If Not IsEmpty Then
        At(UpperBound) = NewValue
    End If
End Property

Public Property Get LowerBound() As Long
    LowerBound = 0
End Property

Public Property Get Length() As Long
    Length = Size
End Property

Public Function Pop() As Variant
    If IsEmpty Then
        Pop = Null
        Exit Function
    End If
    
    Pop = Last
    Resize (Size - 1)
End Function

Public Function Push(ByVal NewValue As Variant) As Long
    Me.Resize (Me.Size + 1)
    Last = NewValue
    Push = Size
End Function

Public Function Resize(ByVal NewSize As Long)
    If NewSize > 0 Then
        ReDim Preserve Arr(NewSize - 1)
    Else
        Erase Arr
    End If
End Function

Public Function Shift() As Variant
    If IsEmpty Then
        Shift = Null
        Exit Function
    End If
    
    Shift = First
    
    If Size = 1 Then
        Erase Arr
        Exit Function
    End If
    
    Dim i As Long
    For i = LowerBound To UpperBound - 1
        Arr(i) = Arr(i + 1)
    Next i
    
    Resize (Size - 1)
End Function

Public Property Get Size() As Long
    If IsEmpty Then
        Size = 0
    Else
        Size = UpperBound - LowerBound + 1
    End If
End Property

Public Function Slice(ByVal BeginIndex As Long, Optional ByVal EndIndex As Long) As ArrayEx
    Dim Ret As ArrayEx: Set Ret = New ArrayEx
    
    If IsMissing(EndIndex) Then
        EndIndex = UpperBound
    End If
    
    Select Case BeginIndex
        Case Is < -Size
            BeginIndex = LowerBound
        Case -Size To LowerBound - 1
            BeginIndex = BeginIndex + Size
        Case LowerBound To UpperBound
            ' そのまま
        Case Is > UpperBound
            Set Slice = Ret
            Exit Function
        Case Else
            Err.Raise Number:=1000, Description:="BeginIndex の値が不正です。"
    End Select
    
    Select Case EndIndex
        Case Is < -Size
            Set Slice = Ret
            Exit Function
        Case -Size To LowerBound - 1
            EndIndex = EndIndex + Size
        Case LowerBound To UpperBound
            ' そのまま
        Case Is > UpperBound
            EndIndex = UpperBound
        Case Else
            Err.Raise Number:=1001, Description:="EndIndex の値が不正です。"
    End Select
    
    If BeginIndex > EndIndex Then
        Set Slice = Ret
        Exit Function
    End If
    
    Ret.Resize (EndIndex - BeginIndex + 1)
    
    Dim i As Long
    For i = Ret.LowerBound To Ret.UpperBound
        Ret.At(i) = At(i + BeginIndex)
    Next i
    
    Set Slice = Ret
End Function

Public Function Unshift(ByVal NewValue As Variant) As Long
    If IsEmpty Then
        Resize (1)
        At(0) = NewValue
        Unshift = 1
        Exit Function
    End If
    
    Resize (Me.Size + 1)
    
    Dim i As Long
    For i = UpperBound To LowerBound + 1 Step -1
        At(i) = At(i - 1)
    Next i
    
    Unshift = Size
End Function

Public Property Get UpperBound() As Long
    If IsEmpty Then
        UpperBound = -1
    Else
        UpperBound = UBound(Arr)
    End If
End Property
