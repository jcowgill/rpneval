VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'A stack of objects
' You can add any type of object to this stack EXCEPT Types
' If you want to add a compound type, create a class module to store the data in
'

Private my_data() As Variant
Private my_len As Long

Private Sub EnsureCapacity(ByVal capacity As Long)
    'Ensure enough capacity
    If UBound(my_data) < capacity Then
        'Double array size
        ReDim Preserve my_data(UBound(my_data) * 2)
    End If
End Sub

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of elements on the stack"
    'Returns the number of items on the stack
    Count = my_len
End Property

Public Property Get IsEmpty() As Boolean
Attribute IsEmpty.VB_Description = "Returns true if the stack is empty"
    'Returns true if the stack is empty
    IsEmpty = (my_len = 0)
End Property

Public Sub Clear()
Attribute Clear.VB_Description = "Clears all the elements from the stack"
    'Clears the stack (makes it empty)
    Erase my_data
    ReDim my_data(16)
    my_len = 0
End Sub

Public Sub Push(ByVal value As Variant)
Attribute Push.VB_Description = "Pushes a value onto the stack"
    'Push a value onto the stack
    
    'Ensure capacity
    EnsureCapacity my_len + 1
    
    'Add item
    If IsObject(value) Then
        Set my_data(my_len) = value
    Else
        my_data(my_len) = value
    End If
    
    my_len = my_len + 1
End Sub

Public Function Pop() As Variant
Attribute Pop.VB_Description = "Pops a value from the stack"
    'Pop a value off the stack
    ' Raises error 5 (invalid procedure call) if there are no values left
    
    'Any elements?
    If IsEmpty Then Err.Raise 5, "Stack"
    
    'Objects must be treated specially here
    If IsObject(my_data(my_len - 1)) Then
        Set Pop = my_data(my_len - 1)
    Else
        Pop = my_data(my_len - 1)
    End If
    
    'Remove object
    my_len = my_len - 1
    my_data(my_len) = Empty
End Function

Public Property Get Top() As Variant
Attribute Top.VB_Description = "Gets or sets the value at the top of the stack"
    'Gets the top value from the stack without popping it
    
    'Any elements?
    If IsEmpty Then Err.Raise 5, "Stack"
    
    'Objects must be treated specially here
    If IsObject(my_data(my_len - 1)) Then
        Set Top = my_data(my_len - 1)
    Else
        Top = my_data(my_len - 1)
    End If
End Property

Public Property Let Top(ByVal value As Variant)
    'Changes the top value of the stack
    
    'Any elements?
    If IsEmpty Then Err.Raise 5, "Stack"
    
    'Copy value
    my_data(my_len - 1) = value
End Property

Public Property Set Top(ByVal value As Variant)
    'Changes the top value of the stack
    
    'Any elements?
    If IsEmpty Then Err.Raise 5, "Stack"
    
    'Copy value
    Set my_data(my_len - 1) = value
End Property

Private Sub Class_Initialize()
    'Reset array
    Clear
End Sub

