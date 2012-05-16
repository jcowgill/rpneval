VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EvalContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary

'RPN Evaluator
' James Cowgill
'

'This class represents the context expressions run in
' the class stores any variables between calls to Evaluate
'

'List of tokens used during evaluation
Private Enum TokenID
    'Tokens appearing on the stack (must be first)
    TokBraketLeft
    TokExponent
    TokMultiply
    TokDivide
    TokAdd
    TokSubtract
    TokAssign
    
    'Not produced by the scanner (but used in the stack)
    TokNegate
    
    'Other tokens
    TokBraketRight
    TokNumber
    TokVariable
    TokEof
End Enum

'Last operator token
Private Const LastOperator As Long = TokNegate

'Error codes
Private Const ErrBase As Long = 60000
Private Const ErrInvalidChar As Long = ErrBase
Private Const ErrMultiplePoints As Long = ErrBase + 1
Private Const ErrVariableNotDefined As Long = ErrBase + 2
Private Const ErrMismatchedBrakets As Long = ErrBase + 3
Private Const ErrStackUnderflow As Long = ErrBase + 4
Private Const ErrTooManyOperands As Long = ErrBase + 5
Private Const ErrAssignToNumber As Long = ErrBase + 6

'Represents a token in the input stream
Private Type Token
    ID As TokenID
    Data As Variant     'Used by TokNumber and TokVariable only
End Type

'Operator precedences and assocativity
Private OpPrecedence(0 To LastOperator) As Long
Private OpRightAssociative(0 To LastOperator) As Boolean

'Variables
Private vars As Dictionary

Private Sub Class_Initialize()
    'Setup precedances
    OpPrecedence(TokNegate) = 4
    OpPrecedence(TokExponent) = 3
    OpPrecedence(TokMultiply) = 2
    OpPrecedence(TokDivide) = 2
    OpPrecedence(TokAdd) = 1
    OpPrecedence(TokSubtract) = 1
    OpPrecedence(TokAssign) = 0
    
    OpPrecedence(TokBraketLeft) = -100  'Must be the lowest
    
    'Setup right associative operators (defaults to left associative)
    OpRightAssociative(TokNegate) = True
    OpRightAssociative(TokExponent) = True
    OpRightAssociative(TokAssign) = True
    
    'Create variable container
    Set vars = New Dictionary
End Sub

'Gets the variables dictionary
Public Property Get Variables() As Dictionary
    Set Variables = vars
End Property

'Gets the next token in the input stream
' pos must be a variable which keeps a record of where we are in the string
Private Function TokenScan(ByVal str As String, ByRef pos As Long) As Token
    Dim c As String         'Current character
    Dim startPos As Long
    Dim hadDecimalPoint As Boolean
    
    'Ignore any leading whitespace
    Do While pos <= Len(str)
        'Get current character
        c = Mid$(str, pos, 1)
        
        If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then
            'Not whitespace so exit loop
            Exit Do
        End If
    
        'Advance position
        pos = pos + 1
    Loop
    
    'End of string?
    If pos > Len(str) Then
        TokenScan.ID = TokEof
        Exit Function
    End If
    
    'What is the first character?
    ' (c is set in the leading whitespace loop)
    Select Case c
        Case "+"
            'Add
            TokenScan.ID = TokAdd
            pos = pos + 1
        
        Case "-"
            'Subtract
            TokenScan.ID = TokSubtract
            pos = pos + 1
        
        Case "*"
            'Times
            TokenScan.ID = TokMultiply
            pos = pos + 1
        
        Case "/"
            'Divide
            TokenScan.ID = TokDivide
            pos = pos + 1
        
        Case "^"
            'Exponent
            TokenScan.ID = TokExponent
            pos = pos + 1
        
        Case "("
            'Braket Open
            TokenScan.ID = TokBraketLeft
            pos = pos + 1
        
        Case ")"
            'Braket Close
            TokenScan.ID = TokBraketRight
            pos = pos + 1
        
        Case "="
            'Assignment
            TokenScan.ID = TokAssign
            pos = pos + 1
            
        Case "0" To "9"
            'Number
            TokenScan.ID = TokNumber
            
            'Suck up numbers and decimal point
            startPos = pos
            pos = pos + 1
            Do While pos <= Len(str)
                'Get current character
                c = Mid$(str, pos, 1)
                
                If c = "." Then
                    'Only allow one decimal point
                    If hadDecimalPoint Then
                        Err.Raise ErrMultiplePoints, "EvalContext", "Numbers must contain only 1 decimal point"
                    Else
                        hadDecimalPoint = True
                    End If
                    
                ElseIf c < "0" Or c > "9" Then
                    'Not a number char so exit loop
                    Exit Do
                End If
                
                'Advance position
                pos = pos + 1
            Loop
            
            'Store number
            TokenScan.Data = CDbl(Mid$(str, startPos, pos - startPos))
            
        Case "A" To "Z", "a" To "z", "_"
            'Variable
            TokenScan.ID = TokVariable
            
            'Suck up letters, numbers and the underscore (0 or more)
            startPos = pos
            pos = pos + 1
            Do While pos <= Len(str)
                'Get current character
                c = Mid$(str, pos, 1)
                
                If (c < "a" Or c > "z") And (c < "A" Or c > "Z") And (c < "0" Or c > "9") And c <> "_" Then
                    'Not a variable char so exit loop
                    Exit Do
                End If
                
                'Advance position
                pos = pos + 1
            Loop
            
            'Store variable name
            TokenScan.Data = Mid$(str, startPos, pos - startPos)
            
        Case Else
            'Invalid character
            Err.Raise ErrInvalidChar, "EvalContext", "Invalid character """ & c & """ at position " & pos
            
    End Select
End Function

'Returns true if otherOp should be popped off the operator stack now
Private Function ShouldEvaluateOperator(ByVal myOp As TokenID, ByVal otherOp As TokenID)
    If OpRightAssociative(myOp) Then
        ShouldEvaluateOperator = OpPrecedence(myOp) < OpPrecedence(otherOp)
    Else
        ShouldEvaluateOperator = OpPrecedence(myOp) <= OpPrecedence(otherOp)
    End If
End Function

'Pops a value from the RPN stack as a number
Private Function PopNumber(ByVal rpnStack As Stack) As Double
    'Pop the value
    Dim value As Variant
    value = rpnStack.Pop
    
    'Number or variable?
    If VarType(value) = vbString Then
        'Lookup variable
        If vars.Exists(value) Then
            PopNumber = vars(value)
        Else
            Err.Raise ErrVariableNotDefined, "EvalContext", "Variable not defined: " & value
        End If
    Else
        'Use number directly
        PopNumber = value
    End If
End Function

'Evaluates an operator using the given RPN stack
Private Sub EvaluateOperator(ByVal op As TokenID, ByVal rpnStack As Stack)
    Dim result As Double
    Dim value1 As Double
    Dim value2 As Double
                
    'How many operands?
    If op = TokNegate Then
        'One operand
        If rpnStack.Count < 1 Then GoTo stackUnderflow
        
        'Do the negate inline
        result = -PopNumber(rpnStack)
        
    ElseIf op = TokAssign Then
        'Assignment
        If rpnStack.Count < 2 Then GoTo stackUnderflow
        
        'Get variable name and result to store
        Dim varName As Variant
        result = PopNumber(rpnStack)
        varName = rpnStack.Pop
        
        'Must be a variable and not a number
        If VarType(varName) <> vbString Then
            Err.Raise ErrAssignToNumber, "EvalContext", "You cannot assign to a number"
        End If
        
        'Do assignment
        vars(varName) = result
    
    Else
        'Two operands
        If rpnStack.Count < 2 Then GoTo stackUnderflow
        
        'Get operands
        value2 = PopNumber(rpnStack)
        value1 = PopNumber(rpnStack)
        
        'Do the operation
        Select Case op
            Case TokAdd
                result = value1 + value2
                
            Case TokSubtract
                result = value1 - value2
                
            Case TokMultiply
                result = value1 * value2
                
            Case TokDivide
                result = value1 / value2
                
            Case TokExponent
                result = value1 ^ value2
        End Select
    End If
    
    'Push result
    rpnStack.Push result
    Exit Sub
    
stackUnderflow:
    Err.Raise ErrStackUnderflow, "EvalContext", "Not enough operands left on the stack"
End Sub

'Evaluates the given expression
' This method implements both the shunting yard algorithm and an RPN evaluator
Public Function Evaluate(ByVal str As String) As Double
    Dim opStack As Stack        'Shunting yard and RPN stacks
    Dim rpnStack As Stack
    Dim currToken As Token      'Current token
    Dim pos As Long             'Position in input string
    Dim unaryMinus As Boolean   'True if next token could be an unary minus
    
    'Create stacks
    Set opStack = New Stack
    Set rpnStack = New Stack
    
    'First token can be an unary minus
    unaryMinus = True
    pos = 1
    
    'Start processing tokens
    Do
        'Get next token
        currToken = TokenScan(str, pos)
        
        'What is it?
        Select Case currToken.ID
            Case TokEof
                'Exit loop
                Exit Do
                
            Case TokNumber, TokVariable
                'Push the number / variable name onto the RPN stack
                rpnStack.Push currToken.Data
                unaryMinus = False
                
            Case TokBraketLeft
                'Push ( immediately
                opStack.Push TokBraketLeft
            
            Case TokBraketRight
                'Pop any operators until a (
                Do
                    'Any tokens?
                    If opStack.IsEmpty Then
                        'Did not find a (
                        Err.Raise ErrMismatchedBrakets, "EvalContext", "Mismatched brakets"
                    End If
                    
                    'Open braket?
                    If opStack.Top = TokBraketLeft Then
                        Exit Do
                    End If
                    
                    'Pop and evaluate operator
                    EvaluateOperator opStack.Pop, rpnStack
                Loop
                
                'Pop the braket
                opStack.Pop
                unaryMinus = False
                
            Case Is <= LastOperator
                'An operator
                
                'If this is a minus and can be unary, convert it to a negation
                If currToken.ID = TokSubtract And unaryMinus Then
                    currToken.ID = TokNegate
                End If
                
                'Do precedence loop
                Do While Not opStack.IsEmpty
                    'Check precedence of operator at the top of the stack
                    If Not ShouldEvaluateOperator(currToken.ID, opStack.Top) Then
                        'Exit loop
                        Exit Do
                    End If
                    
                    'Pop operator and evaluate it
                    EvaluateOperator opStack.Pop, rpnStack
                Loop
                
                'Push this operator onto the stack
                opStack.Push currToken.ID
                unaryMinus = True
    
        End Select
    Loop
    
    'Pop any operators from the op stack
    Do While Not opStack.IsEmpty
        'Open braket?
        If opStack.Top = TokBraketLeft Then
            Err.Raise ErrMismatchedBrakets, "EvalContext", "Mismatched brakets"
        End If
    
        'Pop and evaluate operator
        EvaluateOperator opStack.Pop, rpnStack
    Loop
    
    'RPN stack should have 1 operand
    If rpnStack.Count > 1 Then
        Err.Raise ErrTooManyOperands, "EvalContext", "Too many operands left on the stack (not enough operators?)"
    End If
    
    'Make empty strings equal to zero
    If rpnStack.IsEmpty Then
        Evaluate = 0
    Else
        Evaluate = PopNumber(rpnStack)
    End If
End Function
