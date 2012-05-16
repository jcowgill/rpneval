VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EvalForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expression Evaluator"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   4815
   End
   Begin MSComctlLib.ListView lstVars 
      Height          =   1455
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Evaluate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtExpression 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Variables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "EvalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private context As EvalContext

Private Sub cmdReset_Click()
    'Reset evaluation context
    Set context = New EvalContext
    
    'Set built in variables
    context.Variables.Add "e", 2.71828182845905
    context.Variables.Add "pi", 3.14159265358979
    
    'Reset vars list
    UpdateVars
    
    'Wipe input and result
    txtExpression = ""
    txtResult = ""
End Sub

Private Sub UpdateVars()
    'Update variable list
    lstVars.ListItems.Clear
    Set vars = context.Variables
    keyArray = vars.Keys
    For i = 0 To vars.Count - 1
        'Add item
        Set item = lstVars.ListItems.Add(, , keyArray(i))
        item.SubItems(1) = vars.item(keyArray(i))
    Next
End Sub

Private Sub cmdRun_Click()
    Dim vars As Dictionary
    Dim keyArray As Variant
    Dim i As Long
    Dim item As ListItem

    'Evaluate string and handle any errors
    On Error GoTo ErrHandle
    
    txtResult = ""
    txtResult = context.Evaluate(txtExpression)
    
    On Error GoTo 0
    
    UpdateVars
    Exit Sub
    
ErrHandle:
    'Oh noes
    ' One of ours?
    If Err.Number >= 60000 Then
        MsgBox "Error" & vbCrLf & "  " & Err.Description, vbExclamation
    Else
        MsgBox "Unexpected Error!" & vbCrLf & vbTab & Err.Description, vbCritical
    End If
End Sub

Private Sub Form_Load()
    'Reset context
    cmdReset_Click
End Sub
