VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmAddExpense 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAddExpense.frx":0000
   ScaleHeight     =   4005
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtExpenseAmount 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtExpenses 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1320
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   81264641
      CurrentDate     =   40905
   End
   Begin glxpbuttonz.UserButtonz cmdSave 
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Save"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16711935
      ColorButtonUp   =   15309136
      ColorButtonDown =   16711935
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz cmdCancel 
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16711935
      ColorButtonUp   =   15309136
      ColorButtonDown =   16711935
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Expense"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Expense:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "frmAddExpense.frx":2EE27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmAddExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varId As Integer
Public varExpenses As String
Public varExpenseAmount As String
Public varAddEdit As Boolean
Public varOk As Boolean
Private TExpenses As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtExpenseAmount_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKey0 To vbKey9
  Case vbKeyBack, vbKeyClear, vbKeyDelete
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
    KeyAscii = 0
    Beep
    MsgBox "Please input numbers only...", vbOKOnly + vbInformation
End Select
End Sub

Private Sub cmdSave_Click()
    
    If txtExpenses.Text = "" Then
        MsgBox "Please enter new expense...", vbOKOnly + vbInformation
        txtExpenses.SetFocus
        Exit Sub
   
    ElseIf txtExpenseAmount.Text = "" Then
        MsgBox "Please enter amount...", vbOKOnly + vbInformation
        txtExpenseAmount.SetFocus
        Exit Sub
    End If
        
    If varAddEdit Then
        If MsgBox("Add new record?...", vbYesNo + vbInformation, "adding...") = vbNo Then Exit Sub
        TExpenses.AddNew
    ElseIf MsgBox("Update this record?...", vbYesNo + vbInformation, "updating..") = vbNo Then Exit Sub
    End If
    
    With TExpenses
        !Expenses = txtExpenses.Text
        !ExpenseAmount = txtExpenseAmount.Text
        !ExpenseDate = DateValue(DTPicker1.Value)
        .Update
    End With
    
    If varAddEdit Then
        MsgBox "New record added..", vbOKOnly + vbInformation
    Else: MsgBox "New record updated...", vbOKOnly + vbInformation
    End If
    
    varAddEdit = False
    varOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    If varAddEdit Then
        conTable TExpenses, "SELECT * FROM tExpenses"
    Else: conTable TExpenses, "SELECT * FROM tExpenses WHERE id = " & varId
    End If
    
    If Not varAddEdit Then setFields
    DTPicker1.Value = DateValue(Now)
End Sub

Private Sub setFields()
    If TExpenses.EOF Then Exit Sub
    TExpenses.MoveFirst
    txtExpenses.Text = varExpenses
    txtExpenseAmount.Text = varExpenseAmount
End Sub

