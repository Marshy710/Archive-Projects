VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmExpenses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmExpenses.frx":0000
   ScaleHeight     =   6945
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4695
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16761087
      HeadLines       =   2
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Expenses"
         Caption         =   "EXPENSES"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "ExpenseAmount"
         Caption         =   "AMOUNT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ExpenseDate"
         Caption         =   "DATE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   3225.26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin glxpbuttonz.UserButtonz cmdDate 
      Height          =   375
      Left            =   8040
      TabIndex        =   3
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
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
      Caption         =   "..."
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   -2147483635
      ColorButtonUp   =   16744576
      ColorButtonDown =   -2147483635
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz cmdPrint 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
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
      Caption         =   "&Print"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   -2147483635
      ColorButtonUp   =   16744576
      ColorButtonDown =   -2147483635
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz cmdAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
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
      Caption         =   "  &Add New"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   -2147483635
      ColorButtonUp   =   16744576
      ColorButtonDown =   -2147483635
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz cmdRefresh 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
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
      Caption         =   "  &Refresh"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   -2147483635
      ColorButtonUp   =   16744576
      ColorButtonDown =   -2147483635
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz cmdEdit 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3240
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
      Caption         =   "  &Edit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   -2147483635
      ColorButtonUp   =   16744576
      ColorButtonDown =   -2147483635
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz cmdDelete 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3840
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
      Caption         =   "  &Delete"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   -2147483635
      ColorButtonUp   =   16744576
      ColorButtonDown =   -2147483635
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz cmdClose 
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   6360
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
      Caption         =   "  &Close"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   -2147483635
      ColorButtonUp   =   16744576
      ColorButtonDown =   -2147483635
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Expenses List"
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
      TabIndex        =   10
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Expenses:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "frmExpenses.frx":2EE27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TExpensesList As ADODB.Recordset
Private varSqlString As String


Private Sub cmdAdd_Click()
    With frmAddExpense
        .varAddEdit = True
        .Show vbModal
        If .varOk Then
          RefreshList
        End If
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDate_Click()
     With frmDateRange
        .Show vbModal
       If .varOk = True Then
           varSqlString = "SELECT * FROM tExpenses WHERE format(ExpenseDate,'mm/dd/yyyy') >= '" & _
           Format(CDate(.varDateFrom), "mm/dd/yyyy") & "' AND format(ExpenseDate,'mm/dd/yyyy') <= '" & Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            conTable TExpensesList, varSqlString
            Set DataGrid1.DataSource = TExpensesList
       End If
    End With
    Label2.Caption = "0"
    With TExpensesList
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                Label2.Caption = Format(CDbl(Label2.Caption) + CDbl(!ExpenseAmount), "#,##0.00")
                .MoveNext
                DoEvents
            Loop
        End If
    End With
    
    Set DataGrid1.DataSource = TExpensesList
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Delete this record?...", vbYesNo + vbCritical, "deleting...") = vbNo Then Exit Sub
    TExpensesList.Delete
    MsgBox "Record was deleted...", vbOKOnly + vbInformation
    RefreshList
End Sub

Private Sub cmdEdit_Click()
    If TExpensesList.EOF Or TExpensesList.BOF Then
        MsgBox "No record found on the list...", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    With frmAddExpense
        .varAddEdit = False
        .varId = TExpensesList!id
        .varExpenses = TExpensesList!Expenses
        .varExpenseAmount = TExpensesList!ExpenseAmount
        .Show vbModal
        If .varOk Then
            RefreshList
        End If
    End With
End Sub

Private Sub cmdPrint_Click()
        DataEnvironment1.rssqlExpensesReport.Open varSqlString
        rptExpensesReport.Sections("Section5").Controls("lblTotalAmount").Caption = Label2.Caption
        rptExpensesReport.Show vbModal
        DataEnvironment1.rssqlExpensesReport.Close
End Sub

Private Sub cmdRefresh_Click()
    RefreshList
End Sub

Private Sub Form_Load()
    RefreshList
    subTotal
    enableMe
End Sub

Private Sub RefreshList()
    varSqlString = "SELECT * FROM tExpenses ORDER BY expenseDate DESC"
    conTable TExpensesList, varSqlString
    Set DataGrid1.DataSource = TExpensesList

End Sub

Private Sub subTotal()
    Label2.Caption = "0"
    With TExpensesList
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                Label2.Caption = Format(CDbl(Label2.Caption) + CDbl(!ExpenseAmount), "#,##0.00")
                .MoveNext
                DoEvents
            Loop
        End If
    End With
    
    Set DataGrid1.DataSource = TExpensesList
End Sub

Private Sub enableMe()
Dim varEnable As Boolean
If varUserAdmin Or varAddEdit Then varEnable = True

cmdPrint.Visible = varEnable
cmdEdit.Visible = varEnable
cmdDelete.Visible = varEnable
End Sub
