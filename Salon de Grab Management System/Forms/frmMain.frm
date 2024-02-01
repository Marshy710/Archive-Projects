VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SALON DE GRAB MANAGEMENT SYSTEM"
   ClientHeight    =   10650
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   16245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   16245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin glxpbuttonz.UserButtonz UserButtonz7 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   14520
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16761087
      ColorButtonUp   =   16761087
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz UserButtonz8 
      CausesValidation=   0   'False
      Height          =   615
      Left            =   14280
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16761087
      ColorButtonUp   =   16744703
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16777152
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz UserButtonz10 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   14520
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16761087
      ColorButtonUp   =   16761087
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz UserButtonz16 
      CausesValidation=   0   'False
      Height          =   615
      Left            =   1440
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2760
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16761087
      ColorButtonUp   =   16744703
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16777152
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz UserButtonz20 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16761087
      ColorButtonUp   =   16761087
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz UserButtonz9 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16761087
      ColorButtonUp   =   16761087
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz UserButtonz14 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4080
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16761087
      ColorButtonUp   =   16761087
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz UserButtonz12 
      Height          =   735
      Left            =   12840
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Reports"
      IconHighLiteColor=   16761087
      CaptionHighLiteColor=   0
      Picture         =   "frmMain.frx":0000
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
   Begin glxpbuttonz.UserButtonz UserButtonz3 
      Height          =   735
      Left            =   10560
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Edit"
      IconHighLiteColor=   16761087
      CaptionHighLiteColor=   0
      Picture         =   "frmMain.frx":0984
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
   Begin glxpbuttonz.UserButtonz UserButtonz2 
      CausesValidation=   0   'False
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&New Transaction "
      IconHighLiteColor=   16761087
      CaptionHighLiteColor=   0
      Picture         =   "frmMain.frx":1308
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
   Begin glxpbuttonz.UserButtonz UserButtonz4 
      Height          =   735
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Sales Record "
      IconHighLiteColor=   16761087
      CaptionHighLiteColor=   0
      Picture         =   "frmMain.frx":1C8C
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
   Begin glxpbuttonz.UserButtonz UserButtonz1 
      Height          =   735
      Left            =   5880
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Beauticians' Record"
      IconHighLiteColor=   16777215
      CaptionHighLiteColor=   16777215
      Picture         =   "frmMain.frx":2610
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
   Begin glxpbuttonz.UserButtonz UserButtonz5 
      Height          =   735
      Left            =   8280
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Expenses"
      IconHighLiteColor=   16761087
      CaptionHighLiteColor=   0
      Picture         =   "frmMain.frx":2F94
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H80000008&
      Height          =   8640
      Left            =   1680
      Picture         =   "frmMain.frx":3918
      ScaleHeight     =   8640
      ScaleWidth      =   12960
      TabIndex        =   1
      Top             =   2880
      Width           =   12960
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "All Rights Reserved I Project v1.0.1"
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   5040
         TabIndex        =   19
         Top             =   7440
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Salon De Grab Management System  "
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
         Left            =   4680
         TabIndex        =   18
         Top             =   7200
         Width           =   3375
      End
      Begin VB.Image Image4 
         Height          =   8325
         Left            =   0
         Picture         =   "frmMain.frx":3273F
         Top             =   0
         Width           =   12960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "All Rights Reserved I Project v1.0.0"
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   5520
         TabIndex        =   9
         Top             =   8160
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Salon De Grab Automation System  "
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
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   7920
         Width           =   3255
      End
   End
   Begin glxpbuttonz.UserButtonz UserButtonz11 
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   10680
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " "
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   4210816
      ColorButtonUp   =   4210816
      ColorButtonDown =   4210816
      BorderBrightness=   0
      ColorBright     =   4210816
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz UserButtonz6 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   14640
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tekton Pro Ext"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16761087
      ColorButtonUp   =   16761087
      ColorButtonDown =   16761087
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Image Image2 
      Height          =   2085
      Left            =   480
      Picture         =   "frmMain.frx":191AE3
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   15195
   End
   Begin VB.Image Image3 
      Height          =   1500
      Left            =   3480
      Picture         =   "frmMain.frx":1A7AB7
      Top             =   0
      Width           =   9000
   End
   Begin VB.Image Image1 
      Height          =   12000
      Left            =   0
      Picture         =   "frmMain.frx":1D3A1B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16260
   End
   Begin VB.Menu Files 
      Caption         =   "&Files"
      Begin VB.Menu TransactionS 
         Caption         =   "Services"
         Shortcut        =   ^E
      End
      Begin VB.Menu TransactionP 
         Caption         =   "Products"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu RReports 
      Caption         =   "&Reports"
      Begin VB.Menu SalesRecord 
         Caption         =   "Sales Record"
      End
      Begin VB.Menu BeauticianRecord 
         Caption         =   "Beauticians' Record"
      End
      Begin VB.Menu RExpenses 
         Caption         =   "Expenses"
      End
      Begin VB.Menu Summary 
         Caption         =   "Summary"
         Begin VB.Menu RSummary 
            Caption         =   "Report Summary"
         End
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "&Tools"
      Begin VB.Menu UserAccount 
         Caption         =   "User Account"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBeautician_Click()
    With frmBeautician
        .varAddEdit = True
        .Show vbModal
        
    End With
End Sub

Private Sub cmdProduct_Click()
    With frmInputProduct
        .Show vbModal
    End With
End Sub



Private Sub cmdProductList_Click()
    With frmProductsList
        .Show vbModal
    End With

End Sub

Private Sub cmdServicesList_Click()
    With frmServicesList
        .Show vbModal
    End With

End Sub

Private Sub About_Click()
    With frmAbout
        .Show vbModal
    End With
End Sub


Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub BeauticianRecord_Click()
  If varUserAdmin = True Then
     With frmBeauticianList
          .Show vbModal
     End With
  Else

Dim varId As Integer
Dim varName As String
Dim TSubTotal As Recordset
   With frmsearchbox
        .varRecSet = 1
        .Show vbModal
        If .varOk Then
            varId = .varId
            varName = .varName
        Else: Exit Sub
        End If
  End With
  
  With frmDateRange
    .Show vbModal
    If .varOk Then
    
            conTable TSubTotal, "SELECT SUM(trn.serviceAmount) As subTotal " & _
            "FROM (transactionRecord As trn) LEFT JOIN tservicestransaction As tsvr " & _
            "ON trn.transId = tsvr.id WHERE trn.beauticianId = " & varId & _
            " AND format(transactionDate,'mm/dd/yyyy') >= '" & Format(CDate(.varDateFrom), "mm/dd/yyyy") & _
            "' AND format(transactionDate,'mm/dd/yyyy') <= '" & Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            
            
            DataEnvironment1.rssqlBeauticianIncome.Open _
            "SELECT trn.id, tsvr.transactiondate, tsvr.transactionNo, svr.servicename, svr.description, trn.serviceAmount " & _
            "FROM ((transactionRecord AS trn) LEFT JOIN tservices As svr ON trn.serviceId = svr.id) " & _
            "LEFT JOIN tservicestransaction As tsvr ON trn.transId = tsvr.id WHERE trn.beauticianId = " & varId & _
            " AND format(transactionDate,'mm/dd/yyyy') >= '" & Format(CDate(.varDateFrom), "mm/dd/yyyy") & _
            "' AND format(transactionDate,'mm/dd/yyyy') <= '" & Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            rptBeauticianIncome.Sections("Section5").Controls("lblTotalAmount").Caption = Format(TSubTotal!subTotal, "#,##0.00")
            rptBeauticianIncome.Sections("Section4").Controls("lblBeautician").Caption = varName
            rptBeauticianIncome.Show vbModal
            DataEnvironment1.rssqlBeauticianIncome.Close
    End If
  End With
  End If
End Sub





Private Sub Form_Load()
    UserButtonz3.Enabled = varUserAdmin
   UserButtonz12.Enabled = varUserAdmin
End Sub



Private Sub Products_Click()
    With frmProductsList
        .Show vbModal
    End With
End Sub

Private Sub RExpenses_Click()
      With frmExpenses
        .Show vbModal
    End With
End Sub

Private Sub RSummary_Click()
Dim TServicesTotal As Recordset
Dim TProductTotal As Recordset
Dim TExpensesTotal As Recordset
Dim varServicesTotal As Double
Dim varProductTotal As Double
Dim varExpensesTotal As Double
Dim varNetTotal As Double

    With frmDateRange
        .Show vbModal
        If .varOk Then
            'Services Total Income
            conTable TServicesTotal, "SELECT SUM(transactionTotalCost) As subTotal " & _
            "FROM tservicestransaction WHERE format(transactionDate,'mm/dd/yyyy') >= '" & _
            Format(CDate(.varDateFrom), "mm/dd/yyyy") & _
            "' AND format(transactionDate,'mm/dd/yyyy') <= '" & _
            Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            If TServicesTotal!subTotal <> "" Then _
            varServicesTotal = CDbl(TServicesTotal!subTotal)
            
            'Product Total Income
            conTable TProductTotal, "SELECT SUM(transactionTotalCost) As subTotal " & _
            "FROM tproductstransaction2 WHERE format(transactionDate,'mm/dd/yyyy') >= '" & _
            Format(CDate(.varDateFrom), "mm/dd/yyyy") & _
            "' AND format(transactionDate,'mm/dd/yyyy') <= '" & _
            Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            If TProductTotal!subTotal <> "" Then _
            varProductTotal = CDbl(TProductTotal!subTotal)
    
            
            'Product Total Expenses
            conTable TExpensesTotal, "SELECT SUM(ExpenseAmount) As subTotal " & _
            "FROM tExpenses WHERE format(expenseDate,'mm/dd/yyyy') >= '" & _
            Format(CDate(.varDateFrom), "mm/dd/yyyy") & _
            "' AND format(expenseDate,'mm/dd/yyyy') <= '" & _
            Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            If TExpensesTotal!subTotal <> "" Then _
            varExpensesTotal = CDbl(TExpensesTotal!subTotal)
            
            'Net Income
            varNetTotal = (varServicesTotal + varProductTotal) - varExpensesTotal
            
            'Title report
            rptIncomeReport.Sections("Section4").Controls("lblIncomeReport").Caption = "INCOME REPORT FROM " & _
            Format(CDate(.varDateFrom), "mm/dd/yyyy") & " TO " & Format(CDate(.varDateTo), "mm/dd/yyyy")
            
            'Total Income
            rptIncomeReport.Sections("Section4").Controls("lblServiceIncome").Caption = Format(varServicesTotal, "#,##0.00")
            rptIncomeReport.Sections("Section4").Controls("lblProductSales").Caption = Format(varProductTotal, "#,##0.00")
            rptIncomeReport.Sections("Section4").Controls("lblGrossIncome").Caption = _
            Format(varServicesTotal + varProductTotal, "#,##0.00") 'Gross Income
            
            'Total Expenses
            rptIncomeReport.Sections("Section4").Controls("lblTotalExpenses").Caption = Format(varExpensesTotal, "#,##0.00")
           
                    
            'Net Income
             rptIncomeReport.Sections("Section4").Controls("lblTotalNetIncome").Caption = Format(varNetTotal, "#,##0.00")
            rptIncomeReport.Show vbModal
        End If
    End With
End Sub

Private Sub SalesRecord_Click()
    With frmSalesRecord
        .Show vbModal
    End With
End Sub

Private Sub TProducts_Click()
    With frmProductsList
        .Show vbModal
    End With
End Sub

Private Sub Services_Click()
    With frmServicesList
        .Show vbModal
    End With
End Sub

Private Sub TransactionP_Click()
    With frmTransactionP
        .varTransactionId2 = 0
        .varAddEdit = True
        .Show vbModal
    End With
End Sub

Private Sub TransactionS_Click()
     With frmTransactionS
        .varTransactionId = 0
        .varAddEdit = True
        .Show vbModal
    End With
End Sub

Private Sub TServices_Click()
    With frmServicesList
        .Show vbModal
    End With
End Sub



Private Sub UserAccount_Click()
     With frmUsersList
        .Show vbModal
    End With

End Sub

Private Sub UserButtonz1_Click()

  If varUserAdmin = True Then
     With frmBeauticianList
          .Show vbModal
     End With
  Else

Dim varId As Integer
Dim varName As String
Dim TSubTotal As Recordset
   With frmsearchbox
        .varRecSet = 1
        .Show vbModal
        If .varOk Then
            varId = .varId
            varName = .varName
        Else: Exit Sub
        End If
  End With
  
  With frmDateRange
    .Show vbModal
    If .varOk Then
    
            conTable TSubTotal, "SELECT SUM(trn.serviceAmount) As subTotal " & _
            "FROM (transactionRecord As trn) LEFT JOIN tservicestransaction As tsvr " & _
            "ON trn.transId = tsvr.id WHERE trn.beauticianId = " & varId & _
            " AND format(transactionDate,'mm/dd/yyyy') >= '" & Format(CDate(.varDateFrom), "mm/dd/yyyy") & _
            "' AND format(transactionDate,'mm/dd/yyyy') <= '" & Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            
            
            DataEnvironment1.rssqlBeauticianIncome.Open _
            "SELECT trn.id, tsvr.transactiondate, tsvr.transactionNo, svr.servicename, svr.description, trn.serviceAmount " & _
            "FROM ((transactionRecord AS trn) LEFT JOIN tservices As svr ON trn.serviceId = svr.id) " & _
            "LEFT JOIN tservicestransaction As tsvr ON trn.transId = tsvr.id WHERE trn.beauticianId = " & varId & _
            " AND format(transactionDate,'mm/dd/yyyy') >= '" & Format(CDate(.varDateFrom), "mm/dd/yyyy") & _
            "' AND format(transactionDate,'mm/dd/yyyy') <= '" & Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            rptBeauticianIncome.Sections("Section5").Controls("lblTotalAmount").Caption = Format(TSubTotal!subTotal, "#,##0.00")
            rptBeauticianIncome.Sections("Section4").Controls("lblBeautician").Caption = varName
            rptBeauticianIncome.Show vbModal
            DataEnvironment1.rssqlBeauticianIncome.Close
    End If
  End With
  End If
End Sub

Private Sub UserButtonz12_Click()
Dim TServicesTotal As Recordset
Dim TProductTotal As Recordset
Dim TExpensesTotal As Recordset
Dim varServicesTotal As Double
Dim varProductTotal As Double
Dim varExpensesTotal As Double
Dim varNetTotal As Double

    With frmDateRange
        .Show vbModal
        If .varOk Then
            'Services Total Income
            conTable TServicesTotal, "SELECT SUM(transactionTotalCost) As subTotal " & _
            "FROM tservicestransaction WHERE format(transactionDate,'mm/dd/yyyy') >= '" & _
            Format(CDate(.varDateFrom), "mm/dd/yyyy") & _
            "' AND format(transactionDate,'mm/dd/yyyy') <= '" & _
            Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            If TServicesTotal!subTotal <> "" Then _
            varServicesTotal = CDbl(TServicesTotal!subTotal)
            
            'Product Total Income
            conTable TProductTotal, "SELECT SUM(transactionTotalCost) As subTotal " & _
            "FROM tproductstransaction2 WHERE format(transactionDate,'mm/dd/yyyy') >= '" & _
            Format(CDate(.varDateFrom), "mm/dd/yyyy") & _
            "' AND format(transactionDate,'mm/dd/yyyy') <= '" & _
            Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            If TProductTotal!subTotal <> "" Then _
            varProductTotal = CDbl(TProductTotal!subTotal)
    
            
            'Product Total Expenses
            conTable TExpensesTotal, "SELECT SUM(ExpenseAmount) As subTotal " & _
            "FROM tExpenses WHERE format(expenseDate,'mm/dd/yyyy') >= '" & _
            Format(CDate(.varDateFrom), "mm/dd/yyyy") & _
            "' AND format(expenseDate,'mm/dd/yyyy') <= '" & _
            Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
            If TExpensesTotal!subTotal <> "" Then _
            varExpensesTotal = CDbl(TExpensesTotal!subTotal)
            
            'Net Income
            varNetTotal = (varServicesTotal + varProductTotal) - varExpensesTotal
            
            'Title report
            rptIncomeReport.Sections("Section4").Controls("lblIncomeReport").Caption = "INCOME REPORT FROM " & _
            Format(CDate(.varDateFrom), "mm/dd/yyyy") & " TO " & Format(CDate(.varDateTo), "mm/dd/yyyy")
            
            'Total Income
            rptIncomeReport.Sections("Section4").Controls("lblServiceIncome").Caption = Format(varServicesTotal, "#,##0.00")
            rptIncomeReport.Sections("Section4").Controls("lblProductSales").Caption = Format(varProductTotal, "#,##0.00")
            rptIncomeReport.Sections("Section4").Controls("lblGrossIncome").Caption = _
            Format(varServicesTotal + varProductTotal, "#,##0.00") 'Gross Income
            
            'Total Expenses
            rptIncomeReport.Sections("Section4").Controls("lblTotalExpenses").Caption = Format(varExpensesTotal, "#,##0.00")
           
                    
            'Net Income
             rptIncomeReport.Sections("Section4").Controls("lblTotalNetIncome").Caption = Format(varNetTotal, "#,##0.00")
            rptIncomeReport.Show vbModal
        End If
    End With
End Sub

Private Sub UserButtonz2_Click()
    With frmSelectTable
        .Show vbModal
    End With
End Sub

Private Sub UserButtonz3_Click()
    With frmModify
        .Show vbModal
    End With
End Sub

Private Sub UserButtonz4_Click()
    With frmSalesRecord
        .Show vbModal
    End With
End Sub

Private Sub UserButtonz5_Click()
    With frmExpenses
        .Show vbModal
    End With
End Sub


