VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmSalesRecord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSalesRecord.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin glxpbuttonz.UserButtonz cmdDate 
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   1680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Enabled         =   0   'False
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
      ColorButtonHover=   16711935
      ColorButtonUp   =   15309136
      ColorButtonDown =   16711935
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16761087
      HeadLines       =   1
      RowHeight       =   21
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "transactionNo"
         Caption         =   "        TRANSACTION NO."
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
         DataField       =   "transactionDate"
         Caption         =   "            DATE"
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
         DataField       =   "transactionTotalCost"
         Caption         =   "           AMOUNT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   2910.047
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   2369.764
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   2385.071
         EndProperty
      EndProperty
   End
   Begin glxpbuttonz.UserButtonz cmdServices 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Services"
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
   Begin glxpbuttonz.UserButtonz cmdProducts 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Products"
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
   Begin glxpbuttonz.UserButtonz cmdDetails 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Details"
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
   Begin glxpbuttonz.UserButtonz cmdRefresh 
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Refresh List"
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
   Begin glxpbuttonz.UserButtonz cmdPrint 
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "   &Print"
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
   Begin glxpbuttonz.UserButtonz cmdClose 
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   5880
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
      Caption         =   "   &Close"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Php"
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
      Left            =   1680
      TabIndex        =   12
      Top             =   5760
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   1800
      Picture         =   "frmSalesRecord.frx":2EE27
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Record"
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
      TabIndex        =   11
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
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
      Left            =   360
      TabIndex        =   8
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Date:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblViewRecords 
      BackStyle       =   0  'Transparent
      Caption         =   "View Records:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   0
      Picture         =   "frmSalesRecord.frx":2F3AC
      Top             =   0
      Width           =   9000
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   360
      Picture         =   "frmSalesRecord.frx":45380
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   7920
   End
End
Attribute VB_Name = "frmSalesRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private varServicesDetails As Boolean
Private varOk As Boolean
Private TTransactionList As ADODB.Recordset
Private varSqlString As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDate_Click()
    With frmDateRange
        .Show vbModal
       If .varOk = True Then
           If varServicesDetails Then
                varSqlString = "SELECT * FROM tservicestransaction WHERE format(transactionDate,'mm/dd/yyyy') >= '" & _
                Format(CDate(.varDateFrom), "mm/dd/yyyy") & "' AND format(transactionDate,'mm/dd/yyyy') <= '" & Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
                conTable TTransactionList, varSqlString
                Set DataGrid1.DataSource = TTransactionList
            Else
                varSqlString = "SELECT * FROM tproductstransaction2 WHERE format(transactionDate,'mm/dd/yyyy') >= '" & _
                Format(CDate(.varDateFrom), "mm/dd/yyyy") & "' AND format(transactionDate,'mm/dd/yyyy') <= '" & Format(CDate(.varDateTo), "mm/dd/yyyy") & "'"
                conTable TTransactionList, varSqlString
                Set DataGrid1.DataSource = TTransactionList
            
            End If
       End If
    End With
    subTotalAmount
End Sub

Private Sub cmdDetails_Click()
    
    If TTransactionList.RecordCount < 1 Then
        MsgBox "No record found on the list...", vbOKOnly + vbInformation
        Exit Sub
    ElseIf TTransactionList.BOF Or TTransactionList.EOF Then
        MsgBox "Please select record to view details...", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    
    
    If varServicesDetails Then
        With frmTransactionS
            .varAddEdit = False
            .varTransactionId = TTransactionList!id
            .Show vbModal
            
        End With
    Else
        With frmTransactionP
            .varAddEdit = False
            .varTransactionId2 = TTransactionList!id
            .Show vbModal
            
        End With
    End If
End Sub

Private Sub cmdPrint_Click()

    If varServicesDetails Then
        DataEnvironment1.rssqlServicesReport.Open varSqlString
        rptServicesReport.Sections("Section5").Controls("lblTotalAmount").Caption = Label3.Caption
        rptServicesReport.Show vbModal
        DataEnvironment1.rssqlServicesReport.Close
        
        Else
            DataEnvironment1.rssqlProductsReport.Open varSqlString
        rptProductsReport.Sections("Section5").Controls("lblTotalAmount").Caption = Label3.Caption
        rptProductsReport.Show vbModal
        DataEnvironment1.rssqlProductsReport.Close
    End If
End Sub

Private Sub cmdProducts_Click()
    varServicesDetails = False
    varSqlString = "SELECT * FROM tproductstransaction2 ORDER BY transactionDate DESC"
    conTable TTransactionList, varSqlString
    Set DataGrid1.DataSource = TTransactionList
    cmdDetails.Enabled = True
    cmdDate.Enabled = True
    cmdPrint.Enabled = varUserAdmin
    cmdRefresh.Enabled = True
    lblViewRecords.Caption = "Product Transactions List"
    subTotalAmount
End Sub

Private Sub cmdRefresh_Click()
    If varServicesDetails Then
        cmdServices_Click
    Else: cmdProducts_Click
    End If
    subTotalAmount
End Sub

Private Sub subTotalAmount()
    Label3.Caption = "0"
    With TTransactionList
        If Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                Label3.Caption = Format(CDbl(Label3.Caption) + CDbl(!transactionTotalCost), "#,##0.00")
                .MoveNext
                DoEvents
            Loop
        End If
    End With
End Sub

Private Sub cmdServices_Click()
    varServicesDetails = True
    varSqlString = "SELECT * FROM tservicestransaction ORDER BY transactionDate DESC"
    conTable TTransactionList, varSqlString
    Set DataGrid1.DataSource = TTransactionList
    cmdDetails.Enabled = True
    cmdDate.Enabled = True
    cmdPrint.Enabled = varUserAdmin
    cmdRefresh.Enabled = True
    lblViewRecords.Caption = "Service Transactions List"
    subTotalAmount
End Sub

Private Sub Form_Load()
    cmdPrint.Visible = varUserAdmin
    
End Sub

Private Sub Image4_Click()
End Sub
