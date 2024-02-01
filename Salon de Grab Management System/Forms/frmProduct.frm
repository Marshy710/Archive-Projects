VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmTransactionP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5745
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   10170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProduct.frx":0000
   ScaleHeight     =   5745
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTransactionNo 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8040
      TabIndex        =   4
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
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
      Format          =   44105729
      CurrentDate     =   40905
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   360
      TabIndex        =   6
      Top             =   1920
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16761087
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "product"
         Caption         =   "                  PRODUCTS"
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
         DataField       =   "description"
         Caption         =   "                 DESCRIPTION"
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
         DataField       =   "quantity"
         Caption         =   " QUANTITY"
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
      BeginProperty Column03 
         DataField       =   "totalCost"
         Caption         =   "       AMOUNT"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   2534.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4259.906
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin glxpbuttonz.UserButtonz cmdAddProducts 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Add New"
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
   Begin glxpbuttonz.UserButtonz cmdRemove 
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Remove"
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
   Begin glxpbuttonz.UserButtonz cmdSave 
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   8400
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
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
      Caption         =   "Products Transaction"
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
      TabIndex        =   8
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8520
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      ForeColor       =   &H80000018&
      Height          =   255
      Left            =   7440
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "Total: Php"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblTransactionNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction No:"
      ForeColor       =   &H80000014&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "frmProduct.frx":2EE27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10200
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   360
      Picture         =   "frmProduct.frx":44DFB
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   9495
   End
End
Attribute VB_Name = "frmTransactionP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varAddEdit As Boolean
Public varTransactionId2 As Integer
Private TTempTable2 As ADODB.Recordset
Private TTransactionRecord2 As ADODB.Recordset
Private TTransactionProducts2 As ADODB.Recordset
Private TTempTableList2 As ADODB.Recordset



Private Sub cmdAddProducts_Click()
     With frmProductsTransaction
        .varOk = False
        .Show vbModal
        If .varOk Then
           TTempTable2.AddNew
           TTempTable2!productid = .varProductNameId
           TTempTable2!productamount = .varPrice
           TTempTable2!quantity = .varQuantity
           TTempTable2.Update
           showDatagridData
        End If
    End With
cmdRemove.Enabled = True
cmdSave.Enabled = True

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdRemove_Click()
    With TTempTable2
        If Not .EOF Then
            .MoveFirst
            .Find "id = " & TTempTableList2!id
            If Not .EOF Then
                .Delete
                .MoveNext
            End If
        Else: MsgBox "No record found..", vbOKOnly + vbInformation
        End If
    End With
    
    showDatagridData
        
End Sub



Private Sub txtTransactionNo_KeyPress(KeyAscii As Integer)
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
    If txtTransactionNo.Text = "" Then
        MsgBox "Please enter transaction no...", vbOKOnly + vbInformation
        txtTransactionNo.SetFocus
        Exit Sub
    
    ElseIf TTempTable2.RecordCount < 1 Then
        MsgBox "No record to save...", vbOKOnly + vbInformation
        cmdAddProducts.SetFocus
        Exit Sub
    End If
    
    If varAddEdit Then
        If MsgBox("Add new record?...", vbYesNo + vbQuestion, "adding...") = vbNo Then Exit Sub
        TTransactionProducts2.AddNew
    ElseIf MsgBox("Update new record?...", vbYesNo + vbQuestion, "updating...") = vbNo Then Exit Sub
    End If
    
    With TTransactionProducts2
        !transactionNo = txtTransactionNo.Text
        !transactionDate = DateValue(DTPicker1.Value)
        !transactionTotalCost = CDbl(Label2.Caption)
        .Update
    End With
    
    
    'get the value current value transaction id..
    varTransactionId2 = TTransactionProducts2!id
    
    With TTransactionRecord2
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    .Delete
                    .MoveNext
                    DoEvents
                Loop
            End If
        End With


    With TTempTable2
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                TTransactionRecord2.AddNew
             '   If Not varAddEdit Then TTransactionRecord2!id = !id
                TTransactionRecord2!transid = varTransactionId2
                TTransactionRecord2!productid = !productid
                TTransactionRecord2!quantity = !quantity
                TTransactionRecord2!productamount = !productamount
                TTransactionRecord2.Update
                .MoveNext
                DoEvents
            Loop
        End If
    End With
    
    'Display message if record has done for updating of new record added..
    If varAddEdit Then
        MsgBox "New record was added...", vbOKOnly + vbInformation
    Else: MsgBox "New record was updated...", vbOKOnly + vbInformation
    End If
    
    With frmBillingTransaction2
        .varTotalAmount = CDbl(Label2.Caption)
        .Show vbModal
    End With
    Unload Me
End Sub

Private Sub Form_Load()

    conTable TTransactionProducts2, _
    "SELECT * FROM tproductstransaction2 " & _
    "WHERE id = " & varTransactionId2
    
    conTable TTransactionRecord2, _
    "SELECT * FROM transactionRecord2 " & _
    "WHERE transId = " & varTransactionId2
    
    conTable TTempTable2, _
    "SELECT * FROM temptransactionRecord2"
    
    With TTransactionRecord2
        If .RecordCount > 0 Or Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                TTempTable2.AddNew
                TTempTable2!id = !id
                TTempTable2!transid = !transid
                TTempTable2!productid = !productid
                TTempTable2!quantity = !quantity
                TTempTable2!productamount = !productamount
                TTempTable2.Update
                .MoveNext
                DoEvents
            Loop
        End If
    End With
    
     'setting-up date if not eddit..
    If Not varAddEdit Then
        DTPicker1.Value = DateValue(TTransactionProducts2!transactionDate)
        txtTransactionNo.Text = TTransactionProducts2!transactionNo
    Else: DTPicker1.Value = DateValue(Now)
    End If
    showDatagridData
    enableMe
End Sub
    
    'setting-up date if not eddit..
  '  If varAddEdit Then DTPicker1.Value = DateValue(Now)
    
   ' showDatagridData
'End Sub

Private Sub showDatagridData()
    conTable TTempTableList2, "SELECT trn.id, srv.product, srv.description, trn.quantity,trn.quantity * trn.productAmount as totalCost " & _
    "FROM temptransactionRecord2 AS trn LEFT JOIN tproducts As srv ON trn.productid = srv.id"
    
   Label2.Caption = "0"
    With TTempTableList2
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Label2.Caption = Format(CDbl(Label2.Caption) + CDbl(!totalCost), "#,##0.00")
                .MoveNext
                DoEvents
            Loop
        End If
    End With
    
    Set DataGrid1.DataSource = TTempTableList2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Removing record from temporary transaction record transaction...
    With TTempTable2
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
                DoEvents
            Loop
        End If
    End With
    
    Set TTempTable2 = Nothing
    Set TTempTableList2 = Nothing
    Set TTransactionRecord2 = Nothing
    Set TTransactionProducts2 = Nothing
End Sub

Private Sub enableMe()
Dim varEnable As Boolean
If varUserAdmin Or varAddEdit Then varEnable = True

cmdAddProducts.Visible = varEnable
cmdRemove.Visible = varEnable
cmdSave.Visible = varEnable
txtTransactionNo.Enabled = varEnable
DTPicker1.Enabled = varEnable
End Sub

