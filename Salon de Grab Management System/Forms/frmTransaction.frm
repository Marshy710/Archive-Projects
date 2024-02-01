VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmTransactionS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10185
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTransaction.frx":0000
   ScaleHeight     =   5715
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   960
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
      Format          =   44105729
      CurrentDate     =   40905
   End
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
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16761087
      ForeColor       =   0
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
         DataField       =   "servicename"
         Caption         =   "               SERVICES"
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
         Caption         =   "                    DESCRIPTION"
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
         DataField       =   "name"
         Caption         =   "             BEAUTICIAN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "serviceAmount"
         Caption         =   "        AMOUNT"
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
         RecordSelectors =   0   'False
         BeginProperty Column00 
            DividerStyle    =   4
            ColumnWidth     =   2264.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3165.166
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   2385.071
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   1679.811
         EndProperty
      EndProperty
   End
   Begin glxpbuttonz.UserButtonz cmdAddServices 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   4440
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
      Top             =   4440
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
      Caption         =   "Services Transaction"
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
   Begin VB.Label Label3 
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
      ForeColor       =   &H8000000F&
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   8280
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblServicesRendered 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVICES RENDERED:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblTransactionNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction No:"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "frmTransaction.frx":2EE27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10320
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   360
      Picture         =   "frmTransaction.frx":44DFB
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   9495
   End
End
Attribute VB_Name = "frmTransactionS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varAddEdit As Boolean
Public varTransactionId As Integer
Private TTempTable As ADODB.Recordset
Private TTransactionRecord As ADODB.Recordset
Private TTransactionServices As ADODB.Recordset
Private TTempTableList As ADODB.Recordset
Private varBeauticianId As Integer


Private Sub cmdAddServices_Click()
    With frmServicesTransaction
        .varOk = False
        .Show vbModal
        If .varOk Then
            TTempTable.AddNew
            TTempTable!serviceid = .varServiceNameId
            TTempTable!beauticianId = .varBeauticianId
            TTempTable!serviceamount = .varPrice
            TTempTable.Update
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
    With TTempTable
        If Not .EOF Then
            .MoveFirst
            .Find "id = " & TTempTableList!id
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
    
    ElseIf TTempTable.RecordCount < 1 Then
        MsgBox "No record to save...", vbOKOnly + vbInformation
        cmdAddServices.SetFocus
        Exit Sub
    End If
    
    If varAddEdit Then
        If MsgBox("Add new record?...", vbYesNo + vbQuestion, "adding...") = vbNo Then Exit Sub
        TTransactionServices.AddNew
    ElseIf MsgBox("Update new record?...", vbYesNo + vbQuestion, "updating...") = vbNo Then Exit Sub
    End If
    
    With TTransactionServices
        !transactionNo = txtTransactionNo.Text
        !transactionDate = DateValue(DTPicker1.Value)
        !transactionTotalCost = CDbl(Label2.Caption)
        .Update
    End With
    
    
    'get the value current value transaction id..
    varTransactionId = TTransactionServices!id
    
        With TTransactionRecord
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    .Delete
                    .MoveNext
                    DoEvents
                Loop
            End If
        End With

    
    'Original temporary table and trying to update or add record for original & _
    record for transaction record
    
    With TTempTable
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                TTransactionRecord.AddNew
                TTransactionRecord!transid = varTransactionId
                TTransactionRecord!serviceid = !serviceid
                TTransactionRecord!beauticianId = !beauticianId
                TTransactionRecord!serviceamount = !serviceamount
                TTransactionRecord.Update
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
    
    With frmBillingTransaction
        .varTotalAmount = CDbl(Label2.Caption)
        .Show vbModal
    End With
    
    Unload Me
End Sub



Private Sub Form_Load()

    conTable TTransactionServices, _
    "SELECT * FROM tservicestransaction " & _
    "WHERE id = " & varTransactionId
    
    conTable TTransactionRecord, _
    "SELECT * FROM transactionRecord " & _
    "WHERE transId = " & varTransactionId
    
    conTable TTempTable, _
    "SELECT * FROM temptransactionRecord"
    
    With TTransactionRecord
        If .RecordCount > 0 Or Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                TTempTable.AddNew
                TTempTable!transid = !transid
                TTempTable!serviceid = !serviceid
                TTempTable!beauticianId = !beauticianId
                TTempTable!serviceamount = !serviceamount
                TTempTable.Update
                .MoveNext
                DoEvents
            Loop
        End If
    End With
    
    'setting-up date if not eddit..
    If Not varAddEdit Then
        DTPicker1.Value = DateValue(TTransactionServices!transactionDate)
        txtTransactionNo.Text = TTransactionServices!transactionNo
    Else: DTPicker1.Value = DateValue(Now)
    End If
    
    showDatagridData
    enableMe
End Sub

Private Sub showDatagridData()
    conTable TTempTableList, "SELECT trn.id, svr.servicename, svr.description, trn.serviceAmount, cian.name " & _
    "FROM ((temptransactionRecord AS trn) LEFT JOIN tbeautician As cian ON trn.beauticianid = cian.id) " & _
    "LEFT JOIN tservices As svr ON trn.serviceId = svr.id"
    
    Label2.Caption = "0"
    With TTempTableList
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                Label2.Caption = Format(CDbl(Label2.Caption) + CDbl(!serviceamount), "#,##0.00")
                .MoveNext
                DoEvents
            Loop
        End If
    End With
    
    Set DataGrid1.DataSource = TTempTableList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Removing record from temporary transaction record transaction...
    With TTempTable
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
                DoEvents
            Loop
        End If
    End With
    
    Set TTempTable = Nothing
    Set TTempTableList = Nothing
    Set TTransactionRecord = Nothing
    Set TTransactionServices = Nothing
End Sub

Private Sub enableMe()
Dim varEnable As Boolean
If varUserAdmin Or varAddEdit Then varEnable = True

cmdAddServices.Visible = varEnable
cmdRemove.Visible = varEnable
cmdSave.Visible = varEnable
txtTransactionNo.Enabled = varEnable
DTPicker1.Enabled = varEnable
End Sub
