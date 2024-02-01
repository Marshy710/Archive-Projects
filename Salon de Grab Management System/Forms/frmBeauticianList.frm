VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmBeauticianList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBeauticianList.frx":0000
   ScaleHeight     =   5955
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16761087
      HeadLines       =   2
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
         DataField       =   "name"
         Caption         =   "            NAME"
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
         DataField       =   "address"
         Caption         =   "              ADDRESS"
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
         DataField       =   "contactNo"
         Caption         =   " CONTACT NO."
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
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   -1  'True
            ColumnWidth     =   2715.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3509.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1814.74
         EndProperty
      EndProperty
   End
   Begin glxpbuttonz.UserButtonz cmdOk 
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   840
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
      ColorButtonHover=   16711935
      ColorButtonUp   =   15309136
      ColorButtonDown =   16711935
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz cmdAdd 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   5400
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
      Caption         =   "  &Add New"
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
   Begin glxpbuttonz.UserButtonz cmdDelete 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   5400
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
      Caption         =   "  &Delete"
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
   Begin glxpbuttonz.UserButtonz cmdEdit 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   5400
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
      Caption         =   "  &Edit"
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
      TabIndex        =   7
      Top             =   5400
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
      Caption         =   "  &Close"
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "View Records:"
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
      Left            =   6840
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Beautician List"
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
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Picture         =   "frmBeauticianList.frx":2EE27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmBeauticianList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TBeauticianList As ADODB.Recordset
Public varRecSet As Integer
Public varId As Integer
Public varName As String
Public vardescription As String
Public varPrice As Currency
Public varOk As Boolean
Public varQuantity As String
Private varTSearch As ADODB.Recordset
Private TTransactionProducts2 As ADODB.Recordset


Private Sub cmdAdd_Click()
    With frmBeautician
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

Private Sub cmdDelete_Click()
    If MsgBox("Delete this record?...", vbYesNo + vbCritical, "deleting...") = vbNo Then Exit Sub
    TBeauticianList.Delete
    MsgBox "Record was deleted...", vbOKOnly + vbInformation
    RefreshList
End Sub

Private Sub cmdEdit_Click()
    If TBeauticianList.EOF Or TBeauticianList.BOF Then
        MsgBox "No record found on the list...", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    With frmBeautician
        .varAddEdit = False
        .varId = TBeauticianList!id
        .varName = TBeauticianList!Name
        .varAddress = TBeauticianList!address
        .varContactNo = TBeauticianList!contactNo
        .Show vbModal
        If .varOk Then
            RefreshList
        End If
    End With
End Sub

Private Sub cmdRefresh_Click()
    RefreshList
End Sub

Private Sub cmdOK_Click()
Dim TSubTotal As Recordset

varId = TBeauticianList!id
varName = TBeauticianList!Name
  
  
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
End Sub

Private Sub Form_Load()
RefreshList
End Sub

Private Sub RefreshList()
        
  conTable TBeauticianList, "SELECT * FROM tbeautician"
    Set DataGrid1.DataSource = TBeauticianList

End Sub
