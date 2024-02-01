VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmServicesList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6750
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmServicesList.frx":0000
   ScaleHeight     =   6750
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox xSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5760
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin MSDataGridLib.DataGrid datagrid1 
      Height          =   4215
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16761087
      BorderStyle     =   0
      HeadLines       =   2
      RowHeight       =   21
      RowDividerStyle =   0
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
         DataField       =   "servicename"
         Caption         =   "        SERVICES"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "description"
         Caption         =   "          DESCRIPTION"
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
         DataField       =   "Price"
         Caption         =   "    PRICE"
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
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2220.094
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4724.788
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1305.071
         EndProperty
      EndProperty
   End
   Begin glxpbuttonz.UserButtonz cmdAdd 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   6120
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
   Begin glxpbuttonz.UserButtonz cmdDelete 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   6120
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
      Caption         =   "Delete"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   6120
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
      Caption         =   "Edit"
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
      Left            =   7440
      TabIndex        =   7
      Top             =   6120
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
      Caption         =   "Close"
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
   Begin VB.Image Image4 
      Height          =   480
      Left            =   1800
      Picture         =   "frmServicesList.frx":2EE27
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   8400
      Picture         =   "frmServicesList.frx":2F38A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label lblname 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   360
      Picture         =   "frmServicesList.frx":2F92C
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   8550
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Services List"
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
      Top             =   240
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "frmServicesList.frx":45900
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9360
   End
End
Attribute VB_Name = "frmServicesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TservicesList As ADODB.Recordset


Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub cmdAdd_Click()
  With frmAddService
        .varAddEdit = True
        .Show vbModal
        If .varOk Then
          RefreshList xSearch.Text
        End If
    End With
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Delete this record?...", vbYesNo + vbCritical, "deleting...") = vbNo Then Exit Sub
    TservicesList.Delete
    MsgBox "Record was deleted...", vbOKOnly + vbInformation
    RefreshList xSearch.Text
End Sub

Private Sub cmdEdit_Click()
    If TservicesList.EOF Or TservicesList.BOF Then
        MsgBox "No record found on the list...", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    With frmAddService
        .varAddEdit = False
        .varId = TservicesList!id
        .varservicename = TservicesList!servicename
        .vardescription = TservicesList!Description
        .varPrice = TservicesList!price
        .Show vbModal
        If .varOk Then
            RefreshList xSearch.Text
        End If
    End With
End Sub



Private Sub cmdRefresh_Click()
    RefreshList xSearch.Text
End Sub

Private Sub Form_Load()
    RefreshList xSearch.Text
End Sub

Private Sub RefreshList(xSearch As String)
    conTable TservicesList, "SELECT id, servicename As servicename, description, price " & _
        "FROM tservices WHERE servicename like '%" & xSearch & "%' " & _
        "ORDER BY servicename ASC"
        
    Set DataGrid1.DataSource = TservicesList
End Sub

Private Sub xSearch_Change()
    RefreshList xSearch.Text
End Sub

