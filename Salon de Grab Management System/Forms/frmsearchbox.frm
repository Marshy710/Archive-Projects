VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmsearchbox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmsearchbox.frx":0000
   ScaleHeight     =   5205
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2775
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16761087
      BorderStyle     =   0
      HeadLines       =   0
      RowHeight       =   21
      RowDividerStyle =   0
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "NAME"
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "sname"
         Caption         =   ""
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
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   4020.095
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSearch 
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
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin glxpbuttonz.UserButtonz cmdOk 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   4560
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
      Caption         =   "Ok"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   4560
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
   Begin VB.Image Image3 
      Height          =   420
      Left            =   4560
      Picture         =   "frmsearchbox.frx":2EE27
      Stretch         =   -1  'True
      Top             =   1080
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
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label label 
      BackStyle       =   0  'Transparent
      Caption         =   "Records"
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
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   720
      Picture         =   "frmsearchbox.frx":2F3C9
      Stretch         =   -1  'True
      Top             =   960
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      Picture         =   "frmsearchbox.frx":4539D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmsearchbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varRecSet As Integer
Public varId As Integer
Public varName As String
Public vardescription As String
Public varPrice As Currency
Public varOk As Boolean
Public varQuantity As String
Private varTSearch As ADODB.Recordset
Private TTransactionProducts2 As ADODB.Recordset



Private Sub cmdCancel_Click()
    Unload Me
    varOk = False
End Sub

Private Sub cmdOK_Click()
    varId = varTSearch!id
    If varRecSet = 1 Then
        varName = varTSearch!sname
        
    ElseIf varRecSet = 2 Then
        varPrice = varTSearch!price
        varName = varTSearch!sname
        vardescription = varTSearch!Description
        
          ElseIf varRecSet = 3 Then
        varPrice = varTSearch!price
        varName = varTSearch!sname
        vardescription = varTSearch!Description
    End If
    
    varOk = True
    Unload Me
End Sub


Private Sub Form_Load()
   loadData txtSearch.Text
End Sub

Private Sub loadData(xSearch As String)
    If varRecSet = 1 Then
        conTable varTSearch, "SELECT id, name As sname " & _
        "FROM tbeautician WHERE name like '%" & xSearch & "%' " & _
        "ORDER BY name"
        lblname.Caption = "Search Name:"
        label.Caption = "Select Beautician"
    ElseIf varRecSet = 2 Then
        conTable varTSearch, "SELECT id, servicename As sname, description, price " & _
        "FROM tservices WHERE servicename like '%" & xSearch & "%' " & _
        "ORDER BY servicename"
        lblname.Caption = "Search Service:"
        label.Caption = "Select Service"
    ElseIf varRecSet = 3 Then
        conTable varTSearch, "SELECT id, product As sname, description, price " & _
        "FROM tproducts WHERE product like '%" & xSearch & "%' " & _
        "ORDER BY product"
        lblname.Caption = "Search Product:"
        label.Caption = "Select Product"
    End If
    
    Set DataGrid1.DataSource = varTSearch

End Sub


Private Sub txtSearch_Change()
    loadData txtSearch.Text
End Sub
