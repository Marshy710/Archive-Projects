VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmUsersList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmUsersList.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmUsersList.frx":2EE27
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16761087
      HeadLines       =   2
      RowHeight       =   18
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "username"
         Caption         =   "                USER NAME"
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
         DataField       =   "useradmin"
         Caption         =   "      TYPE"
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
            Alignment       =   2
            ColumnAllowSizing=   -1  'True
            ColumnWidth     =   3945.26
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2399.811
         EndProperty
      EndProperty
   End
   Begin glxpbuttonz.UserButtonz cmdEdit 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   5040
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
   Begin glxpbuttonz.UserButtonz cmdAdd 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   5040
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
      Left            =   3240
      TabIndex        =   5
      Top             =   5040
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
   Begin glxpbuttonz.UserButtonz cmdExit 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   5040
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
   Begin VB.Image Image2 
      Height          =   420
      Left            =   6360
      Picture         =   "frmUsersList.frx":2F79B
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "User's Account List"
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
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "frmUsersList.frx":2FCCD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9120
   End
End
Attribute VB_Name = "frmUsersList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TUsersList As ADODB.Recordset

Private Sub cmdAdd_Click()
    With frmAddUser
        .varAddEdit = True
        .Show vbModal
        If .varOk Then
            loadUsers
        End If
    End With
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Delete this user?...", vbYesNo + vbInformation, "Deleting...") = vbNo Then Exit Sub
    TUsersList.Delete
End Sub

Private Sub cmdEdit_Click()
    With frmAddUser
        .varAddEdit = False
        .varUserId = TUsersList!id
        .txtUserName.Text = TUsersList!UserName
        .txtPassword.Text = TUsersList!userpass
        .txtConfirm.Text = TUsersList!userpass
        .chkAdmin.Value = CInt(TUsersList!uadmin)
        .Show vbModal
        If .varOk Then
            loadUsers
        End If
    End With
End Sub

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub Form_Load()
    loadUsers
End Sub

Private Sub loadUsers()
    conTable TUsersList, "SELECT id, username, userpass, " & _
    "uadmin, IIF(uadmin = 1,'admin','user') As useradmin FROM UsersAccount"
    Set DataGrid1.DataSource = TUsersList

End Sub
