VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2175
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   1285.061
   ScaleMode       =   0  'User
   ScaleWidth      =   3957.657
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   240
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   645
      Width           =   2325
   End
   Begin glxpbuttonz.UserButtonz cmdOk 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
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
      Left            =   2640
      TabIndex        =   5
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
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   360
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   720
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Private TUserLogin As Recordset


Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login

    Unload Me
End Sub

Private Sub cmdOK_Click()
    
        If txtUserName.Text = "" Then
        MsgBox "Please enter username...", vbOKOnly + vbInformation
        txtUserName.SetFocus
        Exit Sub
    
    ElseIf txtPassword.Text = "" Then
        MsgBox "Please enter Password...", vbOKOnly + vbInformation
        txtPassword.SetFocus
        Exit Sub
    End If
    
    'check for correct password
    conTable TUserLogin, "SELECT * FROM UsersAccount WHERE username = '" & txtUserName.Text & _
    "' AND userpass = '" & txtPassword.Text & "'"

    If TUserLogin.RecordCount = 1 Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        varUserAdmin = CInt(TUserLogin!uadmin)
        varUsersId = TUserLogin!id
        Unload Me
  
        Unload Me
    Else
        MsgBox "Invalid Username/ Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

