VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmLogin2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SALON DE GRAB MANAGEMENT SYSTEM"
   ClientHeight    =   5730
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3240
      Width           =   2325
   End
   Begin glxpbuttonz.UserButtonz cmdOK 
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3720
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
      Caption         =   "Log In"
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
      Caption         =   "Project Version 1.0.1"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Management System "
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Salon De Grab Business Shop"
      BeginProperty Font 
         Name            =   "Perpetua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3480
      TabIndex        =   5
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Footlight MT Light"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.Image frmLogin2 
      Height          =   1500
      Left            =   -360
      Picture         =   "frmLogin2.frx":0000
      Top             =   120
      Width           =   9000
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   7305
      Left            =   -120
      Picture         =   "frmLogin2.frx":2BF64
      Stretch         =   -1  'True
      Top             =   600
      Width           =   8520
   End
End
Attribute VB_Name = "frmLogin2"
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
    LoginSucceeded = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
        If txtUserName.Text = "" Then
        MsgBox "Invalid Username/ Password, try again!", , "Login"
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
    Else
        MsgBox "Invalid Username/ Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub
