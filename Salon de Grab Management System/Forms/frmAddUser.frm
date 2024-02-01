VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmAddUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAddUser.frx":0000
   ScaleHeight     =   4095
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAdmin 
      Caption         =   "Check1"
      Height          =   210
      Left            =   1920
      TabIndex        =   6
      Top             =   2640
      Width           =   210
   End
   Begin VB.TextBox txtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin glxpbuttonz.UserButtonz cmdAdd 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "&Add"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16744576
      ColorButtonUp   =   16744576
      ColorButtonDown =   16744576
      BorderBrightness=   0
      ColorBright     =   16761024
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz cmdCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
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
      Caption         =   "&Cancel"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   16744576
      ColorButtonUp   =   16744576
      ColorButtonDown =   16744576
      BorderBrightness=   0
      ColorBright     =   16761024
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmed Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "User's Form"
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
      TabIndex        =   10
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmAddUser.frx":2EE27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public varAddEdit As Boolean
Public varUserId As Integer
Public varOk As Boolean
Private TUsers As Recordset

Private Sub cmdAdd_Click()
    If txtUserName.Text = "" Then
        MsgBox "Please enter user name...", vbOKOnly + vbInformation
        txtUserName.SetFocus
        Exit Sub
    ElseIf txtPassword.Text = "" Then
        MsgBox "Please enter password...", vbOKOnly + vbInformation
        txtPassword.SetFocus
        Exit Sub
    ElseIf txtConfirm.Text = "" Then
        MsgBox "Please enter confirm password...", vbOKOnly + vbInformation
        txtConfirm.SetFocus
        Exit Sub
    ElseIf txtPassword.Text <> txtConfirm.Text Then
        MsgBox "mismatch password. Please verify...", vbOKOnly + vbInformation
        txtPassword.SetFocus
        Exit Sub
    End If
    
    If varAddEdit Then
        If MsgBox("Add new user?...", vbYesNo + vbInformation, "Adding...") = vbNo Then Exit Sub
        TUsers.AddNew
    ElseIf MsgBox("Update user account?...", vbYesNo + vbInformation, "Updating...") = vbNo Then Exit Sub
    End If
    
    With TUsers
        !UserName = txtUserName.Text
        !userpass = txtPassword.Text
        !uadmin = chkAdmin.Value
        .Update
    End With
    
    If varAddEdit Then
        MsgBox "New user added...", vbOKOnly + vbInformation
    Else: MsgBox "New user updated...", vbOKOnly + vbInformation
    End If
    
    varOk = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    conTable TUsers, "SELECT * FROM UsersAccount WHERE id = " & varUserId
    If varUserId = varUsersId Then
        chkAdmin.Enabled = False
    Else: chkAdmin.Enabled = True
    End If
    
    If varAddEdit = False Then
        Label5 = "Edit Account"
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set TUsers = Nothing
End Sub

