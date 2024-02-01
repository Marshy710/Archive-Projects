VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmBeautician 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBeautician.frx":0000
   ScaleHeight     =   3735
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtContactNo 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
   Begin glxpbuttonz.UserButtonz cmdSave 
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3000
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
      Left            =   3600
      TabIndex        =   8
      Top             =   3000
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
   Begin VB.Label lblBeautician 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Beautician"
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
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblContactNo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No.:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -120
      Picture         =   "frmBeautician.frx":2EE27
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmBeautician"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varId As Integer
Public varName As String
Public varAddress As String
Public varContactNo As String
Public varAddEdit As Boolean
Public varOk As Boolean
Private TBeautician As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
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
    
    If txtName.Text = "" Then
        MsgBox "Please input beautician name...", vbOKOnly + vbInformation
        txtName.SetFocus
        Exit Sub
   
    ElseIf txtAddress.Text = "" Then
        MsgBox "Please input address...", vbOKOnly + vbInformation
        txtAddress.SetFocus
        Exit Sub
        
    ElseIf txtContactNo.Text = "" Then
        MsgBox "Please input Contact No.", vbOKOnly + vbInformation
        txtContactNo.SetFocus
        Exit Sub
    End If
    
    If varAddEdit Then
        If MsgBox("Add new record?...", vbYesNo + vbInformation, "adding...") = vbNo Then Exit Sub
        TBeautician.AddNew
    ElseIf MsgBox("Update this record?...", vbYesNo + vbInformation, "updating..") = vbNo Then Exit Sub
    End If
    
    With TBeautician
        !Name = txtName.Text
        !address = txtAddress.Text
        !contactNo = txtContactNo.Text
        .Update
    End With
    
    If varAddEdit Then
        MsgBox "New record added..", vbOKOnly + vbInformation
    Else: MsgBox "New record updated...", vbOKOnly + vbInformation
    End If
    
    varAddEdit = False
    varOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    If varAddEdit Then
        conTable TBeautician, "SELECT * FROM tbeautician"
    Else: conTable TBeautician, "SELECT * FROM tbeautician WHERE id = " & varId
    End If
    If varAddEdit = False Then
        lblBeautician = "Edit Profile"
    End If
    If Not varAddEdit Then setFields
End Sub

Private Sub setFields()
    If TBeautician.EOF Then Exit Sub
    TBeautician.MoveFirst
    txtName.Text = varName
    txtAddress.Text = varAddress
    txtContactNo.Text = varContactNo
End Sub
