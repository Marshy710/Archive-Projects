VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmBillingTransaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBillingTransaction.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtChange 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4320
      Width           =   4575
   End
   Begin VB.TextBox txtCash 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   33
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      MaxLength       =   11
      TabIndex        =   0
      Top             =   2880
      Width           =   4575
   End
   Begin VB.TextBox txtTotalAmount 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   33
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4575
   End
   Begin glxpbuttonz.UserButtonz cmdOk 
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   5640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BILLING..."
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
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   3960
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Received"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "frmBillingTransaction.frx":2EE27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmBillingTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varTotalAmount As Currency
Dim varTotalCash As Currency
Dim varTotalChange As Currency

Private Sub cmdOK_Click()
If MsgBox("Print Transaction?...", vbYesNo + vbQuestion, "Receipt...") = vbYes Then
     
        DataEnvironment1.rssqlServiceReceipt.Open
        rptServiceTransaction.Sections("Section5").Controls("lblTotalAmount").Caption = Format(txtTotalAmount.Text, "#,##0.00")
        rptServiceTransaction.Sections("Section5").Controls("labelCash").Caption = Format(txtCash.Text, "#,##0.00")
        rptServiceTransaction.Sections("Section5").Controls("labelChange").Caption = Format(txtChange.Text, "#,##0.00")
        rptServiceTransaction.PrintReport True
        DataEnvironment1.rssqlServiceReceipt.Close
        Unload rptServiceTransaction
        
Else: Unload Me
End If
   Unload Me
End Sub

Private Sub Form_Load()
    txtTotalAmount.Text = varTotalAmount
End Sub


Private Sub txtCash_Change()
    If txtCash.Text = "" Then Exit Sub
    txtChange.Text = CDbl(txtCash.Text) - varTotalAmount
    If (CDbl(txtCash.Text) - varTotalAmount) < 1 Then
        txtChange.Text = "(" & txtChange.Text & ")"
        txtChange.ForeColor = &HFF&
    Else: txtChange.ForeColor = &H0&
    End If
End Sub

Private Sub txtCash_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKey0 To vbKey9
  Case vbKeyBack, vbKeyClear, vbKeyDelete, 13
  Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab
  Case Else
    KeyAscii = 0
    Beep
    MsgBox "Please input numbers only...", vbOKOnly + vbInformation
End Select
    Select Case KeyAscii
        Case vbKey0 To vbKey9
        Case vbKeyBack
    End Select
    
    If KeyAscii = 13 Then cmdOK_Click
End Sub



