VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmProductsTransaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProductsTransaction.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQuantity 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      MaxLength       =   11
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtProductName 
      BackColor       =   &H00FFC0FF&
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtPrice 
      BackColor       =   &H00FFC0FF&
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin glxpbuttonz.UserButtonz cmdOk 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   3720
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
      Left            =   3240
      TabIndex        =   10
      Top             =   3720
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
   Begin glxpbuttonz.UserButtonz cmdProducts 
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   1080
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
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
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity:"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label txtProducts 
      BackStyle       =   0  'Transparent
      Caption         =   "Product:"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   3000
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   0
      Picture         =   "frmProductsTransaction.frx":2EE27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmProductsTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varProductNameId As Integer
Public varPrice As Currency
Public varQuantity As Integer
Public vardescription As String
Public varOk As Boolean
Private TTransactionProducts2 As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
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

Private Sub cmdOK_Click()
    If txtProductName.Text = "" Then
     MsgBox "Please select product...", vbOKOnly + vbInformation
        txtProductName.SetFocus
        Exit Sub
        
    ElseIf txtQuantity.Text = "" Then
        MsgBox "Please input quantity...", vbOKOnly + vbInformation
        txtQuantity.SetFocus
        Exit Sub
    End If
   varQuantity = CInt(txtQuantity.Text)
    varOk = True
   Unload Me
End Sub

Private Sub cmdProducts_Click()
    With frmsearchbox
        .varOk = False
        .varRecSet = 3
        .Show vbModal
    
        If .varOk Then
            varProductNameId = .varId
            txtProductName.Text = .varName
            txtDescription.Text = .vardescription
            txtPrice.Text = .varPrice
            varPrice = .varPrice
            
        End If
    End With
End Sub

