VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmAddProduct 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInputProduct.frx":0000
   ScaleHeight     =   4350
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdescription 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox txtprice 
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
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtproduct 
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
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
   Begin glxpbuttonz.UserButtonz cmdSave 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
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
      Left            =   3840
      TabIndex        =   8
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Product"
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
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Name:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   0
      Picture         =   "frmInputProduct.frx":2EE27
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmAddProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varId As Integer
Public varproduct As String
Public vardescription As String
Public varPrice As String
Public varAddEdit As Boolean
Public varOk As Boolean
Private TProducts As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtprice_KeyPress(KeyAscii As Integer)
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
    
     If txtproduct.Text = "" Then
        MsgBox "Please input product name...", vbOKOnly + vbInformation
        txtproduct.SetFocus
        Exit Sub
   
    ElseIf txtdescription.Text = "" Then
        MsgBox "Please input description...", vbOKOnly + vbInformation
        txtdescription.SetFocus
        Exit Sub
        
    ElseIf txtprice.Text = "" Then
        MsgBox "Please input price...", vbOKOnly + vbInformation
        txtprice.SetFocus
        Exit Sub
    End If
       
        
    If varAddEdit Then
        If MsgBox("Add new record?...", vbYesNo + vbInformation, "adding...") = vbNo Then Exit Sub
        TProducts.AddNew
    ElseIf MsgBox("Update this record?...", vbYesNo + vbInformation, "updating..") = vbNo Then Exit Sub
    End If
    
    With TProducts
        !product = txtproduct.Text
        !Description = txtdescription.Text
        !price = txtprice.Text
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
        conTable TProducts, "SELECT * FROM tproducts"
    Else: conTable TProducts, "SELECT * FROM tproducts WHERE id = " & varId
    End If
    
    If varAddEdit = False Then
        Label4 = "Edit Product"
    End If
    
    If Not varAddEdit Then setFields
End Sub

Private Sub setFields()
    If TProducts.EOF Then Exit Sub
    TProducts.MoveFirst
    txtproduct.Text = varproduct
    txtdescription.Text = vardescription
    txtprice.Text = varPrice
End Sub


