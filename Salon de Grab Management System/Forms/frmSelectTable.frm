VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmSelectTable 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSelectTable.frx":0000
   ScaleHeight     =   3060
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin glxpbuttonz.UserButtonz cmdServices 
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "Services"
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
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "Products"
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
      Caption         =   "Select Transaction"
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
      Top             =   120
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Picture         =   "frmSelectTable.frx":2EE27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmSelectTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProducts_Click()
With frmTransactionP
        .varTransactionId2 = 0
        .varAddEdit = True
        .Show vbModal
    End With
    Unload Me
End Sub

Private Sub cmdServices_Click()
    With frmTransactionS
        .varTransactionId = 0
        .varAddEdit = True
        .Show vbModal
    End With
    Unload Me
End Sub
