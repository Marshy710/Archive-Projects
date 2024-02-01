VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmServicesTransaction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmServicesTransaction.frx":0000
   ScaleHeight     =   4725
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   " "
      Top             =   3120
      Width           =   1215
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
      Left            =   1440
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmServicesTransaction.frx":2EE27
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtServiceName 
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtBeautician 
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin glxpbuttonz.UserButtonz cmdOk 
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Top             =   3960
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
      Left            =   3720
      TabIndex        =   11
      Top             =   3960
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
   Begin glxpbuttonz.UserButtonz cmdServices 
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   1560
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
   Begin glxpbuttonz.UserButtonz cmdBeautician 
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   960
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
      Caption         =   "Service"
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
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label txtServices 
      BackStyle       =   0  'Transparent
      Caption         =   "Services:"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Beautician:"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmServicesTransaction.frx":2EE29
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmServicesTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public varBeauticianId As Integer
Public varBeautician As String
Public varServiceNameId As Integer
Public varPrice As Currency
Public vardescription As String
Public varOk As Boolean
Private TTransactionServices As ADODB.Recordset


Private Sub cmdBeautician_Click()
    With frmsearchbox
        .varOk = False
        .varRecSet = 1
        .Show vbModal
           If .varOk Then
              varBeauticianId = .varId
              txtBeautician.Text = .varName
           End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If txtBeautician.Text = "" Then
        MsgBox "Please enter beautician...", vbOKOnly + vbInformation
        txtBeautician.SetFocus
        Exit Sub
    ElseIf varBeauticianId = 0 Then
        MsgBox "Please select beautician...", vbOKOnly
        txtBeautician.SetFocus
        Exit Sub
    ElseIf txtServiceName.Text = "" Then
        MsgBox "Please select services...", vbOKOnly + vbInformation
        txtServiceName.SetFocus
        Exit Sub
    End If
    
    varOk = True
    Unload Me
End Sub

Private Sub cmdServices_Click()
    With frmsearchbox
        .varRecSet = 2
        .Show vbModal
        If .varOk Then
            varServiceNameId = .varId
            txtServiceName.Text = .varName
            txtDescription.Text = .vardescription
            txtPrice.Text = .varPrice
            varPrice = .varPrice
        End If
    End With
End Sub
