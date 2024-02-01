VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form frmDateRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3150
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDateRange.frx":0000
   ScaleHeight     =   3150
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      CalendarBackColor=   16761087
      CalendarTitleBackColor=   16744576
      Format          =   40435713
      CurrentDate     =   40965
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      CalendarBackColor=   16761087
      CalendarTitleBackColor=   16744576
      Format          =   40435713
      CurrentDate     =   40965
   End
   Begin glxpbuttonz.UserButtonz cmdOk 
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2400
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
      Left            =   2040
      TabIndex        =   6
      Top             =   2400
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
      Caption         =   "Set Date"
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
      TabIndex        =   4
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   0
      Picture         =   "frmDateRange.frx":2EE27
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "frmDateRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public varOk As Boolean
Public varDateFrom As String
Public varDateTo As String


Private Sub cmdCancel_Click()
    varOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    varDateFrom = DateValue(dtFrom.Value)
    varDateTo = DateValue(dtTo.Value)
    varOk = True
    Unload Me
End Sub

Private Sub Form_Load()
    dtFrom.Value = DateValue(Now)
    dtTo.Value = DateValue(Now)
End Sub

