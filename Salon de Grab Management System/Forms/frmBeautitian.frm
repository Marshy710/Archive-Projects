VERSION 5.00
Begin VB.Form frmBeautitian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BEAUTITIAN FORM"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6420
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtContactNo 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox txtAddress 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblContactNo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No.:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmBeautitian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
