VERSION 5.00
Begin VB.Form frmServices 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SERVICES FORM"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtPrice 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   1575
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lblPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
