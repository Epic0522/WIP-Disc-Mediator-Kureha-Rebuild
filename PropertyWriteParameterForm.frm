VERSION 5.00
Begin VB.Form PropertyWriteParameterForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Write Parameters"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5475
   LinkTopic       =   "PropertyWriteParameterForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame fraCommon
      Caption         =   "Common"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4995
      Begin VB.ComboBox cboWriteCommand
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   1815
      End
      Begin VB.TextBox txtBlockLength
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Text            =   "32"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtBufferRate
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Text            =   "80"
         Top             =   1260
         Width           =   735
      End
      Begin VB.CheckBox chkOptimumPowerControl
         Caption         =   "Optimum power control"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1740
         Width           =   2535
      End
      Begin VB.CheckBox chkForceEject
         Caption         =   "Force eject before write"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2100
         Width           =   2535
      End
      Begin VB.Label lblCommand
         Caption         =   "Write command:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblBlock
         Caption         =   "Block length:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label lblBuffer
         Caption         =   "Buffer rate:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
   End
End
Attribute VB_Name = "PropertyWriteParameterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If cboWriteCommand.ListCount = 0 Then
        cboWriteCommand.AddItem "SAO"
        cboWriteCommand.AddItem "SAO RAW"
        cboWriteCommand.AddItem "DAO RAW+96"
        cboWriteCommand.ListIndex = 0
    End If
End Sub
