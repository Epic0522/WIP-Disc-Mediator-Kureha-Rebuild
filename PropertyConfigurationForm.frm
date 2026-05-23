VERSION 5.00
Begin VB.Form PropertyConfigurationForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5475
   LinkTopic       =   "PropertyConfigurationForm"
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
      TabIndex        =   8
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame fraGeneral
      Caption         =   "General"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4995
      Begin VB.CheckBox chkVerifyCompare
         Caption         =   "Verify written data after write"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   3015
      End
      Begin VB.CheckBox chkEjectAfterMessage
         Caption         =   "Eject media after completion message"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   3495
      End
      Begin VB.ComboBox cboSuccessAction
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   2535
      End
      Begin VB.ComboBox cboFailureAction
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblSuccess
         Caption         =   "On success:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblFailure
         Caption         =   "On failure:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label lblNote
         Caption         =   "Configuration options are staged here before being wired to the rebuilt engine."
         Height          =   435
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   4395
      End
   End
End
Attribute VB_Name = "PropertyConfigurationForm"
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
    If cboSuccessAction.ListCount = 0 Then
        cboSuccessAction.AddItem "Do nothing"
        cboSuccessAction.AddItem "Close application"
        cboSuccessAction.AddItem "Shut down computer"
        cboSuccessAction.ListIndex = 0
    End If

    If cboFailureAction.ListCount = 0 Then
        cboFailureAction.AddItem "Do nothing"
        cboFailureAction.AddItem "Close application"
        cboFailureAction.AddItem "Shut down computer"
        cboFailureAction.ListIndex = 0
    End If
End Sub
