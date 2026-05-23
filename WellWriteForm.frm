VERSION 5.00
Begin VB.Form WellWriteForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Write Image"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   LinkTopic       =   "WellWriteForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSetting
      Caption         =   "Settings"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   3660
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   3660
      Width           =   1095
   End
   Begin VB.Frame fraTarget
      Caption         =   "Target device"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6210
      Begin VB.ComboBox cboWriteDrive
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   3675
      End
      Begin VB.ComboBox cboWriteSpeed
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   900
         Width           =   1815
      End
      Begin VB.CheckBox chkVerify
         Caption         =   "Verify after write"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblDrive
         Caption         =   "Drive:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblSpeed
         Caption         =   "Speed:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.TextBox txtLog
      Height          =   1215
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   2100
      Width           =   6210
   End
   Begin VB.Label lblProgress
      BorderStyle     =   1  'Fixed Single
      Caption         =   ""
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   6210
   End
End
Attribute VB_Name = "WellWriteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSetting_Click()
    PropertyWriteParameterForm.Show vbModal, Me
End Sub

Private Sub Form_Load()
    If cboWriteDrive.ListCount = 0 Then
        cboWriteDrive.AddItem "Default optical drive"
        cboWriteDrive.ListIndex = 0
    End If
    If cboWriteSpeed.ListCount = 0 Then
        cboWriteSpeed.AddItem "Maximum"
        cboWriteSpeed.AddItem "8x"
        cboWriteSpeed.AddItem "4x"
        cboWriteSpeed.ListIndex = 0
    End If
    txtLog.Text = "Waiting for Momiji write pipeline..."
    lblProgress.Caption = "0%"
End Sub
