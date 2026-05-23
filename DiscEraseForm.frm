VERSION 5.00
Begin VB.Form DiscEraseForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disc Erase"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   LinkTopic       =   "DiscEraseForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdErase
      Caption         =   "Erase"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox cboEraseMode
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblEraseMode
      Caption         =   "Erase mode:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label lblSummary
      Caption         =   "Erase rewritable media."
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblHeader
      Caption         =   "Disc erase preview"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "DiscEraseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdErase_Click()
    IsReadyForm.LoadPrompt "Default optical drive", True
    IsReadyForm.Show vbModal, Me
End Sub

Private Sub Form_Load()
    cboEraseMode.AddItem "Quick"
    cboEraseMode.AddItem "Full"
    cboEraseMode.ListIndex = 0
End Sub

Public Sub LoadPreview(ByVal mediaType As String)
    lblSummary.Caption = "Target media: " & mediaType
End Sub

