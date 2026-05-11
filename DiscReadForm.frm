VERSION 5.00
Begin VB.Form DiscReadForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disc Read"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   LinkTopic       =   "DiscReadForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRead
      Caption         =   "Read"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2100
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2100
      Width           =   1095
   End
   Begin VB.ComboBox cboOutputType
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblOutputType
      Caption         =   "Output:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label lblSummary
      Caption         =   "Disc read preview"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label lblHeader
      Caption         =   "Read disc or image into project files."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   3075
   End
End
Attribute VB_Name = "DiscReadForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRead_Click()
    MsgBox "Disc read pipeline is not wired yet.", vbInformation, "Disc Read"
End Sub

Private Sub Form_Load()
    cboOutputType.AddItem "CUE/BIN"
    cboOutputType.AddItem "CCD/IMG"
    cboOutputType.AddItem "ISO"
    cboOutputType.ListIndex = 0
End Sub

Public Sub LoadPreview(ByVal mediaType As String)
    lblSummary.Caption = "Current target media: " & mediaType & vbCrLf & "This rebuilt dialog will host disc/image import behavior."
End Sub

