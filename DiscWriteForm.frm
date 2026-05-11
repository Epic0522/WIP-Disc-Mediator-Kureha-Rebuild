VERSION 5.00
Begin VB.Form DiscWriteForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disc Write"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5925
   LinkTopic       =   "DiscWriteForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdStartWrite
      Caption         =   "Start"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox chkWriteCDText
      Caption         =   "Write CD-TEXT when available"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ComboBox cboWriteMode
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label lblWriteMode
      Caption         =   "Write mode:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1740
      Width           =   1155
   End
   Begin VB.Label lblSummary
      Caption         =   "No project loaded."
      Height          =   1035
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label lblHeader
      Caption         =   "Disc write preview"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "DiscWriteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdStartWrite_Click()
    MsgBox "Write pipeline is not wired yet." & vbCrLf & "This dialog is the rebuilt entry point.", vbInformation, "Disc Write"
End Sub

Private Sub Form_Load()
    cboWriteMode.AddItem "SAO"
    cboWriteMode.AddItem "SAO RAW"
    cboWriteMode.AddItem "DAO RAW+96"
    cboWriteMode.ListIndex = 0
End Sub

Public Sub LoadPreview(ByVal discLabel As String, ByVal mediaType As String, ByVal trackCount As Long, ByVal hasCdText As Boolean)
    lblSummary.Caption = "Label: " & discLabel & vbCrLf & _
        "Media: " & mediaType & vbCrLf & _
        "Tracks: " & CStr(trackCount)
    chkWriteCDText.Value = Abs(hasCdText)
End Sub

