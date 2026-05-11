VERSION 5.00
Begin VB.Form TrackRippingForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Track Ripping"
   ClientHeight    =   2820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5685
   LinkTopic       =   "TrackRippingForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRip
      Caption         =   "Rip"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox cboRipFormat
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label lblRipFormat
      Caption         =   "Format:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1620
      Width           =   855
   End
   Begin VB.Label lblSummary
      Caption         =   "Track ripping preview"
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label lblHeader
      Caption         =   "Extract audio tracks from disc."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "TrackRippingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRip_Click()
    MsgBox "Track ripping pipeline is not wired yet.", vbInformation, "Track Ripping"
End Sub

Private Sub Form_Load()
    cboRipFormat.AddItem "WAV"
    cboRipFormat.AddItem "MP3"
    cboRipFormat.AddItem "FLAC"
    cboRipFormat.ListIndex = 0
End Sub

Public Sub LoadPreview(ByVal trackCount As Long)
    lblSummary.Caption = "Tracks available for ripping: " & CStr(trackCount)
End Sub

