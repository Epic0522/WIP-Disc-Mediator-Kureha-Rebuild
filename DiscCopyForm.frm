VERSION 5.00
Begin VB.Form DiscCopyForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disc Copy"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   LinkTopic       =   "DiscCopyForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCopy
      Caption         =   "Copy"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblTarget
      Caption         =   "Target drive: default"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1140
      Width           =   2895
   End
   Begin VB.Label lblSource
      Caption         =   "Source drive: default"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblSummary
      Caption         =   "Disc copy preview"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label lblHeader
      Caption         =   "Copy one disc directly to another drive."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3075
   End
End
Attribute VB_Name = "DiscCopyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCopy_Click()
    MsgBox "Disc copy pipeline is not wired yet.", vbInformation, "Disc Copy"
End Sub

Public Sub LoadPreview(ByVal mediaType As String)
    lblSummary.Caption = "Current project media: " & mediaType
End Sub

