VERSION 5.00
Begin VB.Form IsReadyForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disc Ready"
   ClientHeight    =   1995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4650
   LinkTopic       =   "IsReadyForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3180
      TabIndex        =   2
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label lblDriveName
      Alignment       =   2  'Center
      Caption         =   "(default drive)"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Width           =   4170
   End
   Begin VB.Label lblMessage
      Caption         =   "Insert writable media into the selected drive."
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4170
   End
End
Attribute VB_Name = "IsReadyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub LoadPrompt(ByVal driveName As String, ByVal writable As Boolean)
    lblDriveName.Caption = driveName
    If writable Then
        lblMessage.Caption = "Insert writable media into the selected drive."
    Else
        lblMessage.Caption = "Insert readable media into the selected drive."
    End If
End Sub
