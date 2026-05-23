VERSION 5.00
Begin VB.Form WellSaveForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Image"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   LinkTopic       =   "WellSaveForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   2460
      Width           =   1095
   End
   Begin VB.TextBox txtLog
      Height          =   1575
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   5895
   End
   Begin VB.Label lblProgress
      BorderStyle     =   1  'Fixed Single
      Caption         =   ""
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label lblTitle
      Caption         =   "Image save task"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "WellSaveForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub LoadTask(ByVal discLabel As String, ByVal trackCount As Long)
    lblTitle.Caption = "Save image: " & discLabel
    txtLog.Text = "Tracks: " & CStr(trackCount) & vbCrLf & "Waiting for image writer integration."
    lblProgress.Caption = "0%"
End Sub
