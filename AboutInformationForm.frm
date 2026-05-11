VERSION 5.00
Begin VB.Form AboutInformationForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5085
   LinkTopic       =   "AboutInformationForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label lblBody
      Caption         =   "VB6 clean-room rebuild for disc authoring research."
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label lblVersion
      Caption         =   "Version 0.1"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   420
      Width           =   1335
   End
   Begin VB.Label lblTitle
      Caption         =   "Kureha VB6 Rebuild"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   2175
   End
End
Attribute VB_Name = "AboutInformationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub LoadPreview(ByVal versionText As String)
    lblVersion.Caption = versionText
End Sub
