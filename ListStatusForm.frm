VERSION 5.00
Begin VB.Form ListStatusForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Status List"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   LinkTopic       =   "ListStatusForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3060
      Width           =   1095
   End
   Begin VB.ListBox lstStatusList
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   5160
   End
   Begin VB.Label lblCaption
      Caption         =   "Status"
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5160
   End
End
Attribute VB_Name = "ListStatusForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Public Sub LoadLines(ByVal captionText As String, ByVal linesText As String)
    Dim lines() As String
    Dim i As Long

    lblCaption.Caption = captionText
    lstStatusList.Clear
    lines = Split(linesText, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        If Trim$(lines(i)) <> "" Then lstStatusList.AddItem lines(i)
    Next i
End Sub
