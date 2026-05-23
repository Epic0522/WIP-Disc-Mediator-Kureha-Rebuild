VERSION 5.00
Begin VB.Form PropertyISOFile
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ISO File Properties"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   LinkTopic       =   "PropertyISOFile"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame fraAll
      Caption         =   "Entry"
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.TextBox txtFileName
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   420
         Width           =   3975
      End
      Begin VB.TextBox txtOriginalFileName
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtFileSize
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   1500
         Width           =   1815
      End
      Begin VB.TextBox txtTimestamp
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox chkHidden
         Caption         =   "Hidden file"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   2580
         Width           =   1815
      End
      Begin VB.Label lblFileName
         Caption         =   "File name:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblOriginal
         Caption         =   "Original:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lblSize
         Caption         =   "Size:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblTimestamp
         Caption         =   "Timestamp:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2100
         Width           =   1095
      End
   End
End
Attribute VB_Name = "PropertyISOFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Public Sub LoadEntry(ByVal fileName As String, ByVal originalName As String, ByVal sizeText As String, ByVal modifiedText As String)
    txtFileName.Text = fileName
    txtOriginalFileName.Text = originalName
    txtFileSize.Text = sizeText
    txtTimestamp.Text = modifiedText
End Sub
