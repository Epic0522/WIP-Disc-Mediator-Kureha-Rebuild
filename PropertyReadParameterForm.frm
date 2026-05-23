VERSION 5.00
Begin VB.Form PropertyReadParameterForm
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Read Parameters"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5475
   LinkTopic       =   "PropertyReadParameterForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame fraCD
      Caption         =   "CD"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4995
      Begin VB.ComboBox cboReadCommandCD
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   1815
      End
      Begin VB.CheckBox chkReadFirstPregap
         Caption         =   "Read first pregap"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   900
         Width           =   2055
      End
      Begin VB.CheckBox chkOnlyFirstSession
         Caption         =   "Read only first session"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1260
         Width           =   2535
      End
      Begin VB.CheckBox chkCollectSubPQ
         Caption         =   "Collect SubPQ data"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1620
         Width           =   2295
      End
      Begin VB.CheckBox chkIgnoreError
         Caption         =   "Ignore read errors"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1980
         Width           =   2175
      End
      Begin VB.Label lblReadCommand
         Caption         =   "Read command:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "PropertyReadParameterForm"
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

Private Sub Form_Load()
    If cboReadCommandCD.ListCount = 0 Then
        cboReadCommandCD.AddItem "READ CD"
        cboReadCommandCD.AddItem "READ(10)"
        cboReadCommandCD.AddItem "READ CD MSF"
        cboReadCommandCD.ListIndex = 0
    End If
End Sub
