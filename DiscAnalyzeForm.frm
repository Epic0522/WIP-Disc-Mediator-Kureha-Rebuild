VERSION 5.00
Begin VB.Form DiscAnalyzeForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disc Analyze"
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5925
   LinkTopic       =   "DiscAnalyzeForm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefresh
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2820
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2820
      Width           =   1095
   End
   Begin VB.TextBox txtSummary
      Height          =   2355
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "DiscAnalyzeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    ListStatusForm.LoadLines "Disc analysis stages", "TOC inspection" & vbCrLf & "Raw TOC inspection" & vbCrLf & "CD-TEXT pack inspection"
    ListStatusForm.Show vbModal, Me
End Sub

Public Sub LoadPreview(ByVal discLabel As String, ByVal mediaType As String, ByVal trackCount As Long, ByVal hasCdText As Boolean)
    txtSummary.Text = "Disc label: " & discLabel & vbCrLf & _
        "Media type: " & mediaType & vbCrLf & _
        "Track count: " & CStr(trackCount) & vbCrLf & _
        "CD-TEXT present: " & IIf(hasCdText, "Yes", "No") & vbCrLf & vbCrLf & _
        "This rebuilt window will eventually host TOC, raw TOC, and CD-TEXT inspection."
End Sub

