Attribute VB_Name = "MainModule"
Option Explicit

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Long
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const OFN_FILEMUSTEXIST As Long = &H1000
Public Const OFN_HIDEREADONLY As Long = &H4
Public Const OFN_OVERWRITEPROMPT As Long = &H2
Public Const OFN_PATHMUSTEXIST As Long = &H800
Public Const OFN_EXPLORER As Long = &H80000
Public Const HWND_NOTOPMOST As Long = -2
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOACTIVATE As Long = &H10

Public Sub Main()
    On Error Resume Next
    ChDrive Left$(App.Path, 1)
    ChDir App.Path
    On Error GoTo 0

    Load MainForm
    MainForm.Show
End Sub

Public Function ShowOpenFileDialog(ByVal ownerHwnd As Long, ByVal filterText As String, ByVal dialogTitle As String, Optional ByVal initialDir As String = "", Optional ByVal defaultExtension As String = "") As String
    Dim ofn As OPENFILENAME
    Dim fileBuffer As String
    Dim fileTitleBuffer As String

    fileBuffer = String$(1024, vbNullChar)
    fileTitleBuffer = String$(260, vbNullChar)

    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = ownerHwnd
        .lpstrFilter = Replace(filterText, "|", vbNullChar) & vbNullChar & vbNullChar
        .nFilterIndex = 1
        .lpstrFile = fileBuffer
        .nMaxFile = Len(fileBuffer)
        .lpstrFileTitle = fileTitleBuffer
        .nMaxFileTitle = Len(fileTitleBuffer)
        .lpstrInitialDir = initialDir
        .lpstrTitle = dialogTitle
        .flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY
        .lpstrDefExt = defaultExtension
    End With

    If GetOpenFileName(ofn) <> 0 Then
        ShowOpenFileDialog = Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
    Else
        ShowOpenFileDialog = ""
    End If
End Function

Public Function ShowSaveFileDialog(ByVal ownerHwnd As Long, ByVal filterText As String, ByVal dialogTitle As String, ByVal defaultFileName As String, Optional ByVal initialDir As String = "", Optional ByVal defaultExtension As String = "") As String
    Dim ofn As OPENFILENAME
    Dim fileBuffer As String
    Dim fileTitleBuffer As String

    If Len(defaultFileName) > 1000 Then defaultFileName = Left$(defaultFileName, 1000)
    fileBuffer = defaultFileName & String$(1024 - Len(defaultFileName), vbNullChar)
    fileTitleBuffer = String$(260, vbNullChar)

    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = ownerHwnd
        .lpstrFilter = Replace(filterText, "|", vbNullChar) & vbNullChar & vbNullChar
        .nFilterIndex = 1
        .lpstrFile = fileBuffer
        .nMaxFile = Len(fileBuffer)
        .lpstrFileTitle = fileTitleBuffer
        .nMaxFileTitle = Len(fileTitleBuffer)
        .lpstrInitialDir = initialDir
        .lpstrTitle = dialogTitle
        .flags = OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT
        .lpstrDefExt = defaultExtension
    End With

    If GetSaveFileName(ofn) <> 0 Then
        ShowSaveFileDialog = Left$(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
    Else
        ShowSaveFileDialog = ""
    End If
End Function

Public Function ShowFolderDialog(ByVal ownerHwnd As Long, ByVal dialogTitle As String) As String
    Dim shellApp As Object
    Dim folderObj As Object

    On Error GoTo FolderError

    Set shellApp = CreateObject("Shell.Application")
    Set folderObj = shellApp.BrowseForFolder(ownerHwnd, dialogTitle, 0)
    If folderObj Is Nothing Then Exit Function

    ShowFolderDialog = folderObj.Self.Path
    Exit Function

FolderError:
    ShowFolderDialog = ""
End Function

Public Sub SetFormAlwaysOnTop(ByVal hwnd As Long, ByVal enabled As Boolean)
    If enabled Then
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    Else
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
    End If
End Sub

Public Function QuoteShellPath(ByVal pathText As String) As String
    QuoteShellPath = Chr$(34) & pathText & Chr$(34)
End Function

