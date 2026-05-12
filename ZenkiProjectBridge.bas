Attribute VB_Name = "ZenkiProjectBridge"
Option Explicit

Private Const ZENKI_LANG_ENGLISH As Long = 0
Private Const ZENKI_LANG_JAPANESE As Long = 1

Private Const ZENKI_INFO_TITLE As Long = 0
Private Const ZENKI_INFO_PERFORMER As Long = 1
Private Const ZENKI_INFO_SONGWRITER As Long = 2
Private Const ZENKI_INFO_COMPOSER As Long = 3
Private Const ZENKI_INFO_ARRANGER As Long = 4
Private Const ZENKI_INFO_MESSAGE As Long = 5

Public Function SyncZenkiTracks(ByVal engine As ZenkiEngine, ByVal tracks As Collection, ByRef syncedCount As Long, ByRef skippedCount As Long) As Boolean
    Dim i As Long
    Dim dllTrackNo As Long
    Dim entry As TrackEntry

    On Error GoTo SyncError

    syncedCount = 0
    skippedCount = 0

    If engine Is Nothing Then Exit Function

    engine.ClearTracks
    engine.EnableTrackText CollectionHasCdText(tracks)

    dllTrackNo = 0
    For i = 1 To tracks.Count
        Set entry = tracks.Item(i)

        If Trim$(entry.FilePath) = "" Then
            skippedCount = skippedCount + 1
        Else
            If Not engine.AddTrackFile(entry.FilePath, MsfToFramesBridge(entry.Pregap), MsfToFramesBridge(entry.Postgap), ZenkiTrackFormatFlag(entry.Source)) Then
                Exit Function
            End If

            dllTrackNo = dllTrackNo + 1
            ApplyTrackTexts engine, dllTrackNo, entry
            syncedCount = syncedCount + 1
        End If
    Next i

    SyncZenkiTracks = True
    Exit Function

SyncError:
    SyncZenkiTracks = False
End Function

Public Function MsfToFramesBridge(ByVal value As String) As Long
    Dim parts() As String
    Dim minutesValue As Long
    Dim secondsValue As Long
    Dim framesValue As Long

    parts = Split(Trim$(value), ":")
    If UBound(parts) >= 0 Then minutesValue = Val(parts(0))
    If UBound(parts) >= 1 Then secondsValue = Val(parts(1))
    If UBound(parts) >= 2 Then framesValue = Val(parts(2))

    If secondsValue < 0 Then secondsValue = 0
    If secondsValue > 59 Then secondsValue = 59
    If framesValue < 0 Then framesValue = 0
    If framesValue > 74 Then framesValue = 74

    MsfToFramesBridge = ((minutesValue * 60) + secondsValue) * 75 + framesValue
End Function

Private Function CollectionHasCdText(ByVal tracks As Collection) As Boolean
    Dim i As Long
    Dim entry As TrackEntry

    For i = 1 To tracks.Count
        Set entry = tracks.Item(i)
        If TextItemHasCdText(entry.EnglishText) Or TextItemHasCdText(entry.JapaneseText) Then
            CollectionHasCdText = True
            Exit Function
        End If
    Next i
End Function

Private Function TextItemHasCdText(ByVal textItem As TOCCDText) As Boolean
    If textItem Is Nothing Then Exit Function
    If Not textItem.LanguageEnabled Then Exit Function

    If textItem.TitleEnabled And Trim$(textItem.Title) <> "" Then
        TextItemHasCdText = True
        Exit Function
    End If
    If textItem.PerformerEnabled And Trim$(textItem.Performer) <> "" Then
        TextItemHasCdText = True
        Exit Function
    End If
    If textItem.SongwriterEnabled And Trim$(textItem.Songwriter) <> "" Then
        TextItemHasCdText = True
        Exit Function
    End If
    If textItem.ComposerEnabled And Trim$(textItem.Composer) <> "" Then
        TextItemHasCdText = True
        Exit Function
    End If
    If textItem.ArrangerEnabled And Trim$(textItem.Arranger) <> "" Then
        TextItemHasCdText = True
        Exit Function
    End If
    If textItem.MessageEnabled And Trim$(textItem.Message) <> "" Then
        TextItemHasCdText = True
    End If
End Function

Private Sub ApplyTrackTexts(ByVal engine As ZenkiEngine, ByVal trackNo As Long, ByVal entry As TrackEntry)
    ApplyTextLanguage engine, trackNo, ZENKI_LANG_ENGLISH, entry.EnglishText
    ApplyTextLanguage engine, trackNo, ZENKI_LANG_JAPANESE, entry.JapaneseText
End Sub

Private Sub ApplyTextLanguage(ByVal engine As ZenkiEngine, ByVal trackNo As Long, ByVal languageNo As Long, ByVal textItem As TOCCDText)
    If textItem Is Nothing Then Exit Sub
    If Not textItem.LanguageEnabled Then Exit Sub

    ApplyTextField engine, trackNo, languageNo, ZENKI_INFO_TITLE, textItem.TitleEnabled, textItem.Title
    ApplyTextField engine, trackNo, languageNo, ZENKI_INFO_PERFORMER, textItem.PerformerEnabled, textItem.Performer
    ApplyTextField engine, trackNo, languageNo, ZENKI_INFO_SONGWRITER, textItem.SongwriterEnabled, textItem.Songwriter
    ApplyTextField engine, trackNo, languageNo, ZENKI_INFO_COMPOSER, textItem.ComposerEnabled, textItem.Composer
    ApplyTextField engine, trackNo, languageNo, ZENKI_INFO_ARRANGER, textItem.ArrangerEnabled, textItem.Arranger
    ApplyTextField engine, trackNo, languageNo, ZENKI_INFO_MESSAGE, textItem.MessageEnabled, textItem.Message
End Sub

Private Sub ApplyTextField(ByVal engine As ZenkiEngine, ByVal trackNo As Long, ByVal languageNo As Long, ByVal informationNo As Long, ByVal enabled As Boolean, ByVal value As String)
    If Not enabled Then Exit Sub
    If Trim$(value) = "" Then Exit Sub
    engine.SetText trackNo, languageNo, informationNo, value
End Sub

Private Function ZenkiTrackFormatFlag(ByVal sourceLabel As String) As Long
    Select Case UCase$(Trim$(sourceLabel))
        Case "WAV"
            ZenkiTrackFormatFlag = 2352
        Case Else
            ZenkiTrackFormatFlag = 0
    End Select
End Function
