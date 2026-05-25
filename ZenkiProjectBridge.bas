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

Public Function SyncZenkiIsoEntries(ByVal engine As ZenkiEngine, ByVal entries As Collection, ByRef syncedCount As Long, ByRef reportText As String) As Boolean
    Dim i As Long
    Dim entry As FileEntry

    On Error GoTo SyncError

    syncedCount = 0
    reportText = ""

    If engine Is Nothing Then Exit Function

    If Not engine.InitIsoFileSystem(0) Then
        reportText = "InitISOFS failed."
        Exit Function
    End If

    For i = 1 To entries.Count
        Set entry = entries.Item(i)
        If entry.IsDirectory Then
            If Not engine.MakeIsoDirectory(CleanIsoDisplayName(entry.Name)) Then
                reportText = "MakeISODirectory failed: " & entry.Name
                Exit Function
            End If
        ElseIf Trim$(entry.OriginalPath) <> "" Then
            If Not engine.AddIsoFile(CleanIsoDisplayName(entry.Name), entry.OriginalPath) Then
                reportText = "AddISOFile failed: " & entry.Name
                Exit Function
            End If
        Else
            If Not engine.AddIsoDummyFile(CleanIsoDisplayName(entry.Name), SizeTextToBytes(entry.SizeText)) Then
                reportText = "AddISODummyFile failed: " & entry.Name
                Exit Function
            End If
        End If
        syncedCount = syncedCount + 1
    Next i

    reportText = "Zenki ISO synced " & CStr(syncedCount) & " item(s)."
    SyncZenkiIsoEntries = True
    Exit Function

SyncError:
    reportText = Err.Description
    SyncZenkiIsoEntries = False
End Function

Public Function VerifyZenkiTracks(ByVal engine As ZenkiEngine, ByVal tracks As Collection, ByRef reportText As String) As Boolean
    Dim i As Long
    Dim dllTrackNo As Long
    Dim entry As TrackEntry
    Dim mismatchText As String

    If engine Is Nothing Then Exit Function

    If engine.TrackCount <> CountBackedTracks(tracks) Then
        reportText = "Track count mismatch. VB=" & CStr(CountBackedTracks(tracks)) & " DLL=" & CStr(engine.TrackCount)
        Exit Function
    End If

    For i = 1 To tracks.Count
        Set entry = tracks.Item(i)
        If Trim$(entry.FilePath) <> "" Then
            dllTrackNo = dllTrackNo + 1
            mismatchText = VerifyTrack(engine, dllTrackNo, entry)
            If mismatchText <> "" Then
                reportText = mismatchText
                Exit Function
            End If
        End If
    Next i

    reportText = "Zenki verified " & CStr(dllTrackNo) & " track(s)."
    VerifyZenkiTracks = True
End Function

Public Function SizeTextToBytes(ByVal sizeText As String) As Long
    Dim valueText As String
    Dim numberValue As Double
    Dim upperText As String

    upperText = UCase$(Trim$(sizeText))
    valueText = Replace(upperText, ",", "")
    valueText = Replace(valueText, "KB", "")
    valueText = Replace(valueText, "KIB", "")
    valueText = Replace(valueText, "MB", "")
    valueText = Replace(valueText, "MIB", "")
    valueText = Replace(valueText, "B", "")
    valueText = Trim$(valueText)

    If valueText = "" Or Left$(valueText, 1) = "<" Then Exit Function

    numberValue = Val(valueText)
    If InStr(upperText, "MB") > 0 Or InStr(upperText, "MIB") > 0 Then
        numberValue = numberValue * 1048576#
    ElseIf InStr(upperText, "KB") > 0 Or InStr(upperText, "KIB") > 0 Then
        numberValue = numberValue * 1024#
    End If

    If numberValue < 0# Then numberValue = 0#
    If numberValue > 2147483647# Then numberValue = 2147483647#
    SizeTextToBytes = CLng(numberValue)
End Function

Private Function CleanIsoDisplayName(ByVal value As String) As String
    CleanIsoDisplayName = Trim$(value)
    If Right$(CleanIsoDisplayName, 1) = "\" Then
        CleanIsoDisplayName = Left$(CleanIsoDisplayName, Len(CleanIsoDisplayName) - 1)
    End If
    If CleanIsoDisplayName = "" Then CleanIsoDisplayName = "UNTITLED"
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

Private Function CountBackedTracks(ByVal tracks As Collection) As Long
    Dim i As Long
    Dim entry As TrackEntry

    For i = 1 To tracks.Count
        Set entry = tracks.Item(i)
        If Trim$(entry.FilePath) <> "" Then CountBackedTracks = CountBackedTracks + 1
    Next i
End Function

Private Function VerifyTrack(ByVal engine As ZenkiEngine, ByVal dllTrackNo As Long, ByVal entry As TrackEntry) As String
    Dim actualPath As String
    Dim expectedPath As String
    Dim actualPathHex As String
    Dim expectedFile As String
    Dim actualFile As String
    Dim expectedFlag As Long
    Dim actualFlag As Long
    Dim actualSectorCount As Long

    actualPath = NormalizePathText(engine.GetTrackPath(dllTrackNo))
    expectedPath = NormalizePathText(entry.FilePath)
    actualPathHex = engine.GetTrackPathHex(dllTrackNo, 96)
    expectedFile = FileNameOnly(expectedPath)
    actualFile = FileNameOnly(actualPath)
    expectedFlag = ZenkiTrackFormatFlag(entry.Source)
    actualFlag = engine.GetTrackControlFlag(dllTrackNo)
    actualSectorCount = engine.GetTrackSectorCount(dllTrackNo)

    If Not SameTrackPath(actualPath, expectedPath) Then
        VerifyTrack = "Track " & CStr(dllTrackNo) & " path mismatch." & vbCrLf & _
            "Expected file: " & expectedFile & vbCrLf & _
            "Actual file: " & actualFile & vbCrLf & _
            "Actual bytes: " & actualPathHex & vbCrLf & _
            "Actual path: " & actualPath
        Exit Function
    End If

    If engine.GetTrackPregapFrames(dllTrackNo) <> MsfToFramesBridge(entry.Pregap) Then
        VerifyTrack = "Track " & CStr(dllTrackNo) & " pregap mismatch."
        Exit Function
    End If

    If engine.GetTrackPostgapFrames(dllTrackNo) <> MsfToFramesBridge(entry.Postgap) Then
        VerifyTrack = "Track " & CStr(dllTrackNo) & " postgap mismatch."
        Exit Function
    End If

    If actualFlag <> expectedFlag Then
        VerifyTrack = "Track " & CStr(dllTrackNo) & " flag mismatch." & vbCrLf & _
            "Expected flag: " & CStr(expectedFlag) & vbCrLf & _
            "Actual flag: " & CStr(actualFlag) & vbCrLf & _
            "Source label: " & entry.Source
        Exit Function
    End If

    If actualSectorCount <= 0 Then
        VerifyTrack = "Track " & CStr(dllTrackNo) & " sector count invalid." & vbCrLf & _
            "Actual sector count: " & CStr(actualSectorCount)
        Exit Function
    End If

    If Not VerifyTextLanguage(engine, dllTrackNo, ZENKI_LANG_ENGLISH, entry.EnglishText, VerifyTrack) Then Exit Function
    If Not VerifyTextLanguage(engine, dllTrackNo, ZENKI_LANG_JAPANESE, entry.JapaneseText, VerifyTrack) Then Exit Function
End Function

Private Function VerifyTextLanguage(ByVal engine As ZenkiEngine, ByVal trackNo As Long, ByVal languageNo As Long, ByVal textItem As TOCCDText, ByRef mismatchText As String) As Boolean
    If textItem Is Nothing Then
        VerifyTextLanguage = True
        Exit Function
    End If

    If Not textItem.LanguageEnabled Then
        VerifyTextLanguage = True
        Exit Function
    End If

    If Not VerifyTextField(engine, trackNo, languageNo, ZENKI_INFO_TITLE, textItem.TitleEnabled, textItem.Title, mismatchText, "title") Then Exit Function
    If Not VerifyTextField(engine, trackNo, languageNo, ZENKI_INFO_PERFORMER, textItem.PerformerEnabled, textItem.Performer, mismatchText, "performer") Then Exit Function
    If Not VerifyTextField(engine, trackNo, languageNo, ZENKI_INFO_SONGWRITER, textItem.SongwriterEnabled, textItem.Songwriter, mismatchText, "songwriter") Then Exit Function
    If Not VerifyTextField(engine, trackNo, languageNo, ZENKI_INFO_COMPOSER, textItem.ComposerEnabled, textItem.Composer, mismatchText, "composer") Then Exit Function
    If Not VerifyTextField(engine, trackNo, languageNo, ZENKI_INFO_ARRANGER, textItem.ArrangerEnabled, textItem.Arranger, mismatchText, "arranger") Then Exit Function
    If Not VerifyTextField(engine, trackNo, languageNo, ZENKI_INFO_MESSAGE, textItem.MessageEnabled, textItem.Message, mismatchText, "message") Then Exit Function

    VerifyTextLanguage = True
End Function

Private Function VerifyTextField(ByVal engine As ZenkiEngine, ByVal trackNo As Long, ByVal languageNo As Long, ByVal informationNo As Long, ByVal enabled As Boolean, ByVal expectedValue As String, ByRef mismatchText As String, ByVal fieldName As String) As Boolean
    Dim actualValue As String

    If Not enabled Then
        VerifyTextField = True
        Exit Function
    End If

    actualValue = engine.GetText(trackNo, languageNo, informationNo)
    If StrComp(actualValue, expectedValue, vbBinaryCompare) <> 0 Then
        mismatchText = "Track " & CStr(trackNo) & " " & fieldName & " mismatch."
        Exit Function
    End If

    VerifyTextField = True
End Function

Private Function NormalizePathText(ByVal value As String) As String
    NormalizePathText = Replace(Trim$(value), "/", "\")
End Function

Private Function SameTrackPath(ByVal leftPath As String, ByVal rightPath As String) As Boolean
    If StrComp(leftPath, rightPath, vbTextCompare) = 0 Then
        SameTrackPath = True
        Exit Function
    End If

    If StrComp(FileNameOnly(leftPath), FileNameOnly(rightPath), vbTextCompare) = 0 Then
        SameTrackPath = True
    End If
End Function

Private Function FileNameOnly(ByVal value As String) As String
    Dim slashPos As Long

    slashPos = InStrRev(value, "\")
    If slashPos > 0 Then
        FileNameOnly = Mid$(value, slashPos + 1)
    Else
        FileNameOnly = value
    End If
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
