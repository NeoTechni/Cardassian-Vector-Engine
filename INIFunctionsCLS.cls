VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'INIfile access

'CloseFile          Clears the current INI data from RAM
'LoadFile           Loads INI data from file
'SaveFile           Saves INI data to file

'CountSections      Counts the number of sections
'CreateSection      Creates a new section
'DeleteSection      Deletes a section and its contents
'EnumerateSections  Populate an array with a list of sections
'SectionExists      Checks if a section exists or not
'SectionAtIndex     Returns the section name at index

'CountKeys          Counts the numbers of keys in a section
'DeleteKey          Deletes a key in a section
'EnumerateKeys      Populates an array with a list of keys in a section and their values
'GetKeyValue        Gets the value of a key in a section
'KeyExists          Checks if a key exists in a section or not
'SetKeyValue        Sets the value of a key in a section, creates the section if it doesn't exist

Private Enum ENTRYTYPE
    Blank
    comment
    Section
    keyvalue
    equalless
End Enum

Private Type ENTRY
    Name As String
    Value As String
    eType As ENTRYTYPE
End Type

Private Type SECTIONENTRY
    Name As String
    Start As Long
    Finish As Long
    Keys As Long
End Type

Private INIfile() As ENTRY, ENTRYCOUNT As Long, SECTIONLIST() As SECTIONENTRY, SECTIONCOUNT As Long, CURRENTSECTION As String, CURRENTFILE As String

Public Function LoadFile(Optional Filename As String, Optional IgnoreEquals As Boolean) As Boolean
    On Error Resume Next
    Dim tempfile As Long, tempstr As String
    If Len(Filename) > 0 Then
        CloseFile
        CURRENTFILE = Filename
        If FileExists(Filename) = True Then
            tempfile = FreeFile
            Open Filename For Input As #tempfile
                Do Until EOF(tempfile)
                    Line Input #tempfile, tempstr
                    AddEntry tempstr, IgnoreEquals
                Loop
            Close #tempfile
            LoadFile = True
        End If
    End If
End Function

Public Sub CloseFile()
    ENTRYCOUNT = 0
    SECTIONCOUNT = 0
    CURRENTSECTION = ""
    CURRENTFILE = ""
    ReDim INIfile(0)
    ReDim SECTIONLIST(0)
End Sub

Public Sub SaveFile(Optional Filename As String)
    Dim tempfile As Long, temp As Long
    If ENTRYCOUNT > 0 Then
        tempfile = FreeFile
        If Len(Filename) = 0 Then Filename = CURRENTFILE
        If Len(Filename) > 0 Then
            Open Filename For Output As #tempfile
                For temp = 0 To ENTRYCOUNT - 1
                    With INIfile(temp)
                        Select Case .eType
                            Case comment: Print #tempfile, .Value
                            Case Section: Print #tempfile, "[" & .Name & "]"
                            Case keyvalue: Print #tempfile, .Name & "=" & .Value
                        End Select
                    End With
                Next
            Close #tempfile
        End If
    End If
End Sub

Public Function ToText() As String
    Dim temp As Long, tempstr As String
    If ENTRYCOUNT > 0 Then
        For temp = 0 To ENTRYCOUNT - 1
            With INIfile(temp)
                Select Case .eType
                    Case comment: tempstr = tempstr & .Value & vbNewLine
                    Case Section: tempstr = tempstr & "[" & .Name & "]" & vbNewLine
                    Case keyvalue: tempstr = tempstr & .Name & "=" & .Value & vbNewLine
                End Select
            End With
        Next
        ToText = tempstr
    End If
End Function

Private Function FileExists(Filename As String) As Boolean
    On Error Resume Next 'Checks to see if a file exists
    Dim temp As Long
    If Len(Filename) > 0 Then
        temp = GetAttr(Filename)
        FileExists = temp > 0 And FileLen(Filename) > 0
    End If
End Function

Private Function GetEntryType(ByVal Text As String, Optional IgnoreEquals As Boolean) As ENTRYTYPE
    GetEntryType = comment
    If Len(Text) > 2 Then
        If Left(Text, 1) = "[" And Right(Text, 1) = "]" Then
            GetEntryType = Section
        ElseIf InStr(Text, "=") > 0 Then
            GetEntryType = keyvalue
        ElseIf IgnoreEquals Then
            Text = Left(Text, 1)
            Select Case Text
                Case "#", "'", "!", "/", " "
                Case Else
                    GetEntryType = equalless
            End Select
        End If
    End If
End Function

Private Function AddEntry(ByVal Text As String, Optional IgnoreEquals As Boolean) As Long
    'Clean
    Text = LTrim(Text)
    Do Until Left(Text, 1) <> vbTab
        Text = Right(Text, Len(Text) - 1)
    Loop
    
    ENTRYCOUNT = ENTRYCOUNT + 1
    ReDim Preserve INIfile(ENTRYCOUNT)
    With INIfile(ENTRYCOUNT - 1)
        .eType = GetEntryType(Text, IgnoreEquals)
        AddEntry = .eType
        Select Case .eType
            Case comment
                AddKey , False
                .Value = Text
            Case Section
                .Name = Mid(Text, 2, Len(Text) - 2)
                CURRENTSECTION = .Name
                AddSection
            Case keyvalue
                .Name = Left(Text, InStr(Text, "=") - 1)
                .Value = Right(Text, Len(Text) - InStr(Text, "="))
                AddKey
            Case equalless
                If IgnoreEquals Then
                    .Name = Text
                    .Value = ""
                    .eType = keyvalue
                    AddKey
                End If
        End Select
    End With
End Function

Private Sub AddSection(Optional Name As String)
    If SECTIONCOUNT > 0 Then SECTIONLIST(SECTIONCOUNT - 1).Finish = ENTRYCOUNT - 2 'Assumes a new line was added with a new section before this was called
    SECTIONCOUNT = SECTIONCOUNT + 1
    ReDim Preserve SECTIONLIST(SECTIONCOUNT)
    With SECTIONLIST(SECTIONCOUNT - 1)
        If Len(Name) = 0 Then
            .Name = INIfile(ENTRYCOUNT - 1).Name
        Else
            .Name = Name
            ENTRYCOUNT = ENTRYCOUNT + 1
            ReDim Preserve INIfile(ENTRYCOUNT)
            With INIfile(ENTRYCOUNT - 1)
                .eType = Section
                .Name = Name
            End With
        End If
        .Start = ENTRYCOUNT - 1
        .Finish = ENTRYCOUNT - 1
    End With
End Sub

Private Sub AddKey(Optional Section As String, Optional isKey As Boolean = True)
    Dim temp As Long
    If Len(Section) = 0 Then
        With SECTIONLIST(SECTIONCOUNT - 1)
            If isKey Then .Keys = SECTIONLIST(SECTIONCOUNT - 1).Keys + 1
            .Finish = ENTRYCOUNT - 1
        End With
    Else
        temp = GetSectionIndex(Section)
        If temp > -1 Then
            With SECTIONLIST(temp)
                .Keys = .Keys + 1
                .Finish = .Finish + 1
            End With
            For temp = temp + 1 To SECTIONCOUNT - 1
                With SECTIONLIST(temp)
                    .Start = .Start + 1
                    .Finish = .Finish + 1
                End With
            Next
        End If
    End If
End Sub

Private Function GetSectionIndex(Section As String) As Long
    Dim temp As Long
    GetSectionIndex = -1
    If SECTIONCOUNT > 0 Then
        For temp = 0 To SECTIONCOUNT
            If StrComp(Section, SECTIONLIST(temp).Name, vbTextCompare) = 0 Then
                GetSectionIndex = temp
                Exit For
            End If
        Next
    End If
End Function

Private Function GetKeyIndex(Section As String, Key As String) As Long
    Dim temp As Long
    GetKeyIndex = -1
    temp = GetSectionIndex(Section)
    If temp > -1 Then
        With SECTIONLIST(temp)
            For temp = .Start To .Finish
                If INIfile(temp).eType = keyvalue Then
                    If StrComp(Key, INIfile(temp).Name, vbTextCompare) = 0 Then
                        GetKeyIndex = temp
                        Exit For
                    End If
                End If
            Next
        End With
    End If
End Function

Public Function GetClosestComment(Section As String, Key As String) As String
    Dim KeyIndex As Long
    KeyIndex = GetKeyIndex(Section, Key)
    GetClosestComment = GetClosestCommentTo(KeyIndex)
End Function

Public Function GetClosestCommentTo(KeyIndex As Long) As String
    If KeyIndex > -1 Then
        If INIfile(KeyIndex - 1).eType = comment Then
            GetClosestCommentTo = INIfile(KeyIndex - 1).Value
        End If
    End If
End Function

Public Sub DeleteClosestCommentTo(Section As String, Key As String)
    Dim KeyIndex As Long
    KeyIndex = GetKeyIndex(Section, Key)
    If KeyIndex > -1 Then
        If INIfile(KeyIndex - 1).eType = comment Then
            DeleteKeyIndex KeyIndex - 1
        End If
    End If
End Sub

Public Function SetKeyName(Section As String, Key As String, NewName As String) As Boolean
    Dim temp As Long, temp2 As Long
    temp = GetKeyIndex(Section, Key)
    temp2 = GetKeyIndex(Section, NewName)
    If temp > -1 And temp2 = -1 Then
        INIfile(temp).Name = NewName
        SetKeyName = True
    End If
End Function

Public Sub SetKeyValue(Section As String, Key As String, Optional Value As String)
    If Len(Section) = 0 Or Len(Key) = 0 Then Exit Sub
    Dim temp As Long, temp2 As Long
    temp = GetKeyIndex(Section, Key)
    If temp > -1 Then
        INIfile(temp).Value = Value
    Else
        temp = GetSectionIndex(Section)
        If temp = -1 Then
            CreateSection Section
            temp = SECTIONCOUNT - 1
        End If
        InsertKey SECTIONLIST(temp).Finish, Key, Value
        AddKey Section
    End If
End Sub

Private Sub InsertKey(Start As Long, Key As String, Value As String)
    Dim temp As Long
    ENTRYCOUNT = ENTRYCOUNT + 1
    ReDim Preserve INIfile(ENTRYCOUNT)
    For temp = ENTRYCOUNT - 2 To Start Step -1
        INIfile(temp + 1) = INIfile(temp)
    Next
    With INIfile(Start)
        .eType = keyvalue
        .Name = Key
        .Value = Value
    End With
End Sub

Public Function GetKeyValue(Section As String, Key As String, Optional Default As String) As String
    Dim temp As Long
    GetKeyValue = Default
    temp = GetKeyIndex(Section, Key)
    If temp > -1 Then GetKeyValue = INIfile(temp).Value
End Function

Public Function KeyExists(Section As String, Key As String) As Boolean
    KeyExists = GetKeyIndex(Section, Key) > -1
End Function

Public Sub CreateSection(Name As String)
    If GetSectionIndex(Name) = -1 And Len(Name) > 0 Then AddSection Name
End Sub

Public Sub DeleteSection(Section As String)
    Dim temp As Long
    temp = GetSectionIndex(Section)
    If temp > -1 Then
        With SECTIONLIST(temp)
            removerange .Start, .Finish
            ShiftUp temp + 1, .Finish + 1 - .Start
        End With
        For temp = temp To SECTIONCOUNT - 2
            SECTIONLIST(temp) = SECTIONLIST(temp + 1)
        Next
        SECTIONCOUNT = SECTIONCOUNT - 1
        ReDim Preserve SECTIONLIST(SECTIONCOUNT)
    End If
End Sub

Public Sub DeleteKey(Section As String, Key As String)
    Dim temp As Long, temp2 As Long
    temp = GetSectionIndex(Section)
    If temp > -1 Then
        temp2 = GetKeyIndex(Section, Key)
        If temp2 > -1 Then
            removerange temp2, temp2
            With SECTIONLIST(temp)
                .Finish = .Finish - 1
                .Keys = .Keys - 1
            End With
            ShiftUp temp + 1
        End If
    End If
End Sub

Private Function GetSectionFromIndex(Index As Long) As String
    Dim temp As Long
    For temp = Index - 1 To 0 Step -1
        If INIfile(temp).eType = Section Then
            GetSectionFromIndex = INIfile(temp).Name
            Exit For
        End If
    Next
End Function
Private Sub DeleteKeyIndex(Index As Long)
   Dim temp As Long, Section As String
   Section = GetSectionFromIndex(Index)
   temp = GetSectionIndex(Section)
   If temp > -1 And Index > -1 Then
        removerange Index, Index
        With SECTIONLIST(temp)
            .Finish = .Finish - 1
            .Keys = .Keys - 1
        End With
        ShiftUp temp + 1
   End If
End Sub

Private Sub ShiftUp(Section As Long, Optional count As Long = 1)
    Dim temp As Long
    If SECTIONCOUNT > Section + 1 Then
        For temp = Section To SECTIONCOUNT - 1
            With SECTIONLIST(temp)
                .Start = .Start - count
                .Finish = .Finish - count
            End With
        Next
    End If
End Sub

Private Sub removerange(Top As Long, Bottom As Long)
    Dim range As Long, count As Long
    If Top > 0 And Bottom > 0 And Top <= ENTRYCOUNT And Bottom <= ENTRYCOUNT And ENTRYCOUNT > 0 Then
        range = Bottom + 1 - Top
        'ENTRYCOUNT = ENTRYCOUNT - range
        'For count = Top To ENTRYCOUNT
        '    INIfile(count) = INIfile(count + range)
        'Next
        'If ENTRYCOUNT > 0 Then ReDim Preserve INIfile(1 To ENTRYCOUNT) Else ReDim INIfile(ENTRYCOUNT)
        For count = Top To Bottom
            INIfile(count).eType = Blank
        Next
    End If
End Sub

Public Function EnumerateSections(strarray() As String) As Long
    Dim temp As Long
    ReDim strarray(SECTIONCOUNT)
    If SECTIONCOUNT > 0 Then
        For temp = 0 To SECTIONCOUNT - 1
            strarray(temp) = SECTIONLIST(temp).Name
        Next
        EnumerateSections = SECTIONCOUNT
    End If
End Function

Public Function CountSections() As Long
    CountSections = SECTIONCOUNT
End Function

Public Function EnumerateKeys(Section As String, strarray() As String) As Long
    Dim temp As Long, temp2 As Long
    temp = GetSectionIndex(Section)
    If temp > -1 Then
        With SECTIONLIST(temp)
            ReDim strarray(0 To 1, .Keys)
            For temp = .Start To .Finish
                If INIfile(temp).eType = keyvalue Then
                    strarray(0, temp2) = INIfile(temp).Name
                    strarray(1, temp2) = INIfile(temp).Value
                    temp2 = temp2 + 1
                End If
            Next
        End With
        EnumerateKeys = temp2
    End If
End Function

Public Function CountKeys(Section As String) As Long
    Dim temp As Long
    temp = GetSectionIndex(Section)
    If temp > -1 Then CountKeys = SECTIONLIST(temp).Keys
End Function

Public Function SectionExists(Section As String) As Boolean
    SectionExists = GetSectionIndex(Section) > -1
End Function

Public Function SectionAtIndex(Index As Long) As String
    If Index > 0 And Index <= SECTIONCOUNT Then SectionAtIndex = SECTIONLIST(Index - 1).Name
End Function

Public Function CheckPath(Path As String, File As String) As String
    CheckPath = Path & IIf(Right(Path, 1) = "\", Empty, "\") & File
End Function
