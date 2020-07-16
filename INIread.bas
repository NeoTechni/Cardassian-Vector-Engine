Attribute VB_Name = "INIread"
Option Explicit


Public Function issection(Value As String) As Boolean
    On Error Resume Next
    If Left(Value, 1) = "[" And Right(Value, 1) = "]" And stripsection(Value) <> Empty Then issection = True Else issection = False
End Function
Public Function isvalue(Value As String) As Boolean
    On Error Resume Next
    If issection(Value) = False And InStr(Value, "=") > 0 Then isvalue = True Else isvalue = False
End Function
Public Function stripsection(Section As String) As String
    On Error Resume Next
    stripsection = Mid(Section, 2, Len(Section) - 2)
End Function
Public Function stripvalue(Value As String) As String
    On Error Resume Next
    stripvalue = Right(Value, Len(Value) - InStr(Value, "="))
End Function
Public Function stripname(Value As String) As String
    On Error Resume Next
    stripname = Left(Value, InStr(Value, "=") - 1)
End Function
Public Function iscomment(Value As String) As Boolean
    On Error Resume Next
    If Left(Value, 1) = "#" Or Left(Value, 1) = "'" Then iscomment = True Else iscomment = False
End Function









Public Function LoadFile(Filename As String) As String
    On Error Resume Next
    Dim tempfile As Long, tempstr As String, CURRENTSECTION As String
    If FileExists(Filename) = True Then
        tempfile = FreeFile
        Open Filename For Input As #tempfile
            Do Until EOF(tempfile) Or Found = True
                Line Input #tempfile, temp
                If iscomment(temp) = False Then
                    If issection(temp) = True Then
                        CURRENTSECTION = LCase(stripsection(temp))
                    Else
                        If CURRENTSECTION = LCase(stripsection(Section)) Then
                            If isvalue(temp) Then
                                If LCase(stripname(temp)) = LCase(Value) Then
                                    getvalue = stripvalue(temp)
                                    Found = True
                                End If
                            End If
                        End If
                    End If
                End If
            Loop
        Close #tempfile
    End If
End Function

Public Function FileExists(Filename As String) As Boolean
    On Error Resume Next 'Checks to see if a file exists
    Dim temp As Long
    temp = GetAttr(Filename)
    FileExists = temp > 0
End Function

Public Sub SaveFile(Filename As String, Text As String, Optional Append As Boolean)
    Dim tempfile As Integer
    tempfile = FreeFile
    If Append Then
        Open Filename For Append As #tempfile
    Else
        Open Filename For Output As #tempfile
    End If
    Print #tempfile, Text
    Close #tempfile
End Sub
