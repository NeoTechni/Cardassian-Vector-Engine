Attribute VB_Name = "API"
Option Explicit

'hPopupWnd = FindWindow("#32768", 0)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long

Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
'Public Const RDW_ALLCHILDREN = &H80
'Public Const RDW_UPDATENOW = &H100

Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Const MF_BYPOSITION As Long = &H400&
Private Const MF_STRING As Long = &H0&
Private Const MF_RIGHTJUSTIFY As Long = &H4000
Private Const MF_BITMAP = 4
Private Const MF_CHECKED = 8
Private Const MF_DISABLED = &H2&

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_GETITEMHEIGHT = &H1A1
Private Const CB_GETITEMHEIGHT = &H154

' Return the height of each entry in a ListBox or ComboBox control (in pixels)
Function GetListItemHeight(ctrl As Control) As Long
    Dim uMsg As Long
    If TypeOf ctrl Is ListBox Then
        uMsg = LB_GETITEMHEIGHT
    ElseIf TypeOf ctrl Is ComboBox Then
        uMsg = CB_GETITEMHEIGHT
    Else
        Exit Function
    End If
    GetListItemHeight = SendMessage(ctrl.hwnd, uMsg, 0, ByVal 0&)
End Function

' the second parameter in this example is the menu position. the first menu is at 0
Public Sub RightAlignMenu(FormHwnd As Long, Index As Long, Text As String)
    Dim hMenu As Long
    hMenu = GetMenu(FormHwnd)
    ModifyMenu hMenu, Index, MF_BYPOSITION Or MF_RIGHTJUSTIFY, MF_STRING, Text
    DrawMenuBar FormHwnd
End Sub

Public Function ToRealDegrees(Angle As Long) As Long
    ToRealDegrees = (CorrectAngle(-Angle) + 90) Mod 360
End Function

Public Function FlipNumber(Text As String, FullText As String, ParamArray FlipOn() As Variant) As String
    Dim temp As Long
    Text = killallexceptnumbers(Text)
    For temp = 0 To UBound(FlipOn)
        If ContainsText(FullText, CStr(FlipOn(temp))) Then
            Text = 1 - Val(Text)
            Exit For
        End If
    Next
    FlipNumber = Text
End Function
Public Function ContainsText(Text As String, Find As String) As Boolean
    ContainsText = InStr(1, Text, Find, vbTextCompare) > 0
End Function
Public Function RemoveAll(Text As String, ParamArray Remove() As Variant) As String
    Dim temp As Long
    For temp = 0 To UBound(Remove)
        Text = Replace(Text, CStr(Remove(temp)), "")
    Next
    RemoveAll = Text
End Function

Public Function Extension(ByVal Filename As String) As String
    Filename = Right(Filename, Len(Filename) - InStrRev(Filename, "\"))
    If InStr(Filename, ".") > 0 Then
        Extension = LCase(Right(Filename, Len(Filename) - InStrRev(Filename, ".")))
    End If
End Function
