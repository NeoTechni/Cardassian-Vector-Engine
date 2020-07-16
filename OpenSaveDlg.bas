Attribute VB_Name = "OpenSaveDlg"
Option Explicit

'Enum for the Flags of the BrowseForFolder API function
Enum BrowseForFolderFlags
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_EDITBOX = &H10
    BIF_RETURNFSANCESTORS = &H8
End Enum
'BrowseInfo is a type used with the SHBrowseForFolder API call
Private Type BROWSEINFO
     hwndOwner As Long
     pidlRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

'Shell APIs from Shell32.dll file:
'SHBrowseForFolder - Gets the Browse For Folder Dialog
Private Declare Function SHBrowseForFolder Lib "Shell32" (lpBI As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'lstrcat API function appends a string to another - that means that some API functions
'need their string in the numeric way like this does, so its kind of converts strings
'to numbers
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long




Public Const imageextentions As String = "*.bmp;*.gif;*.jpg;*.jpeg;*.jpe;*.jfif;*.png"
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
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
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
Public SaveFileDialog As OPENFILENAME
Public OpenFileDialog As OPENFILENAME
Private rv As Long
Private sv As Long
Private Enum CdlgExt_Flags
 cdlCCFullOpen = &H2
 cdlCCHelpButton = &H8
 cdlCCPreventFullOpen = &H4
 cdlCCRGBInit = &H1
End Enum
Private mFlags As CdlgExt_Flags

Private Type CHOOSECOLOR 'Color Dialog
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    RGBResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
'Color
Public Function ShowColor(mhOwner As Long, Optional mRGBResult As Long) As Long
  Dim CC As CHOOSECOLOR, CustomColors() As Byte, uFlag As Long, i As Long, RetValue As Long
  ReDim CustomColors(0 To 16 * 4 - 1) As Byte
  
  For i = LBound(CustomColors) To UBound(CustomColors)
     CustomColors(i) = 255 ' white
  Next i
  
  uFlag = mFlags And (&H1 Or &H2 Or &H4 Or &H8)
  
  With CC
       .lStructSize = Len(CC)
       .hwndOwner = mhOwner
       .hInstance = App.hInstance
       .lpCustColors = StrConv(CustomColors, vbUnicode)
       .Flags = uFlag
       .RGBResult = mRGBResult
       RetValue = ChooseColorAPI(CC)
       If RetValue = 0 Then
            ShowColor = -1
       Else
          CustomColors = StrConv(.lpCustColors, vbFromUnicode)
          mRGBResult = .RGBResult
            ShowColor = mRGBResult
       End If
  End With
End Function

Public Function Open_File(hWnd As Long) As String
   rv& = GetOpenFileName(OpenFileDialog)
   If (rv&) Then
      Open_File = Replace(Trim$(OpenFileDialog.lpstrFile), Chr(0), Empty)
   Else
      Open_File = ""
   End If
End Function

Public Function AutoSaveLoad(hWnd As Long, ByVal Filter As String, Optional Title As String, Optional InitDir As String, Optional Load As Boolean) As String
    Dim tempstr() As String, tempstr2 As String
    Filter = Replace(Filter, "|", Chr(0))
    If Load Then
        If Len(Title) = 0 Then Title = "Load file"
        InitOpen Filter, Title, InitDir
        AutoSaveLoad = Open_File(hWnd)
    Else
        If Len(Title) = 0 Then Title = "Save file"
        InitSave Filter, Title, InitDir
        If InStr(Filter, Chr(0)) > 0 Then
            tempstr = Split(Filter, Chr(0))
            Filter = Replace(tempstr(1), "*.", Empty)
        End If
        AutoSaveLoad = Save_File(hWnd, Filter)
    End If
End Function
Public Function Save_File(hWnd As Long, Optional defaultextention As String) As String
   sv& = GetSaveFileName(SaveFileDialog)
   Dim temp As String
   temp = ""
   If (sv&) Then
      temp = Trim$(SaveFileDialog.lpstrFile)
      temp = Left(temp, Len(temp) - 1)
      If InStrRev(temp, ".") = 0 And Len(defaultextention) > 0 Then temp = temp & "." & defaultextention
      If Dir(temp) <> Empty Then If MsgBox("File already exists. Do you wish to over write it?" & vbNewLine & temp, vbYesNo, "File exists") = vbNo Then temp = ""
      Save_File = temp
   End If
End Function

Public Sub InitSave(Filter As String, Title As String, Optional InitDir As String)
  With SaveFileDialog
     .lStructSize = Len(SaveFileDialog)

     .hInstance = App.hInstance
     .lpstrFilter = Filter
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = IIf(InitDir <> Empty, InitDir, CurDir)
     .lpstrTitle = Title
     .Flags = 0
  End With
End Sub
Public Sub InitOpen(Filter As String, Title As String, Optional InitDir As String)
   Filter = Replace(Filter, "|", Chr(0))
   With OpenFileDialog
     .lStructSize = Len(OpenFileDialog)

     .hInstance = App.hInstance
     .lpstrFilter = Filter
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = IIf(InitDir <> Empty, InitDir, CurDir)
     .lpstrTitle = Title
     .Flags = 0
   End With
End Sub

Public Function BrowseForFolder(hWnd As Long, Optional Title As String, Optional Flags As BrowseForFolderFlags) As String
On Error Resume Next
    'Variables for use:
     Dim iNull As Integer
     Dim IDList As Long
     Dim Result As Long
     Dim Path As String
     Dim bi As BROWSEINFO
     
     If Flags = 0 Then Flags = BIF_RETURNONLYFSDIRS
     
    'Type Settings
     With bi
        .hwndOwner = hWnd
        .lpszTitle = lstrcat(Title, "")
        .ulFlags = Flags
     End With

    'Execute the BrowseForFolder shell API and display the dialog
     IDList = SHBrowseForFolder(bi)
     
    'Get the info out of the dialog
     If IDList Then
        Path = String$(300, 0)
        Result = SHGetPathFromIDList(IDList, Path)
        iNull = InStr(Path, vbNullChar)
        If iNull Then Path = Left$(Path, iNull - 1)
     End If

    'If Cancel button was clicked, error occured or My Computer was selected then Path = ""
     BrowseForFolder = Path
End Function
