VERSION 5.00
Begin VB.Form SVGedit 
   Caption         =   "SVG editor"
   ClientHeight    =   9480
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   ScaleHeight     =   632
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1034
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerTop 
      Interval        =   100
      Left            =   14760
      Top             =   120
   End
   Begin VB.ListBox lstlinenumbers 
      Appearance      =   0  'Flat
      Height          =   4710
      Left            =   12240
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.HScrollBar hscmain 
      Height          =   255
      LargeChange     =   15
      Left            =   12240
      Max             =   359
      TabIndex        =   5
      Top             =   9225
      Width           =   3135
   End
   Begin VB.PictureBox picpreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   12240
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   4
      ToolTipText     =   "Use the scrollbar below to change the angle of rotation"
      Top             =   4800
      Width           =   3000
   End
   Begin VB.PictureBox picsec 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   360
      ScaleHeight     =   495
      ScaleWidth      =   1335
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstmain 
      Appearance      =   0  'Flat
      Height          =   4710
      Left            =   12840
      TabIndex        =   2
      ToolTipText     =   "Double-click to edit a point, Right click to seek to a specific point by it's number"
      Top             =   0
      Width           =   2535
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9000
      Left            =   240
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   240
      Width           =   9000
      Begin VB.Shape shpcursor 
         BorderColor     =   &H8000000D&
         Height          =   855
         Left            =   120
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox picborder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   9480
      Left            =   0
      ScaleHeight     =   630
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   630
      TabIndex        =   1
      Top             =   0
      Width           =   9480
      Begin VB.Line Line2 
         BorderColor     =   &H8000000D&
         X1              =   504
         X2              =   464
         Y1              =   48
         Y2              =   152
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000D&
         X1              =   88
         X2              =   216
         Y1              =   112
         Y2              =   112
      End
      Begin VB.Shape shpmain 
         Height          =   9030
         Left            =   210
         Top             =   210
         Width           =   9030
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuclear 
         Caption         =   "Clear SVG"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save SVG to Clipboard"
      End
      Begin VB.Menu mnuload 
         Caption         =   "Load SVG from Clipboard"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Append SVG from Clipboard"
      End
      Begin VB.Menu mnusplit 
         Caption         =   "Split SVG to Clipboard"
      End
      Begin VB.Menu mnufilesep 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileop 
         Caption         =   "New File"
         Index           =   0
      End
      Begin VB.Menu mnufileop 
         Caption         =   "Load File"
         Index           =   1
      End
      Begin VB.Menu mnufileop 
         Caption         =   "Save File"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnufileop 
         Caption         =   "Save File As"
         Enabled         =   0   'False
         Index           =   3
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnuedittool 
         Caption         =   "Copy Vertical"
         Index           =   0
      End
      Begin VB.Menu mnuedittool 
         Caption         =   "Copy Horizontal"
         Index           =   1
      End
      Begin VB.Menu mnuedittool 
         Caption         =   "Ignore the boundaries"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu chkmain 
         Caption         =   "Show Tracer"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuusetracer 
         Caption         =   "Use Clipboard as Tracer"
      End
      Begin VB.Menu mnuloadtracer 
         Caption         =   "Load File as Tracer"
      End
      Begin VB.Menu mnucopy 
         Caption         =   "Copy points"
      End
      Begin VB.Menu mnuAspect 
         Caption         =   "Square Aspect Ratio"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDoGrid 
         Caption         =   "Grid lines"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "Tools (Add Point)"
      Index           =   0
      Begin VB.Menu mnumode 
         Caption         =   "Add Point"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnumode 
         Caption         =   "Move Point"
         Index           =   1
      End
      Begin VB.Menu mnumode 
         Caption         =   "Region Select"
         Index           =   2
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "Mode"
      Index           =   1
      Begin VB.Menu mnutool 
         Caption         =   "Line"
         Index           =   0
      End
      Begin VB.Menu mnutool 
         Caption         =   "Bezier Curve"
         Index           =   1
      End
      Begin VB.Menu mnutool 
         Caption         =   "Remove Clip"
         Index           =   2
      End
      Begin VB.Menu mnutool 
         Caption         =   "Oval"
         Index           =   4
      End
      Begin VB.Menu mnutool 
         Caption         =   "Break"
         Index           =   5
      End
      Begin VB.Menu mnutool 
         Caption         =   "Bicubic curve"
         Index           =   6
      End
      Begin VB.Menu mnumoretools 
         Caption         =   "-"
      End
      Begin VB.Menu mnumore 
         Caption         =   "Text"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnumore 
         Caption         =   "Change LCARS Color"
         Index           =   1
         Begin VB.Menu mnucolor 
            Caption         =   "Black"
            Index           =   0
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Dark Orange"
            Index           =   1
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Orange"
            Index           =   2
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Light Orange"
            Index           =   3
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Purple"
            Index           =   4
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Light Purple"
            Index           =   5
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Light Blue"
            Index           =   6
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Red"
            Index           =   7
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Yellow"
            Index           =   8
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Dark Blue"
            Index           =   9
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Dark Yellow"
            Index           =   10
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Dark Purple"
            Index           =   11
         End
         Begin VB.Menu mnucolor 
            Caption         =   "White"
            Index           =   12
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Red Alert"
            Index           =   13
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Light Green"
            Index           =   14
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Green"
            Index           =   15
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Lighter Blue"
            Index           =   16
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Blue"
            Index           =   17
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Turq"
            Index           =   18
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Classic Red"
            Index           =   19
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Grey"
            Index           =   20
         End
         Begin VB.Menu mnucolor 
            Caption         =   "LBlue"
            Index           =   21
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Light Yellow"
            Index           =   22
         End
         Begin VB.Menu mnucolor 
            Caption         =   "BORG"
            Index           =   23
         End
         Begin VB.Menu mnucolor 
            Caption         =   "Chrono"
            Index           =   24
         End
         Begin VB.Menu mnucolor 
            Caption         =   "tDark Turq"
            Index           =   25
         End
         Begin VB.Menu mnucolor 
            Caption         =   "tLight Turq"
            Index           =   26
         End
         Begin VB.Menu mnucolor 
            Caption         =   "tYellow"
            Index           =   27
         End
         Begin VB.Menu mnucolor 
            Caption         =   "tOrange"
            Index           =   28
         End
         Begin VB.Menu mnucolor 
            Caption         =   "tDark Orange"
            Index           =   29
         End
         Begin VB.Menu mnucolor 
            Caption         =   "tDarkPurple"
            Index           =   30
         End
         Begin VB.Menu mnucolor 
            Caption         =   "tLightPurple"
            Index           =   31
         End
      End
      Begin VB.Menu mnumore 
         Caption         =   "Change LCARS Font"
         Index           =   2
         Begin VB.Menu mnufont 
            Caption         =   "LCARS"
            Index           =   0
         End
         Begin VB.Menu mnufont 
            Caption         =   "TOS"
            Index           =   1
         End
         Begin VB.Menu mnufont 
            Caption         =   "Enterprise"
            Index           =   2
         End
         Begin VB.Menu mnufont 
            Caption         =   "Motion Picture"
            Index           =   3
         End
         Begin VB.Menu mnufont 
            Caption         =   "Klingon"
            Index           =   4
         End
         Begin VB.Menu mnufont 
            Caption         =   "Star Wars"
            Index           =   6
         End
         Begin VB.Menu mnufont 
            Caption         =   "ChronowerX"
            Index           =   7
         End
         Begin VB.Menu mnufont 
            Caption         =   "Romulan"
            Index           =   8
         End
         Begin VB.Menu mnufont 
            Caption         =   "Cardassian"
            Index           =   14
         End
      End
      Begin VB.Menu mnumore 
         Caption         =   "Change Stroke"
         Index           =   3
      End
   End
   Begin VB.Menu mnusvgs 
      Caption         =   "SVGs"
      Begin VB.Menu mnusvgdelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnusvgdelete 
         Caption         =   "Discard Changes"
         Index           =   1
      End
      Begin VB.Menu mnusvgdelete 
         Caption         =   "New SVG"
         Index           =   2
      End
      Begin VB.Menu mnusvgdelete 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnusvg 
         Caption         =   "0"
         Index           =   0
      End
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Options"
      Begin VB.Menu mnuoption 
         Caption         =   "Preview on Right side"
         Checked         =   -1  'True
         Index           =   0
      End
   End
   Begin VB.Menu mnutest 
      Caption         =   "RIGHT ALIGN"
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "SVGedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const mnutestIndex As Integer = 5

Dim PointsX As New Collection, SelectedTool As Integer, isdown As Boolean, INIfile As New INI, CurrentSVG As Integer, CurrentTracer As String
Dim IsSet As Boolean, LastX As Single, LastY As Single, CurrentMode As Long, SquareAspect As Boolean, CurrentTop As Integer
Dim Points(4) As POINTSNG, FirstPoint As Long, PreviousPoint As Long, CurrentAngle As Long, DoGrid As Boolean, m_PNG As New LoadPNG
Dim DownX As Long, DownY As Long

Private Sub Form_Load()
    SquareAspect = True
    DoGrid = True
    SaveStatus False
    RightAlignMenu Me.hwnd, mnutestIndex, ""
    CurrentTop = -1
End Sub

Private Sub lstmain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempstr As String
    If Button = vbRightButton Then
        tempstr = killallexceptnumbers(InputBox("What point would you like to select?", "Select a point", lstmain.ListIndex))
        If Len(tempstr) > 0 Then
            If tempstr > -1 And tempstr < lstmain.ListCount Then
                lstmain.ListIndex = tempstr
            End If
        End If
    End If
End Sub

Private Sub mnuAspect_Click()
    If IsSet Then
        mnuAspect.Checked = Not mnuAspect.Checked
        SquareAspect = mnuAspect.Checked
        Form_Resize
    End If
End Sub

Private Sub mnucopy_Click()
    Dim tempstr As String, temp As Long
    For temp = 0 To lstmain.ListCount - 1
        If Len(tempstr) = 0 Then
            tempstr = lstmain.List(temp)
        Else
            tempstr = tempstr & "," & lstmain.List(temp)
        End If
    Next
    Clipboard.Clear
    Clipboard.SetText tempstr
    'Debug.Print tempstr
End Sub

Private Sub cmdmain_Click()
    AddPoint -1, 0, False
End Sub

Private Sub cmdremove_Click()
    Dim temp As Long
    If lstmain.ListIndex > -1 Then
        temp = lstmain.ListIndex - 1
        PointsX.Remove lstmain.ListIndex + 1
        lstmain.RemoveItem lstmain.ListIndex
        lstmain.ListIndex = temp
        DrawPoints picmain
    End If
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    Dim temp As Long, Size As Long, Index As Long, Width As Long, Height As Long, ScaleHeight2 As Long
    Dim LeftSide As Long, RightSide As Long
    
    Height = ScaleHeight - lstmain.Height
    If mnuoption(0).Checked Then 'right side
        RightSide = ScaleWidth - Height
    Else
        LeftSide = Height
    End If
    
    If ScaleHeight = 0 Or ScaleWidth = 0 Then Exit Sub
        
    Height = ScaleHeight - picmain.Top * 2
    Width = Height
    If Not SquareAspect And IsSet Then
        Width = Height * picsec.Width / picsec.Height
        temp = ScaleHeight - lstmain.Height
        temp = ScaleWidth - temp - picmain.Top * 2
        If Width > temp Then
            Width = temp
            Height = Width * picsec.Height / picsec.Width
        End If
    End If
    
    picmain.Move LeftSide + picmain.Top, picmain.Top, Width, Height
    ScaleHeight2 = Height + picmain.Top * 2
    picborder.Move LeftSide, 0, picmain.Top * 2 + picmain.Width, ScaleHeight2
    shpcursor.Move LeftSide + picmain.Top - 2, picmain.Top - 2, picmain.Width + 2, picmain.Height + 2
    
    Height = ScaleHeight - lstmain.Height
    picpreview.Move RightSide, lstmain.Height, Height, Height
    hscmain.Move RightSide, ScaleHeight - hscmain.Height, picpreview.Width
    lstlinenumbers.Move RightSide, 0, lstlinenumbers.Width, lstmain.Height
    lstmain.Move RightSide + lstlinenumbers.Width, 0, Height - lstlinenumbers.Width
    
    DrawPoints picmain
    picborder.Cls
    Size = picmain.Width * 0.05
    For temp = -1 To picmain.Width Step Size
        Index = Index + 1
        Width = TextWidth(CStr(Index)) * 0.5
        Height = TextHeight(CStr(Index)) * 0.5
        
        DrawLine picborder, picmain.Top + temp, 0, picmain.Top + temp, picmain.Top, vbBlack, 1
        DrawLine picborder, picmain.Top + temp, picmain.Top + picmain.Height, picmain.Top + temp, picborder.Height, vbBlack, 1
        
        'DrawLine picborder, 0, picmain.Left + temp, picmain.Top, picmain.Left + temp, vbBlack, 1
        'DrawLine picborder, picmain.Top + picmain.Height, picmain.Left + temp, picborder.Height, picmain.Left + temp, vbBlack, 1
        
        If Index < 21 Then
            PrintText picborder, picmain.Top + temp + Size * 0.5 - Width, picmain.Top * 0.5 - Height, CStr(Index)
            PrintText picborder, picmain.Top + temp + Size * 0.5 - Width, picmain.Height + (picmain.Top * 1.5) - Height, CStr(Index)
            
            'PrintText picborder, picmain.Left * 0.5 - Width, picmain.Top + temp + Size * 0.5 - Height, CStr(Index)
            'PrintText picborder, picmain.Left * 1.5 + picmain.Width - Width, picmain.Top + temp + Size * 0.5 - Height, CStr(Index)
        End If
    Next
    
    
    Index = 0
    Size = picmain.Height * 0.05
    For temp = -1 To picmain.Height Step Size
        Index = Index + 1
        Width = TextWidth(CStr(Index)) * 0.5
        Height = TextHeight(CStr(Index)) * 0.5
    
        'DrawLine picborder, picmain.Left + temp, 0, picmain.Left + temp, picmain.Top, vbBlack, 1
        'DrawLine picborder, picmain.Left + temp, picmain.Top + picmain.Height, picmain.Left + temp, picborder.Height, vbBlack, 1

        DrawLine picborder, 0, picmain.Top + temp, picborder.Width, picmain.Top + temp, vbBlack, 1

        'DrawLine picborder, 0, picmain.Left + temp, picmain.Top, picmain.Width, vbBlack, 1
        'DrawLine picborder, picmain.Top + picmain.Height, picmain.Left + temp, picborder.Height, picmain.Left + temp, vbBlack, 1
    
        If Index < 21 Then
            'PrintText picborder, picmain.Left + temp + Size * 0.5 - Width, picmain.Top * 0.5 - Height, CStr(Index)
            'PrintText picborder, picmain.Left + temp + Size * 0.5 - Width, picmain.Height + (picmain.Top * 1.5) - Height, CStr(Index)
    
            PrintText picborder, picmain.Top * 0.5 - Width, picmain.Top + temp + Size * 0.5 - Height, CStr(Index)
            PrintText picborder, picmain.Top * 1.5 + picmain.Width - Width, picmain.Top + temp + Size * 0.5 - Height, CStr(Index)
        End If
    Next
End Sub

Private Sub hscmain_Change()
    DrawPoints picpreview, -1, -1, hscmain.Value
End Sub

Private Sub hscmain_Scroll()
    hscmain_Change
End Sub

Private Sub lstmain_Click()
    DrawPoints picmain
    lstlinenumbers.ListIndex = lstmain.ListIndex - lstmain.TopIndex
End Sub

Private Sub lstmain_DblClick()
    Dim tempstr As String, StartPoint As String, tempstr2() As String, tempstr3 As String, tempstr4() As Double
    StartPoint = lstmain.List(lstmain.ListIndex)
    tempstr = InputBox("Edit point: " & lstmain.ListIndex + 1 & vbNewLine & "If the text contains:" & vbNewLine & "    X: it will be flipped horizontally" & vbNewLine & "    Y: it will be flipped vertically" & vbNewLine & "    No comma/Only 1 number: It will get the data from the point you specify", "EDIT POINT", StartPoint)
    If Len(tempstr) > 0 And tempstr <> StartPoint Then
        If Not ContainsText(tempstr, ",") Then
            tempstr3 = killallexceptnumbers(tempstr)
            If Len(tempstr3) > 0 Then
                If tempstr3 > -1 And tempstr3 < PointsX.count Then
                    tempstr4 = PointsX.Item(Val(tempstr3))
                    tempstr = Replace(tempstr, tempstr3, tempstr4(0) & "," & tempstr4(1))
                Else
                    MsgBox tempstr3 & " is out of range", vbCritical
                    Exit Sub
                End If
            End If
        End If
        tempstr2 = Split(tempstr, ",")
        tempstr2(0) = FlipNumber(tempstr2(0), tempstr, "h", "x")
        tempstr2(1) = FlipNumber(tempstr2(1), tempstr, "v", "y")
        lstmain.List(lstmain.ListIndex) = RemoveAll(tempstr, "h", "v", "x", "y")
        SetCollection PointsX, lstmain.ListIndex + 1, MakeArray(tempstr2(0), tempstr2(1))
        DrawPoints picmain
    End If
End Sub



Private Sub mnuclear_Click()
    lstmain.Clear
    ClearCollection PointsX
    DrawPoints picmain
End Sub

Private Sub mnuDoGrid_Click()
    DoGrid = Not DoGrid
    mnuDoGrid.Checked = DoGrid
    DrawPoints picmain
End Sub

Private Sub mnucolor_Click(Index As Integer)
    AddPoint -2, CSng(Index), False
End Sub


Private Sub mnuedittool_Click(Index As Integer)
    Dim temp As Long, Data() As Double
    Select Case Index
        Case 0 'copy vertical
            For temp = PointsX.count To 1 Step -1
                Data = PointsX.Item(temp)
                If Data(0) >= 0 Then
                    Data(1) = 1 - Data(1)
                End If
                AddPoint Data(0), Data(1), False
            Next
        Case 1 'copy horizontal
            For temp = PointsX.count To 1 Step -1
                Data = PointsX.Item(temp)
                If Data(0) >= 0 Then
                    Data(0) = 1 - Data(0)
                End If
                AddPoint Data(0), Data(1), False
            Next
        Case 2 'boundary
            mnuedittool(1).Checked = Not mnuedittool(1).Checked
    End Select
End Sub

Private Sub mnufileop_Click(Index As Integer)
    Dim tempstr As String
    Select Case Index
        Case 0 'clear
            INIfile.CloseFile
            INIfile.CreateSection "svgs"
            INIfile.SetKeyValue "svgs", "0", ""
            SaveStatus True, True
        Case 1 'load
            tempstr = AutoSaveLoad(Me.hwnd, "INI File|*.ini", "Load File", , True)
            If INIfile.LoadFile(tempstr) Then SaveStatus True
        Case 2 'save
            INIfile.SaveFile
            MsgBox "SVGs saved to: " & tempstr, vbInformation
        Case 3 'save as
            tempstr = AutoSaveLoad(Me.hwnd, "INI File|*.ini", "Save File")
            If Len(tempstr) > 0 Then
                INIfile.SaveFile tempstr
                MsgBox "SVGs saved to: " & tempstr, vbInformation
            End If
    End Select
End Sub
Sub SaveStatus(Status As Boolean, Optional IsNew As Boolean)
    Dim count As Long, temp As Long, tempstr As String
    For temp = mnusvg.UBound To 1 Step -1
        Unload mnusvg(temp)
    Next
    CurrentSVG = -1
    If Status And Not IsNew Then
        If INIfile.SectionExists("svgs") Then
            Do Until Not INIfile.KeyExists("svgs", CStr(temp))
                If temp > 0 Then Load mnusvg(temp)
                With mnusvg(temp)
                    .Enabled = True
                    .Caption = temp
                    .Checked = False
                End With
                tempstr = INIfile.GetClosestComment("svgs", CStr(temp))
                If Len(tempstr) > 0 Then mnusvg(temp).Caption = temp & ": " & Right(tempstr, Len(tempstr) - 1)
                temp = temp + 1
            Loop
        Else
            Status = False
            MsgBox "This is not a valid SVG INI file", vbCritical
        End If
    ElseIf IsNew Then
        mnusvg(0).Enabled = True
        mnusvg(0).Caption = 0
        mnusvg(0).Checked = True
        CurrentSVG = 0
    End If
    mnufileop(2).Enabled = Status
    mnufileop(3).Enabled = Status
    mnusvgs.Visible = Status
End Sub

Private Sub mnuoption_Click(Index As Integer)
    Select Case Index
        Case 0 'right side
            mnuoption(Index).Checked = Not mnuoption(Index).Checked
    End Select
    
    Select Case Index
        Case 0 'right side
            Form_Resize
    End Select
End Sub

Private Sub mnusvg_Click(Index As Integer)
    Dim tempstr As String
    If CurrentSVG > -1 Then
        INIfile.SetKeyValue "svgs", CStr(CurrentSVG), SaveFile
        mnusvg(CurrentSVG).Checked = False
    End If
    CurrentSVG = Index
    tempstr = INIfile.GetKeyValue("svgs", CStr(Index))
    LoadFile tempstr
    DeleteState True, Index
    mnusvg(CurrentSVG).Checked = True
End Sub

Private Sub mnufont_Click(Index As Integer)
    AddPoint -3, CDbl(Index), False
End Sub

Private Sub mnuload_Click()
    loadcliporfile "Load File"
End Sub

Public Function loadcliporfile(Title As String, Optional Clear As Boolean = True) As String
    Dim tempstr As String
    tempstr = Clipboard.GetText
    If Len(tempstr) > 0 Then
        If MsgBox("Use clipboard?" & vbNewLine & vbNewLine & tempstr, vbYesNo, Title) = vbNo Then tempstr = Empty
    End If
    If Len(tempstr) = 0 Then
        tempstr = InputBox("What is the file?", Title, tempstr)
    End If
    If Len(tempstr) > 0 Then LoadFile tempstr, Clear
    loadcliporfile = tempstr
End Function

Private Sub mnuAdd_Click()
    loadcliporfile "Append File", False
End Sub


Sub LoadFile(Contents As String, Optional Clear As Boolean = True)
    Dim tempstr() As String, temp As Long
    If ContainsText(Contents, vbNewLine) Or ContainsText(Contents, "=") Then
        tempstr = Split(Contents, vbNewLine)
        For temp = 0 To UBound(tempstr)
            Contents = Trim(tempstr(temp))
            If ContainsText(Contents, "=") Then
                Contents = Right(Contents, Len(Contents) - InStr(Contents, "="))
            ElseIf ContainsText(Contents, "#") Then
                Contents = ""
            End If
            tempstr(temp) = Contents
        Next
        Contents = Join(tempstr, "|-1,5|")
    End If
    tempstr = Split(Contents, "|")
    If Clear Then
        mnuclear_Click
    Else
        PointsX.Add splitF("-1,5")
        lstmain.AddItem "-1,5"
    End If
    For temp = 0 To UBound(tempstr)
        PointsX.Add splitF(tempstr(temp))
        lstmain.AddItem tempstr(temp)
    Next
    lstmain.ListIndex = lstmain.ListCount - 1
    DrawPoints picmain
End Sub

Function splitF(Text As String) As Double()
    Dim Data() As Double, tempstr() As String, temp As Long
    tempstr = Split(Text, ",")
    ReDim Preserve Data(UBound(tempstr) + 1)
    For temp = 0 To UBound(tempstr)
        Data(temp) = CSng(killallexceptnumbers(tempstr(temp)))
    Next
    splitF = Data
End Function

Private Sub mnuloadtracer_Click()
    Dim tempstr As String
    tempstr = AutoSaveLoad(Me.hwnd, "Image files" & Chr(0) & "*.jpg;*.png;*.bmp;*.gif;*.jpeg", "Load an image", , True)
    If Len(tempstr) > 0 Then
        If Extension(tempstr) = "png" Then
            m_PNG.PicBox = picsec
            m_PNG.OpenPNG tempstr
        Else
            picsec.Picture = LoadPicture(tempstr)
        End If
        CurrentTracer = tempstr
        IsSet = True
        DrawPoints picmain
    End If
End Sub

Private Sub mnumode_Click(Index As Integer)
    shpcursor.Visible = Index = 2
    mnutools(0).Caption = "Tools (" & mnumode(Index).Caption & ")"
    mnumode(SelectedTool).Checked = False
    mnumode(Index).Checked = True
    SelectedTool = Index
End Sub

Private Sub mnumore_Click(Index As Integer)
    Dim tempstr As String, Data() As Double
    Select Case Index
        Case 0 'text
            tempstr = Left(InputBox("What letter?"), 1)
            If Len(tempstr) = 1 Then Data = MakeArray(-1, 3, Asc(tempstr))
        
        'case 1,2 color/font
        
        Case 3 'stroke
            tempstr = InputBox("What Stroke? (0 = solid, <0 = clip path)")
            If IsNumeric(tempstr) And Len(tempstr) > 0 Then Data = MakeArray(-5, tempstr)
        
        Case Else
            Exit Sub
    End Select
    
    PointsX.Add Data
    lstmain.AddItem JoinSng(Data)
    DrawPoints picmain
End Sub

Function JoinSng(Data() As Double) As String
    Dim tempstr As String, temp As Long
    tempstr = Data(0)
    For temp = 1 To UBound(Data) - 1
        tempstr = tempstr & "," & Data(temp)
    Next
    JoinSng = tempstr
End Function


Private Sub mnusave_Click()
    Dim tempstr As String
    tempstr = SaveFile
   ' Debug.Print tempstr
    Clipboard.Clear
    Clipboard.SetText tempstr
End Sub

Public Function SaveFile() As String
    Dim tempstr As String, temp As Long, tempstr2 As String, Data() As Double
    For temp = 1 To PointsX.count
        Data = PointsX.Item(temp)
        tempstr2 = JoinSng(Data)
        If temp = 1 Then
            tempstr = tempstr2
        Else
            tempstr = tempstr & "|" & tempstr2
        End If
    Next
    SaveFile = tempstr
End Function

Private Sub mnusplit_Click()
    Dim tempstr As String, StartingNumber As String, CurrentLine As Long, Lines() As String, temp As Long
    StartingNumber = Trim(InputBox("What would you like the starting line number to be?" & vbNewLine & "(Leave blank for no line numbers)", "Starting Line Number"))
    tempstr = Replace(SaveFile, "|-1,5|", vbNewLine)
    If Len(StartingNumber) > 0 Then
        CurrentLine = killallexceptnumbers(StartingNumber)
        Lines = Split(tempstr, vbNewLine)
        For temp = 0 To UBound(Lines)
            Lines(temp) = CurrentLine & "=" & Lines(temp)
            CurrentLine = CurrentLine + 1
        Next
        tempstr = Join(Lines, vbNewLine)
    End If
    Clipboard.Clear
    Clipboard.SetText tempstr
End Sub

Private Sub mnusvgdelete_Click(Index As Integer)
    Select Case Index
        Case 0 'delete
            If MsgBox("Are you sure you want to delete SVG '" & mnusvg(CurrentSVG).Caption & "'?", vbQuestion, "Delete SVG") = vbYes Then
                INIfile.DeleteClosestCommentTo "svgs", CStr(CurrentSVG)
                INIfile.DeleteKey "svgs", CStr(CurrentSVG)
                CurrentSVG = CurrentSVG + 1
                Do Until Not INIfile.SetKeyName("svgs", CStr(CurrentSVG), CStr(CurrentSVG - 1))
                    CurrentSVG = CurrentSVG + 1
                Loop
                SaveStatus True
            End If
        Case 1 'discard changes
            DeleteState False
            CurrentSVG = -1
        Case 2 'new SVG
            Load mnusvg(mnusvg.UBound + 1)
            With mnusvg(mnusvg.UBound)
                .Caption = mnusvg.UBound
                .Enabled = True
                .Visible = True
                .Checked = False
            End With
            INIfile.SetKeyValue "svgs", CStr(mnusvg.UBound), ""
    End Select
End Sub
Public Sub DeleteState(State As Boolean, Optional Index As Integer = -1)
    mnusvgdelete(0).Enabled = State And Index > 0
    mnusvgdelete(1).Enabled = State
    If CurrentSVG > -1 Then mnusvg(CurrentSVG).Checked = False
End Sub

Private Sub mnutest_Click()
    Clipboard.Clear
    Clipboard.SetText mnutest.Tag
    'Debug.Print "COPIED: " & mnutest.Tag
End Sub

Private Sub mnutool_Click(Index As Integer)
    AddPoint -1, CSng(Index), False
    If lstmain.ListIndex = -1 Then
        lstmain.ListIndex = lstmain.ListCount - 1
    End If
End Sub

Private Sub mnuusetracer_Click()
    If Clipboard.GetFormat(vbCFBitmap) Then
        picsec.Picture = Clipboard.GetData
        IsSet = True
        DrawPoints picmain
        CurrentTracer = "clipboard"
    End If
End Sub

Private Sub picborder_KeyDown(KeyCode As Integer, Shift As Integer)
    picmain_KeyDown KeyCode, Shift
End Sub

Private Sub picborder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picmain_MouseDown Button, Shift, X - picmain.Top - picborder.Left, Y - picmain.Top - picborder.Top
End Sub

Private Sub picborder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picmain_MouseMove Button, Shift, X - picmain.Top - picborder.Left, Y - picmain.Top - picborder.Top
End Sub

Public Function Closest(X As Single, Y As Single, Optional Excluding As Long = 0) As Long
    Dim Data() As Double, tempDistance As Long, cDistance As Long, temp As Long, tempPoint As POINTSNG
    cDistance = -1
    For temp = 1 To PointsX.count
        Data = PointsX.Item(temp)
        If Data(0) > -1 And temp <> Excluding Then
            tempPoint.X = Data(0) * picmain.Width
            tempPoint.Y = Data(1) * picmain.Height
            tempDistance = Distance(X, Y, CSng(tempPoint.X), CSng(tempPoint.Y))
            If cDistance = -1 Or tempDistance < cDistance Then
                cDistance = tempDistance
                Closest = temp
            End If
        End If
    Next
End Function

Private Sub picborder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picmain_MouseUp Button, Shift, X, Y
End Sub

Private Sub picmain_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim DontDoIt As Boolean
    Select Case KeyCode
        Case 8, 46 'delete,backspace
            cmdremove_Click
    End Select
    If SelectedTool = 2 Then 'region
        If Shift = 0 Then
            Select Case KeyCode
                Case 37: shpcursor.Width = Max(1, shpcursor.Width - 1) 'left
                Case 38: shpcursor.Height = Max(1, shpcursor.Height - 1) 'up
                Case 39: shpcursor.Width = Min(picmain.Width - shpcursor.Left, shpcursor.Width + 1)  'right
                Case 40: shpcursor.Height = Min(picmain.Height - shpcursor.Top, shpcursor.Height + 1) 'down
                Case Else: DontDoIt = True
            End Select
        Else
            Select Case KeyCode
                Case 37: shpcursor.Left = Max(0, shpcursor.Left - 1) 'left
                Case 38: shpcursor.Top = Max(0, shpcursor.Top - 1) 'up
                Case 39: shpcursor.Left = Min(shpcursor.Left + 1, picmain.Width - shpcursor.Width) 'right
                Case 40: shpcursor.Top = Min(shpcursor.Top + 1, picmain.Height - shpcursor.Height) 'down
                Case Else: DontDoIt = True
            End Select
        End If
        If Not DontDoIt Then picmain_MouseUp 0, 0, 0, 0
    End If
End Sub
Private Sub lstmain_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 8, 46 'delete,backspace
            cmdremove_Click
    End Select
End Sub


Private Sub chkmain_Click()
    chkmain.Checked = Not chkmain.Checked
    DrawPoints picmain
End Sub

Public Sub AddPoint(X As Double, Y As Double, Optional doScale As Boolean = True)
    'If X >= 0 Then
        LastX = X
        LastY = Y
        If doScale Then
            X = X / picmain.Width
            Y = Y / picmain.Height
        End If
        If lstmain.ListIndex = -1 Then lstmain.ListIndex = lstmain.ListCount - 1
        If lstmain.ListIndex = lstmain.ListCount - 1 Then
            lstmain.AddItem X & "," & Y
            lstmain.ListIndex = lstmain.ListCount - 1
            PointsX.Add MakeArray(X, Y)
        Else
            lstmain.AddItem X & "," & Y, lstmain.ListIndex + 1
            lstmain.ListIndex = lstmain.ListIndex + 1
            PointsX.Add MakeArray(X, Y), , , lstmain.ListIndex
        End If
    'End If
End Sub

Public Function MakeArray(ParamArray ArgList() As Variant) As Double()
    Dim ARR() As Double, temp As Long
    ReDim ARR(UBound(ArgList) + 1)
    For temp = 0 To UBound(ArgList)
        ARR(temp) = CSng(ArgList(temp))
    Next
    MakeArray = ARR
End Function


Private Function ConvertXY(Shift As Integer, X As Single, Y As Single) As POINTSNG
    Dim tempPoint As POINTSNG, Y2 As Long
    'Y2 = picmain.Width * 0.0125
    'X = (X \ Y2) * Y2
    'If LastX = -1 Then Y2 = picmain.Width * 0.025
    'Y = (Y \ Y2) * Y2
    If mnuedittool(1).Checked Then
        If X < 0 Then X = 0
        If Y < 0 Then Y = 0
        If X > picmain.Width Then X = picmain.Width
        If Y > picmain.Height Then Y = picmain.Height
    End If
    tempPoint.X = X
    tempPoint.Y = Y
    If Shift = 1 And LastX > -1 Then
        If Abs(X - LastX) < Abs(Y - LastY) Then
            tempPoint.X = LastX
        Else
            tempPoint.Y = LastY
        End If
    End If
    ConvertXY = tempPoint
End Function









Private Sub picmain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempPoint As POINTSNG, temp As Double
    tempPoint = ConvertXY(Shift, X, Y)
    
    Line1.X1 = 0
    Line1.X2 = picborder.Width
    Line1.Y1 = Y + picmain.Top
    Line1.Y2 = Line1.Y1
    
    Line2.X1 = X + picmain.Top
    Line2.X2 = Line2.X1
    Line2.Y1 = 0
    Line2.Y2 = picborder.Height
    
    If isdown Then
        'picmain_MouseDown Button, Shift, X, Y
        Select Case SelectedTool
            Case 1 'select closest point
                X = tempPoint.X / picmain.Width
                Y = tempPoint.Y / picmain.Height
                SetCollection PointsX, lstmain.ListIndex + 1, MakeArray(X, Y)
                lstmain.List(lstmain.ListIndex) = X & "," & Y
            Case 2 'region
                If X < 0 Then X = 0
                If X > picmain.Width Then X = picmain.Width
                If X < DownX Then
                    shpcursor.Left = X
                    shpcursor.Width = DownX - X
                Else
                    shpcursor.Left = DownX
                    shpcursor.Width = X - DownX
                End If
                If Y < 0 Then Y = 0
                If Y > picmain.Height Then Y = picmain.Height
                If Y < DownY Then
                    shpcursor.Top = Y
                    shpcursor.Height = DownY - Y
                Else
                    shpcursor.Top = DownY
                    shpcursor.Height = Y - DownY
                End If
        End Select
        DrawPoints picmain
    End If
End Sub

Private Sub picmain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tempPoint As POINTSNG
    If Button = vbLeftButton Then
        tempPoint = ConvertXY(Shift, X, Y)
        Select Case SelectedTool
            Case 0 'add point
                AddPoint tempPoint.X, tempPoint.Y
            Case 1 'select closest point
                lstmain.ListIndex = Closest(X, Y) - 1
                isdown = True
            Case 2
                isdown = True
                shpcursor.Move X, Y, 5, 5
                DownX = X
                DownY = Y
        End Select
        DrawPoints picmain
    ElseIf Button = vbRightButton Then
        Dim CenterX As Single, CenterY As Single, Angle As Long, TheDistance As Long
        CenterX = picmain.Width * 0.5
        CenterY = picmain.Height * 0.5
        Angle = ToRealDegrees(GetAngle(CenterX, CenterY, X, Y))
        TheDistance = Distance(CenterX, CenterY, X, Y)
        
        CopyRightMenu "X=" & (X / picmain.Width) & ", Y=" & (Y / picmain.Height) & " (" & Angle & "°, dX=" & (TheDistance / picmain.Width * 2) & ", dY=" & (TheDistance / picmain.Height * 2) & ")"
    End If
    'Debug.Print Button
End Sub

Public Sub CopyRightMenu(Text As String)
    mnutest.Tag = Text
    RightAlignMenu Me.hwnd, mnutestIndex, mnutest.Tag
    mnutest_Click
End Sub

Private Sub picmain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isdown = False
    If SelectedTool = 2 Then
        CopyRightMenu (shpcursor.Left / picmain.Width) & ", " & (shpcursor.Top / picmain.Height) & ", " & (shpcursor.Width / picmain.Width) & ", " & (shpcursor.Height / picmain.Height)
    End If
End Sub

Public Sub DrawGrid(Dest As PictureBox, Xs As Long, Ys As Long, Color As OLE_COLOR, Thickness As Long)
    Dim temp As Long
    For temp = Xs To Dest.Width Step Xs
        DrawLine Dest, temp, 0, temp, Dest.Height, vbRed, 1
    Next
    For temp = Ys To Dest.Height Step Ys
        DrawLine Dest, 0, temp, Dest.Width, temp, vbRed, 1
    Next
End Sub
Private Function makepointdegree(X As Long, Y As Long, Distance As Long, Degree As Long) As POINTSNG
    makepointdegree = makepoint(DegreesToXY(X, Y, Degree, Distance, Distance, True), DegreesToXY(X, Y, Degree, Distance, Distance, False))
End Function
Private Function makepoint(X As Long, Y As Long) As POINTSNG
    Dim temp As POINTSNG
    temp.X = X
    temp.Y = Y
    makepoint = temp
End Function
Private Sub DrawPoint2(Dest As PictureBox, XY As POINTSNG, Color As OLE_COLOR, Optional Number As Long = -1)
    DrawPoint Dest, XY.X, XY.Y, Color, Number
End Sub
Public Sub DrawPoints(Dest As PictureBox, Optional X As Single = -1, Optional Y As Single = -1, Optional Angle As Long = -1)
    Dim cx As Double, cy As Double, temp As Long, Data() As Double, PreviousPoint As Long, NextPoint As Long, PrevData() As Double, ResetPoint As Boolean, Stroke As Long
    Dim CurrentPoint As POINTSNG, disabled As Long, HasBroke As Boolean, Radius As Long, Multiplier As Double
    Radius = Dest.Width * 0.5
    Stroke = 1
    CurrentAngle = Angle
    Dest.Cls
    CurrentMode = 0
            
    If Angle = -1 Then
        If IsSet And chkmain.Checked Then Dest.PaintPicture picsec, 0, 0, Dest.Width, Dest.Height
        If DoGrid Then
            temp = Dest.Width * 0.05
            DrawGrid Dest, temp, Dest.Height * 0.05, vbRed, 1
        End If
        Points(0) = makepoint(0, 0)
        Points(1) = makepoint(Dest.Width, 0)
        Points(2) = makepoint(0, Dest.Height)
        Points(3) = makepoint(Dest.Width, Dest.Height)
        Multiplier = 1.41
    Else
        Multiplier = 1.2
        Points(0) = makepointdegree(Radius, Radius, Radius, Angle - 45)
        Points(1) = makepointdegree(Radius, Radius, Radius, Angle + 45)
        Points(2) = makepointdegree(Radius, Radius, Radius, Angle - 135)
        Points(3) = makepointdegree(Radius, Radius, Radius, Angle + 135)
        
        DrawPoint Dest, Points(0).X, Points(0).Y, vbRed
        DrawPoint Dest, Points(1).X, Points(1).Y, vbRed
        DrawPoint Dest, Points(2).X, Points(2).Y, vbRed
        DrawPoint Dest, Points(3).X, Points(3).Y, vbRed
    End If
    
    FirstPoint = 0
    If PointsX.count > 1 Then
        PreviousPoint = 1
        Initialize
        For temp = 1 To PointsX.count
           Data = PointsX.Item(temp)
           If Data(0) < 0 Then
                Select Case Data(0)
                    Case -1
                        CurrentMode = Data(1)
                        Select Case CurrentMode
                            Case 2, 5 'break
                                If CurrentMode = 2 Then
                                    HasBroke = True
                                End If
                                DrawPath Dest, vbBlue, Stroke, Not HasBroke
                                If CurrentMode = 5 Then HasBroke = False
                                Initialize
                                CurrentMode = 0 'line
                                FirstPoint = temp + 1
                                PreviousPoint = FirstPoint
                        End Select
                    Case -5: Stroke = Data(1)
                End Select
                PrevData = Data
           Else
                If FirstPoint = 0 Then FirstPoint = temp
                CurrentPoint = CalculatePointXY(Data(0), Data(1), Dest)
                If Angle = -1 Then
                    DrawPoint Dest, CSng(CurrentPoint.X), CSng(CurrentPoint.Y), vbBlue, temp, disabled = 0
                    If lstmain.ListIndex = temp - 1 Then
                        DrawLine Dest, 0, CurrentPoint.Y, Dest.Width, CurrentPoint.Y, vbGreen, 1
                        DrawLine Dest, CurrentPoint.X, 0, CurrentPoint.X, Dest.Height, vbGreen, 1
                    End If
                End If
                If disabled = 0 Then
                    Select Case CurrentMode
                        Case 0 'line
                            AddPathPoint CurrentPoint
                        Case 1 'curve
                            NextPoint = 1
                            If PointsX.count >= temp + 1 Then NextPoint = temp + 1
                            Draw3PointCurve Dest, 10, getPoint(PreviousPoint, Dest), CurrentPoint, getPoint(NextPoint, Dest)
                            disabled = 1
                            
                        
                        Case 3 'text
                            DrawText Dest, temp, PrevData
                        Case 4 'oval
                            DrawOval Dest, temp, CurrentMode, Radius, Angle, Multiplier
                            disabled = 2
                            
                        Case 5 'break
                            
                        
                        Case 6 'bicubic
                            disabled = getBicubic(Dest, temp)
                    End Select
                    PreviousPoint = temp
                Else
                    disabled = disabled - 1
                End If
           End If
        Next
        DrawPath Dest, vbBlue, Stroke, Not HasBroke
        HasBroke = False
    End If
    If Angle = -1 Then
        'If X > -1 And Y > -1 And cX > -1 And PointsX.Count > 0 Then
        '    ConnectLines Dest, PointsX.Count, -1, X, Y
        'Else
        '    DrawPoint Dest, X, Y, vbRed
        'End If
        hscmain_Change
    End If
    'DoEvents
End Sub

Public Function doesswitchmodes(Index As Long, Points As Long) As Boolean
    Dim temp As Long, Data() As Double
    For temp = Index To Index + Points - 1
        Data = PointsX.Item(Index)
        If Data(0) < 0 Then
            doesswitchmodes = True
            Exit For
        End If
    Next
End Function
Private Function getBicubic(Dest As Object, Index As Long) As Long
    Dim bPoints() As POINTSNG, temp As Long, Data() As Double
    getBicubic = 2
    If CanGrabPoints(Index, 4) Then
          bPoints = GrabPointsQualified(Index, 4, Dest) ' GrabPoints(dest, Index, 4)
          Bicubic Dest, CLng(bPoints(0).X), CLng(bPoints(0).Y), CLng(bPoints(1).X), CLng(bPoints(1).Y), CLng(bPoints(3).X), CLng(bPoints(3).Y), CLng(bPoints(2).X), CLng(bPoints(2).Y), 10
    End If
    If PointsX.count >= Index + 4 Then
        Data = PointsX.Item(Index + 4)
        If Data(0) < 0 Then getBicubic = 3
    End If
End Function

Public Function CanGrabPoints(Index As Long, Points As Long) As Boolean
    Dim Data() As Double
    Do Until Points = 0 Or Index >= PointsX.count
        Data = PointsX.Item(Index)
        If Data(0) >= 0 Then Points = Points - 1
    Loop
    CanGrabPoints = Points = 0
End Function
Private Function GrabPoints(Dest As Object, Index As Long, Points As Long) As POINTSNG()
    Dim ovalPoints() As POINTSNG, temp As Long, Data() As Double
    ReDim ovalPoints(Points) As POINTSNG
    Do Until Points = 0
        Data = PointsX.Item(Index + temp)
        If Data(0) >= 0 Then
            ovalPoints(temp) = CalculatePointXY(Data(0), Data(1), Dest)
            Points = Points - 1
        End If
    Loop
    GrabPoints = ovalPoints
End Function

Private Function GrabAPoint(Index As Long, Radius As Long, Optional Multiplier As Double = 1.4) As POINTSNG
    Dim Data() As Double, RET As POINTSNG
    If Index <= PointsX.count Then
        Data = PointsX.Item(Index)
        RET.X = Data(0) * Radius * Multiplier
        RET.Y = Data(1) * Radius * Multiplier
    Else
        RET.X = 0
        RET.Y = 0
    End If
    GrabAPoint = RET
End Function

Public Sub DrawOval(Dest As Object, Index As Long, Optional Shape As Long, Optional Radius As Long, Optional Angle As Long, Optional Multiplier As Double = 1.41)
    Dim ovalPoints() As POINTSNG, temp As Long, Points(3) As POINTSNG, Width As Long, Height As Long  ', XY As POINTSNG
    If CanGrabPoints(Index, 3) Then
        ovalPoints = GrabPointsQualified(Index, 3, Dest)
        DrawPoint Dest, ovalPoints(0).X, ovalPoints(0).Y, vbRed 'CENTER X/Y
        DrawPoint Dest, ovalPoints(1).X, ovalPoints(1).Y, vbRed
        DrawPoint Dest, ovalPoints(2).X, ovalPoints(2).Y, vbRed
        
        For temp = 0 To 2
            Points(temp) = GrabAPoint(Index + temp, Radius, Multiplier)
        Next
        
        Width = Max(Distance2(Points(0).X, Points(1).X), Distance2(Points(0).X, Points(2).X)) * Multiplier
        Height = Max(Distance2(Points(0).Y, Points(1).Y), Distance2(Points(0).Y, Points(2).Y)) * Multiplier
        
        Select Case Shape
            Case 4 'oval
                DrawEllipse Dest, vbBlack, ovalPoints(0).X, ovalPoints(0).Y, Width, Height, Angle, 1
        End Select
    End If
End Sub


Private Function Distance2(Start As Double, Point As Double) As Long
    If Point > Start Then
        Distance2 = Point - Start
    Else
        Distance2 = Start - Point
    End If
End Function

Private Sub DrawText(Dest As Object, Index As Long, Data() As Double)
    Dim XY As POINTSNG
    XY = getPoint(Index, Dest)
    PrintRotatedText Dest, CLng(XY.X), CLng(XY.Y), 360 - CurrentAngle, Chr(Data(2))
End Sub
Private Function CalculatePointXY(PercentX As Double, PercentY As Double, Dest As PictureBox) As POINTSNG
    Dim Point1 As POINTSNG, Point2 As POINTSNG
    Point1 = CalculatePointX(Points(0), Points(1), PercentX)
    Point2 = CalculatePointX(Points(2), Points(3), PercentX)
    'DrawLine Dest, Point1.X, Point1.Y, Point2.X, Point2.Y, vbRed
    CalculatePointXY = CalculatePointX(Point1, Point2, PercentY)
End Function
Private Function CalculatePointX(Point1 As POINTSNG, Point2 As POINTSNG, Percent As Double) As POINTSNG
    CalculatePointX = makepoint(CalculatePoint(CLng(Point1.X), CLng(Point2.X), Percent), CalculatePoint(CLng(Point1.Y), CLng(Point2.Y), Percent))
End Function

Private Function getPoint(Index As Long, Dest As PictureBox, Optional allowfirst = True) As POINTSNG
    Dim Data() As Double
    If Index <= PointsX.count Then
        Data = PointsX.Item(Index)
        If Data(0) < 0 Then
            If allowfirst And FirstPoint <> Index Then
                getPoint = getPoint(FirstPoint, Dest)
            Else
                getPoint = makepoint(-1, -1)
            End If
        Else
            getPoint = CalculatePointXY(Data(0), Data(1), Dest)
        End If
    End If
End Function

Private Function GrabPointsQualified(Index As Long, Points As Long, Dest As Object) As POINTSNG()
    Dim ovalPoints() As POINTSNG, temp As Long, Data() As Double 'FirstPoint
    ReDim ovalPoints(Points) As POINTSNG
    
    'Do Until temp = Points
    '    Data = PointsX.Item(index)
    '    If Data(0) > -1 Then
    '        ovalPoints(temp) = CalculatePointXY(Data(0), Data(1), dest)
    '        temp = temp + 1
    '    ElseIf Data(0) = -1 And Data(1) = 5 Then
    '        index = FirstPoint - 1
    '    End If
    '    index = index + 1
    '    If index >= PointsX.Count Then index = FirstPoint
    'Loop
    'ovalPoints(temp).X = index
    
    For temp = 0 To Points - 1
        ovalPoints(temp) = getPoint(Index + temp, Dest)
    Next
    GrabPointsQualified = ovalPoints
End Function












Public Function ConnectLines(Dest As PictureBox, Point1 As Long, Point2 As Long, Optional X2 As Double, Optional Y2 As Double, Optional Color As OLE_COLOR = vbBlue) As Long
    On Error Resume Next
    Dim X1 As Long, Y1 As Long, Thickness As Long, pt As POINTSNG, Point3 As Long, Data() As Double, Data2() As Double, tempstr As String
    Data = PointsX.Item(Point1)
    
    Thickness = 1
    pt = CalculatePointXY(Data(0), Data(1), Dest)
    X1 = pt.X 'PointsX.Item(Point1) * Dest.Width
    Y1 = pt.Y 'PointsY.Item(Point1) * Dest.Height
    DrawPoint Dest, CSng(X1), CSng(Y1), vbBlue, Point1
    
    If Point2 > -1 Then
        Data2 = PointsX.Item(Point2)
        pt = CalculatePointXY(Data2(0), Data2(1), Dest)
        X2 = pt.X 'PointsX.Item(Point2) * Dest.Width
        Y2 = pt.Y 'PointsY.Item(Point2) * Dest.Height
        Thickness = 2
    End If
    Select Case CurrentMode
        Case 0 'line
            DrawLine Dest, X1, Y1, X2, Y2, Color, 2 'lines
        Case 1 'curve
            If Point1 > -1 And Point2 > -1 And FirstPoint > -1 Then
                Point3 = Point1 + 1
                If Point3 >= lstmain.ListCount Then Point3 = FirstPoint
                'Debug.Print Point1 & " - " & Point2 & " - " & Point3
                Data = PointsX.Item(Point3)
                pt = CalculatePointXY(Data(0), Data(1), Dest)
                Draw3PointCurve Dest, 10, makepoint(CSng(X2), CSng(Y2)), makepoint(X1, Y1), pt
                ConnectLines = 2
            End If
        Case 3 'letter
            tempstr = Chr(Data2(2))
            PrintRotatedText Dest, CLng(X1), CLng(Y1), 360 - CurrentAngle, tempstr
    End Select
End Function

Private Sub TimerTop_Timer()
    If CurrentTop <> lstmain.TopIndex Then RefreshLineNumbers
End Sub

Sub RefreshLineNumbers()
    Dim temp As Long, count As Long
    count = lstmain.Height / GetListItemHeight(lstmain)
    CurrentTop = lstmain.TopIndex
    lstlinenumbers.Clear
    For temp = 0 To count - 1
        lstlinenumbers.AddItem temp + CurrentTop + 1
    Next
End Sub

Public Function Min(Val1 As Double, Val2 As Double) As Double
    Min = IIf(Val1 < Val2, Val1, Val2)
End Function
Public Function Max(Val1 As Double, Val2 As Double) As Double
    Max = IIf(Val1 > Val2, Val1, Val2)
End Function




