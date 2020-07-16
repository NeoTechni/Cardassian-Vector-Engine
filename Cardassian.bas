Attribute VB_Name = "Cardassian"
'http://www.vb-helper.com/index_graphics.html#3d

Public Const LCAR_Black As Long = 0 'RGB(0, 0, 0)
Public Const LCAR_DarkOrange As Long = 27607   'RGB(215, 107, 0)
Public Const LCAR_Orange As Long = 39421 ' rgb(253,153,0)  33023 'RGB(255, 128, 0)
Public Const LCAR_LightOrange As Long = 33023 '65535 'RGB(255, 255, 0)
Public Const LCAR_Purple As Long = 16711935 'rgb(255,0,255)
Public Const LCAR_LightPurple As Long = 13408716 ' rgb(204,153,204)
Public Const LCAR_LightBlue As Long = 13408665 'rgb(153,153,204)
Public Const LCAR_Red As Long = 6710988 'rgb(204,102,102)
Public Const LCAR_Yellow As Long = 10079487 'rgb(255,204,153)
Public Const LCAR_DarkBlue As Long = 16751001 'rgb(153,153,255)
Public Const LCAR_DarkYellow As Long = 6724095 'rgb(255,153,102)
Public Const LCAR_DarkPurple As Long = 10053324 'rgb(204,102,153)
Public Const LCAR_White As Long = 16777215

'ROTATED TEXT
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFaceName As String * 33
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private hFont As Long, oldangle As Long, f As LOGFONT, FontSize As Integer, IsSetUp As Boolean, hPrevFont As Long

Public Type POINTSNG
    X As Double
    Y As Double
End Type

'http://www.functionx.com/win32/Lesson12.htm
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Path() As POINTSNG, PointCount As Long

Public Function GetColor(Index As Integer) As Long
    Select Case Index
        Case 0: GetColor = LCAR_Black
        Case 1: GetColor = LCAR_DarkOrange
        Case 2: GetColor = LCAR_Orange
        Case 3: GetColor = LCAR_LightOrange
        Case 4: GetColor = LCAR_Purple
        Case 5: GetColor = LCAR_LightPurple
        Case 6: GetColor = LCAR_LightBlue
        Case 7: GetColor = LCAR_Red
        Case 8: GetColor = LCAR_Yellow
        Case 9: GetColor = LCAR_DarkBlue
        Case 10: GetColor = LCAR_DarkYellow
        Case 11: GetColor = LCAR_DarkPurple
        Case 12: GetColor = LCAR_White
    End Select
End Function

Public Sub Initialize()
    PointCount = 0
End Sub
Public Sub AddPathPoint(XY As POINTSNG)
    PointCount = PointCount + 1
    ReDim Preserve Path(PointCount)
    Path(PointCount - 1) = XY
End Sub
Public Sub AddPathPoint2(X As Double, Y As Double)
    Dim XY As POINTSNG
    XY.X = X
    XY.Y = Y
    AddPathPoint XY
End Sub
Public Sub DrawPath(Dest As Object, Color As OLE_COLOR, Optional Stroke As Long = 1, Optional ConnectToFirst As Boolean = True)
    On Error Resume Next
    Dim temp As Long, FirstPoint As POINTSNG, CurrentPoint As POINTSNG, APIPoints() As POINTAPI
    FirstPoint = Path(0)
    If Stroke = 0 Then
        ReDim Preserve APIPoints(0 To PointCount)
        For temp = 0 To PointCount - 1
            CurrentPoint = Path(temp)
            APIPoints(temp).X = CLng(CurrentPoint.X)
            APIPoints(temp).Y = CLng(CurrentPoint.Y)
        Next
        Dest.FillStyle = vbSolid
        Dest.ForeColor = Color
        Dest.FillColor = Color
        Dest.DrawStyle = 6
        Polygon Dest.hDC, APIPoints(0), PointCount
    Else
        For temp = 1 To PointCount - 1
            CurrentPoint = Path(temp)
            DrawLine Dest, CurrentPoint.X, CurrentPoint.Y, FirstPoint.X, FirstPoint.Y, Color, Stroke
            FirstPoint = CurrentPoint
        Next
        If ConnectToFirst Then
            FirstPoint = Path(0)
            DrawLine Dest, CurrentPoint.X, CurrentPoint.Y, FirstPoint.X, FirstPoint.Y, Color, Stroke
        End If
    End If
End Sub
Public Function killallexceptnumbers(Text As String) As Double
    Dim temp As Long, tempstr As String
    For temp = 1 To Len(Text)
        Select Case Mid(Text, temp, 1)
            Case "-"
                If temp = 1 Then tempstr = "-"
            Case "."
                If InStr(tempstr, ".") = 0 And temp < Len(Text) Then tempstr = tempstr & "."
            Case "E"
                If InStr(tempstr, "E") = 0 Then tempstr = tempstr & "E-"
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
                tempstr = tempstr & Mid(Text, temp, 1)
        End Select
    Next
    If Len(tempstr) > 0 Then killallexceptnumbers = CDbl(tempstr)
End Function
Public Sub DrawPoint(Dest As PictureBox, X As Double, Y As Double, Color As OLE_COLOR, Optional Number As Long = -1, Optional CrossStyle As Boolean = True)
    Const s As Integer = 5
    If CrossStyle Then
        DrawLine Dest, X - s, Y, X + s + 1, Y, Color
        DrawLine Dest, X, Y - s, X, Y + s + 1, Color
    Else
        Dest.Circle (X, Y), s, Color
    End If
    If Number > -1 Then
        Dest.CurrentX = X
        Dest.CurrentY = Y
        Dest.Print Number
    End If
End Sub
Public Sub DrawLine(Dest As PictureBox, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Color As OLE_COLOR, Optional Thickness As Long = 1)
    If (X1 = 0 And Y1 = 0) Then 'Or (X2 = 0 And Y2 = 0) Then
        Exit Sub
    End If
    Dest.DrawWidth = Abs(Thickness)
    Dest.Line (X1, Y1)-(X2, Y2), Color
End Sub

Sub ClearCollection(LST As Collection, Optional DownTo As Long = 1)
    Dim temp As Long
    For temp = LST.count To DownTo Step -1
        LST.Remove temp
    Next
End Sub
Sub SetCollection(LST As Collection, Index As Long, Value)
    'Debug.Print "Before: " & object & " Set to:" & Value
    On Error Resume Next
    If Index = LST.count Then
        LST.Remove LST.count
        LST.Add Value
    Else
        LST.Remove Index
        LST.Add Value, , Index
    End If
    'Debug.Print "After: " & LST.Item(Index)
End Sub

Function CalculatePoint(pt1 As Long, PT2 As Long, Percent As Double) As Long
    CalculatePoint = (PT2 - pt1) * Percent + pt1
End Function

Sub DrawPercentagePoint(Dest As PictureBox, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Percent As Double, Optional Color As OLE_COLOR = vbBlue)
    DrawLine Dest, X1, Y1, X2, Y2, vbBlack, 1
    X1 = CalculatePoint(X1, X2, Percent)
    Y1 = CalculatePoint(Y1, Y2, Percent)
    DrawPoint Dest, CDbl(X1), CDbl(Y1), Color
End Sub

Sub Bicubic(Dest As PictureBox, X1 As Long, Y1 As Long, aX1 As Long, aY1 As Long, X2 As Long, Y2 As Long, aX2 As Long, aY2 As Long, Increment As Long, Optional Color As Long = -1, Optional Thickness As Long = 1)
    Dim temp As Long, Percent As Double, X3 As Long, Y3 As Long, X4 As Long, Y4 As Long, PrevPoint As POINTSNG
    For temp = 0 To 100 Step Increment
        Percent = temp * 0.01
        
        X3 = CalculatePoint(X1, aX1, Percent)
        Y3 = CalculatePoint(Y1, aY1, Percent)
        
        X4 = CalculatePoint(aX2, X2, Percent)
        Y4 = CalculatePoint(aY2, Y2, Percent)
        
        X3 = CalculatePoint(X3, X4, Percent)
        Y3 = CalculatePoint(Y3, Y4, Percent)
        
        If Color = -1 Then
            AddPathPoint2 CDbl(X3), CDbl(Y3)
        Else
            If temp > 0 Then
                DrawLine Dest, X3, Y3, PrevPoint.X, PrevPoint.Y, Color, Thickness
            End If
            PrevPoint.X = X3
            PrevPoint.Y = Y3
        End If
    Next
End Sub

' Parametric X function for drawing a degree 3 Bezier curve.
Private Function X(ByVal t As Double, ByVal x0 As Double, ByVal X1 As Double, ByVal X2 As Double, ByVal X3 As Double) As Double
    X = CDbl(x0 * (1 - t) ^ 3 + X1 * 3 * t * (1 - t) ^ 2 + X2 * 3 * t ^ 2 * (1 - t) + X3 * t ^ 3)
End Function

' Parametric Y function for drawing a degree 3 Bezier curve.
Private Function Y(ByVal t As Double, ByVal y0 As Double, ByVal Y1 As Double, ByVal Y2 As Double, ByVal Y3 As Double) As Double
    Y = CDbl(y0 * (1 - t) ^ 3 + Y1 * 3 * t * (1 - t) ^ 2 + Y2 * 3 * t ^ 2 * (1 - t) + Y3 * t ^ 3)
End Function

Public Sub Draw3PointCurve(ByVal pic As Object, Steps As Long, pt0 As POINTSNG, pt1 As POINTSNG, PT2 As POINTSNG)
    'DrawPoint pic, pt0.X, pt0.Y, vbRed
    'DrawPoint pic, pt1.X + 1, pt1.Y + 1, vbBlue
    'DrawPoint pic, pt2.X - 1, pt2.Y - 1, vbGreen
    
    DrawBezier pic, 1 / Steps, pt0, pt1, pt1, PT2
End Sub

' Draw the Bezier curve.
Public Sub DrawBezier(ByVal pic As Object, ByVal dt As Double, pt0 As POINTSNG, pt1 As POINTSNG, PT2 As POINTSNG, PT3 As POINTSNG)
    ' Debugging code.
    ' Draw the control lines.
    Dim p As PictureBox
    'pic.DrawStyle = vbDot
    'pic.Line (pt0.X, pt0.Y)-(pt1.X, pt1.Y)
    'pic.Line (pt2.X, pt2.Y)-(pt3.X, pt3.Y)
    pic.DrawStyle = vbSolid

    ' Draw the curve.
    Dim t As Double
    Dim x0 As Double
    Dim y0 As Double
    Dim X1 As Double
    Dim Y1 As Double

    t = 0#
    X1 = X(t, pt0.X, pt1.X, PT2.X, PT3.X)
    Y1 = Y(t, pt0.Y, pt1.Y, PT2.Y, PT3.Y)
    AddPathPoint2 X1, Y1
    t = t + dt
    Do While t < 1#
        x0 = X1
        y0 = Y1
        X1 = X(t, pt0.X, pt1.X, PT2.X, PT3.X)
        Y1 = Y(t, pt0.Y, pt1.Y, PT2.Y, PT3.Y)
        'pic.Line (x0, y0)-(x1, y1)
        AddPathPoint2 X1, Y1
        t = t + dt
    Loop

    ' Connect to the final point.
    t = 1#
    x0 = X1
    y0 = Y1
    X1 = X(t, pt0.X, pt1.X, PT2.X, PT3.X)
    Y1 = Y(t, pt0.Y, pt1.Y, PT2.Y, PT3.Y)
    'pic.Line (x0, y0)-(x1, y1)
    AddPathPoint2 X1, Y1
End Sub


' Draw a cardinal spline built from connected Bezier curves.
' dt = fraction of curve inpercent that each step takes up (1/steps)
Public Sub DrawCurve(ByVal pic As Object, ByVal dt As Double, ByVal tension As Double, pts() As POINTSNG)
    Dim control_scale As Double
    Dim pt As POINTSNG
    Dim pt_before As POINTSNG
    Dim pt_after As POINTSNG
    Dim pt_after2 As POINTSNG
    Dim Di As POINTSNG
    Dim DiPlus1 As POINTSNG
    Dim p1 As POINTSNG
    Dim p2 As POINTSNG
    Dim p3 As POINTSNG
    Dim p4 As POINTSNG
    Dim i As Integer

    control_scale = CDbl(tension / 0.5 * 0.175)
    For i = LBound(pts) To UBound(pts) - 1
        pt_before = pts(Max(i - 1, 0))
        pt = pts(i)
        pt_after = pts(i + 1)
        pt_after2 = pts(Min(i + 2, UBound(pts)))

        p1 = pts(i)
        p4 = pts(i + 1)

        Di.X = pt_after.X - pt_before.X
        Di.Y = pt_after.Y - pt_before.Y
        p2.X = pt.X + control_scale * Di.X
        p2.Y = pt.Y + control_scale * Di.Y

        DiPlus1.X = pt_after2.X - pts(i).X
        DiPlus1.Y = pt_after2.Y - pts(i).Y
        p3.X = pt_after.X - control_scale * DiPlus1.X
        p3.Y = pt_after.Y - control_scale * DiPlus1.Y

        DrawBezier pic, dt, p1, p2, p3, p4
    Next i
End Sub

'ROTATED TEXT
Public Sub PrintRotatedText(Dest As Object, X As Long, Y As Long, Angle As Long, Text As String)
    FontSize = Dest.Font.Size
    If Not IsSetUp Or Angle <> oldangle Then
        If IsSetUp Then DeleteObject hFont
        f.lfEscapement = 10 * Angle 'rotation angle, in tenths
        f.lfFaceName = Dest.Font.Name + Chr$(0)
        f.lfHeight = (FontSize * -20) / Screen.TwipsPerPixelY
        hFont = CreateFontIndirect(f)
    End If
    hPrevFont = SelectObject(Dest.hDC, hFont)
    Dest.CurrentX = X
    Dest.CurrentY = Y
    Dest.Print Text
    hFont = SelectObject(Dest.hDC, hPrevFont)
End Sub
Public Sub PrintText(Dest As Object, X As Long, Y As Long, Text As String)
    Dest.CurrentX = X
    Dest.CurrentY = Y
    Dest.Print Text
End Sub

Public Function toArray(Optional Text As String = "0,0|0.2983333,0.4983333|0,1|0.2483333,1|0.3,0.5|0.25,0")
    Dim LeftSide As String
    LeftSide = "array as Float("
    toArray = LeftSide & LeftSide & Replace(Text, "|", "), " & LeftSide) & "))"
End Function
