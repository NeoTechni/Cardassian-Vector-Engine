Attribute VB_Name = "Trigonometry"
Option Explicit
Public Const PI = 3.14159265358979
Public Const convert As Double = PI / 180
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Byte

Public Function DegreesToXY(CenterX As Long, CenterY As Long, Degree As Long, radiusX As Long, radiusY As Long, Optional isX As Boolean = True) As Long
    Degree = CorrectAngle(Degree)
    If isX Then DegreesToXY = CenterX - (Sin(-Degree * convert) * radiusX) Else DegreesToXY = CenterY - (Sin((90 + (Degree)) * convert) * radiusY)
End Function

Public Function CorrectAngle(Degree As Long) As Long
    If Degree < 0 Then Degree = Degree + 360
    If Degree > 360 Then Degree = Degree Mod 360
    CorrectAngle = Degree
End Function

'Trigonometic functions from iPod
Public Function Distance(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
    On Error Resume Next
    If Y2 - Y1 = 0 Then Distance = Abs(X2 - X1): Exit Function
    If X2 - X1 = 0 Then Distance = Abs(Y2 - Y1): Exit Function
    Distance = Abs(Y2 - Y1) / Sin(Atn(Abs(Y2 - Y1) / Abs(X2 - X1)))
End Function


' Draw an ellipse centered at (cx, cy) with dimensions, wid and hgt rotated angle degrees.
Public Sub DrawEllipse(Dest As Object, Color As OLE_COLOR, ByVal cx As Single, ByVal cy As Single, ByVal Width As Single, ByVal Height As Single, ByVal Angle As Single, Optional Thickness As Long = 1, Optional Sections As Long = 50)
    Dim sin_angle As Single, cos_angle As Single, theta As Single, dtheta As Single, X As Single, Y As Single, RX As Single, RY As Single, CurrentX As Single, CurrentY As Single, FirstX As Single, FirstY As Single
    If Angle > 0 Then Angle = 360 - Angle
    Angle = Angle * PI / 180
    sin_angle = Sin(Angle)
    cos_angle = Cos(Angle)
    dtheta = 2 * PI / Sections
    Do While theta < 2 * PI
        X = Width * Cos(theta)
        Y = Height * Sin(theta)
        RX = cx + X * cos_angle + Y * sin_angle
        RY = cy - X * sin_angle + Y * cos_angle
        If theta = 0 Then
            FirstX = RX
            FirstY = RY
        Else
            DrawLine Dest, CurrentX, CurrentY, RX, RY, Color, Thickness
        End If
        CurrentX = RX
        CurrentY = RY
        theta = theta + dtheta
    Loop
    DrawLine Dest, CurrentX, CurrentY, FirstX, FirstY, Color, Thickness
End Sub


Public Function Min(V1 As Long, V2 As Long) As Long
    If V1 < V2 Then Min = V1 Else Min = V2
End Function
Public Function Max(V1 As Long, V2 As Long) As Long
    If V1 > V2 Then Max = V1 Else Max = V2
End Function



Public Function GetAngle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Long
    GetAngle = CorrectAngle(AngleBySection(X1, Y1, X2, Y2, RadiansToDegrees(Angle(X1, Y1, X2, Y2))) - 180)
End Function

Public Function findXY(X As Single, Y As Single, Distance As Single, AngleRADIANS As Double, Optional isX As Boolean = True) As Single
    If isX = True Then findXY = X + Sin(AngleRADIANS) * Distance Else findXY = Y + Cos(AngleRADIANS) * Distance
End Function

Public Function Angle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Double
    On Error Resume Next
    Angle = Atn((Y2 - Y1) / (X1 - X2))
End Function

Public Function RadiansToDegrees(ByVal Radians As Double) As Double 'Converts Radians to Degrees.
    RadiansToDegrees = Radians * (180 / PI)
End Function
Public Function DegreesToRadians(ByVal Degrees As Double) As Double 'Converts Degrees to Radians.
    DegreesToRadians = Degrees * (PI / 180)
End Function

Public Function AngleSection(Angle As Long) As Long
    Select Case Angle
        Case Is > 270: AngleSection = 3
        Case Is > 180: AngleSection = 2
        Case Is > 90: AngleSection = 1
    End Select
End Function

Public Function AngleBySection(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, ByVal Angle As Long) As Double
    Angle = Abs(Angle)
    If X1 < X2 Then 'the point is at the left of Center
        If Y1 = Y2 Then AngleBySection = 180
        If Y1 < Y2 Then AngleBySection = 180 - Angle
        If Y1 > Y2 Then AngleBySection = 180 + Angle
    End If
    
    If X1 > X2 Then 'the point is at the right of Center
        If Y1 > Y2 Then AngleBySection = 360 - Angle
        If Y1 < Y2 Then AngleBySection = Angle
    End If
    
    If X1 = X2 Then
        If Y1 < Y2 Then AngleBySection = 90
        If Y1 > Y2 Then AngleBySection = 270
    End If
End Function

Public Function AngleBetweenAngles(Angle1 As Single, Angle2 As Single) As Single
    Dim temp As Long
    temp = Angle2 - Angle1
    If AngleSection(Round(Angle1)) = 0 And AngleSection(Round(Angle2)) = 3 Then temp = 360 - temp
    If AngleSection(Round(Angle1)) = 3 And AngleSection(Round(Angle2)) = 0 Then temp = 360 + temp
    AngleBetweenAngles = temp
End Function

Public Function IsWithinAngle(Angle1 As Single, Angle2 As Single, Angle As Single) As Boolean
    IsWithinAngle = Abs(AngleBetweenAngles(Angle1, Angle2)) <= Abs(Angle)
End Function

Public Function RoundAngle(Angle1 As Single, Angle As Single) As Single
    Dim temp As Single, rAngle As Single, rDistance As Long, temp2 As Long
    rDistance = 360
    'temp = Angle
    'Do While temp <= 360
    For temp = 0 To 360 Step Angle
        temp2 = AngleBetweenAngles(Angle1, temp)
        If Abs(temp2) < rDistance Then
            rDistance = Abs(temp2)
            rAngle = temp
        End If
        'temp = temp + Angle
    Next
    'Loop
    RoundAngle = rAngle
End Function

Public Function GetXYIntercept(X1 As Single, Y1 As Single, Angle As Single, X3 As Single, Y3 As Single, X4 As Single, Y4 As Single, ByRef X As Single, ByRef Y As Single) As Boolean
    Dim X2 As Single, Y2 As Single
    Const Distance As Single = 100 'Example number
    X2 = findXY(X1, Y1, Distance, CDbl(Angle), True)
    Y2 = findXY(X1, Y1, Distance, CDbl(Angle), False)
    GetXYIntercept = LineLineIntercept(X1, Y1, X2, Y2, X3, Y3, X4, Y4, CLng(X), CLng(Y))
End Function
'Intersections, obtained elsewhere
Public Function LineLineIntercept(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single, X4 As Single, Y4 As Single, ByRef X As Long, ByRef Y As Long) As Boolean
    Dim a1 As Single, a2 As Single, b1 As Single, b2 As Single, c1 As Single, c2 As Single, denom As Single
    'Translated from Pascal, lost source
    a1 = Y2 - Y1
    b1 = X1 - X2
    c1 = X2 * Y1 - X1 * Y2 '  { a1*x + b1*y + c1 = 0 is line 1 }

    a2 = Y4 - Y3
    b2 = X3 - X4
    c2 = X4 * Y3 - X3 * Y4 '  { a2*x + b2*y + c2 = 0 is line 2 }

    denom = a1 * b2 - a2 * b1

    If denom <> 0 Then
        LineLineIntercept = True
        X = (b1 * c2 - b2 * c1) / denom
        Y = (a2 * c1 - a1 * c2) / denom
    End If
End Function

Public Function LineCircleIntersept(cx As Single, cy As Single, Radius As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, ByRef ix1 As Single, ByRef iy1 As Single, Optional ByRef ix2 As Single, Optional ByRef iy2 As Single, Optional OneResult As Boolean) As Integer
    Dim dx As Single, dy As Single, a As Single, b As Single, C As Single, det As Single, t As Single
    'http://www.vb-helper.com/howto_line_circle_intersections.html
    dx = X2 - X1
    dy = Y2 - Y1

    a = dx * dx + dy * dy
    b = 2 * (dx * (X1 - cx) + dy * (Y1 - cy))
    C = (X1 - cx) * (X1 - cx) + (Y1 - cy) * (Y1 - cy) - Radius * Radius

    det = b * b - 4 * a * C
    If (a <= 0.0000001) Or (det < 0) Then
        ' No real solutions.
    ElseIf det = 0 Then
        ' One solution.
        LineCircleIntersept = 1
        t = -b / (2 * a)
        ix1 = X1 + t * dx
        iy1 = Y1 + t * dy
    Else
        ' Two solutions.
        LineCircleIntersept = 2
        t = (-b + Sqr(det)) / (2 * a)
        ix1 = X1 + t * dx
        iy1 = Y1 + t * dy
        If Not OneResult Then 'Check if I only need 1 result
            t = (-b - Sqr(det)) / (2 * a)
            ix2 = X1 + t * dx
            iy2 = Y1 + t * dy
        End If
    End If
End Function



