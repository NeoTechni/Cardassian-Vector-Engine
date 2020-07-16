VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

' Constants
Private Const SRCCOPY = &HCC0020
Private Const BI_RGB = 0&
Private Const CBM_INIT = &H4
Private Const DIB_RGB_COLORS = 0

' Types
Private Type RGBTriple
    Red As Byte
    Green As Byte
    Blue As Byte
End Type
Private Type BMPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO_1
    bmiHeader As BMPINFOHEADER
    bmiColors(1) As RGBQUAD
End Type
Private Type BITMAPINFO_2
    bmiHeader As BMPINFOHEADER
    bmiColors(3) As RGBQUAD
End Type
Private Type BITMAPINFO_4
    bmiHeader As BMPINFOHEADER
    bmiColors(15) As RGBQUAD
End Type
Private Type BITMAPINFO_8
    bmiHeader As BMPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type
Private Type BITMAPINFO_16
    bmiHeader As BMPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Type BITMAPINFO_24
    bmiHeader As BMPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Type BITMAPINFO_24a
    bmiHeader As BMPINFOHEADER
    bmiColors As RGBTriple
End Type

' Functions
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateDIBitmap_1 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hDC As Long, lpInfoHeader As BMPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_1, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_2 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hDC As Long, lpInfoHeader As BMPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_2, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_4 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hDC As Long, lpInfoHeader As BMPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_4, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_8 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hDC As Long, lpInfoHeader As BMPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_8, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_16 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hDC As Long, lpInfoHeader As BMPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_16, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_24 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hDC As Long, lpInfoHeader As BMPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_24, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_24a Lib "gdi32" Alias "CreateDIBitmap" (ByVal hDC As Long, lpInfoHeader As BMPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_24a, ByVal wUsage As Long) As Long

'header
Private bm1 As BITMAPINFO_1
Private bm2 As BITMAPINFO_2
Private bm4 As BITMAPINFO_4
Private bm8 As BITMAPINFO_8
Private bm16 As BITMAPINFO_16
Private bm24 As BITMAPINFO_24
Private bm24a As BITMAPINFO_24a
'bitmap handle.
Private hBmp As Long

Private Type ScTw
    Width As Long
    Height As Long
End Type



Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal length As Long, ByVal Fill As Byte)
Private Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    SizeImage As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed As Long
    ClrImportant As Long
End Type
Private RBD As Long
Private IDATData() As Byte
Dim IdataLen As Long
Private Type IHDR
    Width As Long
    Height As Long
    BitDepth As Byte
    ColorType As Byte
    Compression As Byte
    filter As Byte
    Interlacing As Byte
End Type
'For Decompression:
Private Type CodesType
    Lenght() As Long
    code() As Long
End Type
Private m_Backcolor As Long
Private Palettenbyte() As Byte
Private OutStream() As Byte
Private OutPos As Long
Private InStream() As Byte
Private Inpos As Long
Private ByteBuff As Long
Private BitNum As Long
Private BitMask(16) As Long
Private Pow2(16) As Long
Private LC As CodesType
Private dc As CodesType
Private LitLen As CodesType
Private Dist As CodesType
Private TempLit As CodesType
Private TempDist As CodesType
Private LenOrder(18) As Long
Private MinLLenght As Long
Private MaxLLenght As Long
Private MinDLenght As Long
Private MaxDLenght As Long
Private IsStaticBuild As Boolean
Private BPPprivat As Long
Private m_width As Long
Private m_height As Long
Private m_bitdepht As Long
Private m_colortype As Long
Private m_compression As Long
Private m_filter As Long
Private m_interlacing As Long
Private m_ErrorNumber As Long
Private m_sAlpha As Boolean
Private m_hAlpha As Boolean
Private trns() As Byte
Private m_hTrans As Boolean
Private m_sTrans As Boolean
Private Colorused As Long
Private bkgd() As Byte
Private m_hbkgd As Boolean
Private m_bkgdColor As Long
Private m_text As String
Private m_Time As String
Private m_ztext As String
Private m_gama As Long
Private m_Bgx As Long
Private m_Bgy As Long
Private m_BGPic As Object
Private m_OwnBkgnd As Boolean
Private m_OBCol As Long
Private m_PicBox As Object
Private m_settoBG As Boolean
Public Function OpenPNG(Filename As String) As Long
    Dim Stand As Long, Ende As Boolean, Filenumber As Long, Signature(7) As Byte, Test As Long, Länge As Long
    Dim ChunkName As String * 4, ChunkInhalt() As Byte, CRC32Inhalt As Long, Teststring As String
    Dim TestCRC32 As Long, Testint As Integer 'Dim crc32test As New clsCRC
    m_hbkgd = False
    m_hTrans = False
    BPPprivat = 0
    ReDim IDATData(0)
    IdataLen = 0
    Filenumber = FreeFile
    Open Filename For Binary As Filenumber
    Get Filenumber, , Signature
    Test = IsValidSignature(Signature)
    If Test <> -1 Then
        m_ErrorNumber = 1
        Exit Function
    End If
    Do While Ende = False
        Get Filenumber, , Länge
        SwapBytesLong Länge
        Get Filenumber, , ChunkName
        If Länge > 0 Then ReDim ChunkInhalt(Länge - 1)
        Stand = Seek(Filenumber)
        If Stand + Länge > LOF(Filenumber) Then
            m_ErrorNumber = 3
            Exit Function
        End If
        Get Filenumber, , ChunkInhalt
        Get Filenumber, , CRC32Inhalt
        'SwapBytesLong CRC32Inhalt
        'teststring = ChunkName & StrConv(ChunkInhalt, vbUnicode)
        'Testcrc32 = CRC32(teststring) 'reiner VB-Code
        'crc32test.Algorithm = 1
        'TestCRC32 = crc32test.CalculateString(teststring) 'VB und Assembler
        'If CRC32Inhalt <> 0 Then
        'If CRC32Inhalt <> TestCRC32 Then
        'MsgBox "Bad crc32"
        'm_ErrorNumber = 2
        'Exit Function
        'End If
        'End If
        Select Case ChunkName
            Case "IHDR"
                ReadIHDR ChunkInhalt
            Case "PLTE"
                ReDim Palettenbyte(UBound(ChunkInhalt))
                CopyMemory Palettenbyte(0), ChunkInhalt(0), UBound(ChunkInhalt) + 1
            Case "IDAT"
                ReDim Preserve IDATData(IdataLen + UBound(ChunkInhalt))
                CopyMemory IDATData(IdataLen), ChunkInhalt(0), UBound(ChunkInhalt) + 1
                IdataLen = UBound(IDATData) + 1
            Case "IEND"
                Ende = True
            Case "bKGD"
                bkgd = ChunkInhalt
                ReadBkgd
                m_hbkgd = True
            Case "cHRM"
            Case "oFFs"
            Case "pCaL"
            Case "sCAL"
            Case "gAMA"
                CopyMemory ByVal VarPtr(m_gama), ChunkInhalt(0), 4
                SwapBytesLong m_gama
            Case "hIST"
            Case "pHYs"
            Case "sBIT"
            Case "tEXt"
                m_text = m_text & StrConv(ChunkInhalt, vbUnicode) & Chr(0)
            Case "zTXt"
                DecompressText ChunkInhalt
            Case "gIFg"
            Case "gIFx"
            Case "tIME"
                CopyMemory ByVal VarPtr(Testint), ChunkInhalt(0), 2
                Swap Testint
                m_Time = Format(ChunkInhalt(3), "00") & "." & Format(ChunkInhalt(2), "00") & "." & Testint & " " & Format(ChunkInhalt(4), "00") & ":" & Format(ChunkInhalt(5), "00") & ":" & Format(ChunkInhalt(6), "00")
            Case "tRNS"
                m_hTrans = True
                trns = ChunkInhalt
            Case "cTXt"
            Case Else
            'If Asc(Left(ChunkName, 1)) > 65 Then Exit Function 'kritischer Chunk
        End Select
    Loop
    If IdataLen = 0 Then
        m_ErrorNumber = 4
        Exit Function
    End If
    Close Filenumber
    MakePicture
End Function
Private Function IsValidSignature(Signature() As Byte) As Boolean
    If Signature(0) <> 137 Then Exit Function
    If Signature(1) <> 80 Then Exit Function
    If Signature(2) <> 78 Then Exit Function
    If Signature(3) <> 71 Then Exit Function
    If Signature(4) <> 13 Then Exit Function
    If Signature(5) <> 10 Then Exit Function
    If Signature(6) <> 26 Then Exit Function
    If Signature(7) <> 10 Then Exit Function
    IsValidSignature = True
End Function
Private Sub SwapBytesLong(ByteValue As Long)
    Dim Übergabe As Long, i As Long
    For i = 0 To 3
        CopyMemory ByVal VarPtr(Übergabe) + i, ByVal VarPtr(ByteValue) + (3 - i), 1
    Next i
    ByteValue = Übergabe
End Sub
Private Sub ReadIHDR(Bytefeld() As Byte)
    Dim Header As IHDR
    CopyMemory ByVal VarPtr(Header), Bytefeld(0), 13
    SwapBytesLong Header.Width
    SwapBytesLong Header.Height
    m_width = Header.Width
    m_height = Header.Height
    m_bitdepht = Header.BitDepth
    m_colortype = Header.ColorType
    m_compression = Header.Compression
    m_filter = Header.filter
    m_interlacing = Header.Interlacing
End Sub
Public Property Get Width() As Long
    Width = m_width
End Property
Public Property Get Height() As Long
    Height = m_height
End Property
Public Property Get Bitdepht() As Long
    Bitdepht = m_bitdepht
End Property
Public Property Get ColorType() As Long
    ColorType = m_colortype
End Property
Public Property Get Compression() As Long
    Compression = m_compression
End Property
Public Property Get filter() As Long
    filter = m_filter
End Property
Public Property Get Interlacing() As Long
    Interlacing = m_interlacing
End Property
Private Sub MakePicture()
    Dim DataSize As Long
    Dim Buffer() As Byte
    Dim BitCount As Integer
    Dim Bitdepht As Long
    Dim Drehen As Integer
    m_hAlpha = False
    Drehen = 1
    Select Case Me.Interlacing
        Case 0
        DataSize = DataPerRow * Me.Height
        Case 1
        DataSize = (DataPerRow * Me.Height) + Me.Height
        End Select
        ReDim Buffer(UBound(IDATData) - 2)
        CopyMemory Buffer(0), IDATData(2), UBound(IDATData) - 1
        Select Case Me.Compression
        Case 0
        Decompress Buffer, DataSize
        End Select
        Select Case Me.Interlacing
        Case 0
        Buffer = DeFilter(Buffer)
        Drehen = 1
        Case 1
        Buffer = DeFilterInterlaced(Buffer)
        Drehen = 0
        End Select
        BitCount = Me.Bitdepht
        Select Case Me.ColorType
        Case 0 'Grayscale
        Select Case Me.Bitdepht
        Case 16
        Conv16To8 Buffer
        InitColorTable_Grey 8
        BitCount = 8
        BPPprivat = 8
        Case 8, 4, 1
        Select Case Interlacing
        Case 0
        BitCount = Me.Bitdepht
        InitColorTable_Grey Me.Bitdepht, False
        Align32 BitCount, Buffer
        Case Else
        BitCount = 8
        InitColorTable_Grey Me.Bitdepht, True
        End Select
        Case 2
        InitColorTable_Grey 2
        If Me.Interlacing = 0 Then
        Pal2To8 Me.Width, Me.Height, Buffer, DataPerRow
        End If
        BitCount = 8
        BPPprivat = 8
        End Select
        If m_hTrans And m_sTrans Then
        If Me.Bitdepht <> 2 Then
        Align32 BitCount, Buffer
        End If
        PalToRGBA Me.Width, Me.Height, BitCount, Buffer
        BitCount = 32
        BPPprivat = 32
        MakeAlpha m_BGPic, Buffer, m_Bgx, m_Bgy
        BitCount = 24
        BPPprivat = 24
        End If
        Case 2 'RGB
        If Me.Bitdepht = 16 Then Conv16To8 Buffer
        BitCount = 24
        BPPprivat = 24
        ReverseRGB Buffer
        Drehen = 1
        BPPprivat = 8
        Align32 BitCount, Buffer
        BPPprivat = 24
        If m_hTrans And m_sTrans Then
        MakeRGBTransparent Buffer
        MirrorData Buffer, Me.Width * 4
        Drehen = 0
        BitCount = 32
        BPPprivat = 32
        MakeAlpha m_BGPic, Buffer, m_Bgx, m_Bgy
        BitCount = 24
        BPPprivat = 24
        End If
        Case 3 'Palette
        Select Case Me.Bitdepht
        Case 8, 4, 1
        If Me.Interlacing = 1 Then
        BitCount = 8
        BPPprivat = 8
        Align32 BitCount, Buffer
        Else
        BitCount = Me.Bitdepht
        If BitCount >= 8 Then
        Align32 BitCount, Buffer
        End If
        End If
        Case 2
        If Me.Interlacing = 0 Then
        Pal2To8 Me.Width, Me.Height, Buffer, DataPerRow
        BitCount = 8
        BPPprivat = 8
        Else
        BitCount = 8
        BPPprivat = 8
        Align32 BitCount, Buffer
        End If
        End Select
        If m_hTrans And m_sTrans Then
        If Me.Bitdepht <> 2 Then
        Align32 BitCount, Buffer
        End If
        PalToRGBA Me.Width, Me.Height, BitCount, Buffer
        BitCount = 32
        BPPprivat = 32
        MakeAlpha m_BGPic, Buffer, m_Bgx, m_Bgy
        BitCount = 24
        BPPprivat = 24
        End If
        Case 4 'Grayscale + Alpha
        m_hAlpha = True
        If Me.Bitdepht = 16 Then Conv16To8 Buffer
        GrayAToRGBA Buffer
        BPPprivat = 32
        BitCount = 32
        MirrorData Buffer, LineBytes(Me.Width, BitCount)
        Drehen = 0
        If m_sAlpha = True Then
        MakeAlpha m_BGPic, Buffer, m_Bgx, m_Bgy
        BPPprivat = 24
        BitCount = 24
        End If
        Case 6 'RGB + Alpha
        m_hAlpha = True
        If Me.Bitdepht = 16 Then Conv16To8 Buffer
        BitCount = 32
        BPPprivat = 32
        ReverseRGBA Buffer
        MirrorData Buffer, LineBytes(Me.Width, BitCount)
        Drehen = 0
        If m_sAlpha = True Then
        MakeAlpha m_BGPic, Buffer, m_Bgx, m_Bgy
        BPPprivat = 24
        BitCount = 24
        End If
        End Select
        If Not (((Me.ColorType = 3) And (BitCount = 32)) Or _
        (Me.Bitdepht = 2)) Then
        Select Case Me.Bitdepht
        Case 16
        Bitdepht = 8
        Bitdepht = 16
        End Select
        End If
        Select Case BitCount
        Case 1, 2, 4
        Align32 BitCount, Buffer
        End Select
        Select Case BitCount
        Case 1
        Select Case Me.ColorType
        Case 3
        InitColorTable_1Palette Palettenbyte
        Case Else
        InitColorTable_1
        End Select
        CreateBitmap_1 Buffer, Me.Width, Me.Height, True, Colorused
        DrawBitmap Me.Width, Me.Height, m_PicBox, True, m_Bgx, m_Bgy, m_settoBG
        Case 4
            Select Case Me.ColorType
        Case 0
        Case Else
     InitColorTable_4 Palettenbyte
    End Select
        CreateBitmap_4 Buffer, Me.Width, Me.Height, True, Colorused
        DrawBitmap Me.Width, Me.Height, m_PicBox, True, m_Bgx, m_Bgy, m_settoBG
    Case 8
        Select Case Me.ColorType
            Case 0, 4
            Case Else
                InitColorTable_8 Palettenbyte
        End Select
        Drehen = 1
        CreateBitmap_8 Buffer, Me.Width, Me.Height, Drehen, Colorused
        DrawBitmap Me.Width, Me.Height, m_PicBox, True, m_Bgx, m_Bgy, m_settoBG
    Case 24
        CreateBitmap_24 Buffer, Me.Width, Me.Height, Drehen, 1
        DrawBitmap Me.Width, Me.Height, m_PicBox, True, m_Bgx, m_Bgy, m_settoBG
    Case 32
        CreateBitmap_24 Buffer, Me.Width, Me.Height, Drehen
        DrawBitmap Me.Width, Me.Height, m_PicBox, True, m_Bgx, m_Bgy, m_settoBG
    End Select
End Sub
Private Function Decompress(bytearray() As Byte, UncompressedSize As Long, Optional ZIP64 As Boolean = False) As Long
'On Error Resume Next

Dim IsLastBlock As Boolean
Dim CompType As Long
Dim char As Long
Dim Nubits As Long
Dim L1 As Long
Dim L2 As Long
Dim X As Long
UncompressedSize = UncompressedSize + 100
InStream = bytearray
Call Init_Decompress(UncompressedSize)
Do
 IsLastBlock = GetBits(1)
 CompType = GetBits(2)
If CompType = 0 Then
If Inpos + 4 > UBound(InStream) Then
 Decompress = -1
 Exit Do
End If
Do While BitNum >= 8
 Inpos = Inpos - 1
 BitNum = BitNum - 8
Loop
 CopyMemory L1, InStream(Inpos), 2&
 CopyMemory L2, InStream(Inpos + 2), 2&
 Inpos = Inpos + 4
If L1 - (Not (L2) And &HFFFF&) Then Decompress = -2
If Inpos + L1 - 1 > UBound(InStream) Then
 Decompress = -1
 Exit Do
End If
If OutPos + L1 - 1 > UBound(OutStream) Then
 Decompress = -1
 Exit Do
End If
 CopyMemory OutStream(OutPos), InStream(Inpos), L1
 OutPos = OutPos + L1
 Inpos = Inpos + L1
 ByteBuff = 0
 BitNum = 0
ElseIf CompType = 3 Then
 Decompress = -1
 Exit Do
Else
If CompType = 1 Then
If Create_Static_Tree <> 0 Then
 MsgBox "Error in tree creation (Static)"
 Exit Function
End If
Else
If Create_Dynamic_Tree <> 0 Then
 MsgBox "Error in tree creation (Static)"
 Exit Function
End If
End If
 Do
 NeedBits MaxLLenght
 Nubits = MinLLenght
Do While LitLen.Lenght(ByteBuff And BitMask(Nubits)) <> Nubits
 Nubits = Nubits + 1
Loop
 char = LitLen.code(ByteBuff And BitMask(Nubits))
 DropBits Nubits
If char < 256 Then
If OutPos > UBound(OutStream) Then Decompress = -1: Exit Function
OutStream(OutPos) = char
 OutPos = OutPos + 1
ElseIf char > 256 Then
 char = char - 257
 L1 = LC.code(char) + GetBits(LC.Lenght(char))
If (L1 = 258) And ZIP64 Then L1 = GetBits(16) + 3
 NeedBits MaxDLenght
 Nubits = MinDLenght
Do While Dist.Lenght(ByteBuff And BitMask(Nubits)) <> Nubits
 Nubits = Nubits + 1
Loop
 char = Dist.code(ByteBuff And BitMask(Nubits))
 DropBits Nubits
 L2 = dc.code(char) + GetBits(dc.Lenght(char))
For X = 1 To L1
If OutPos > UncompressedSize Then
 OutPos = UncompressedSize
 GoTo Stop_Decompression
End If
 OutStream(OutPos) = OutStream(OutPos - L2)
 OutPos = OutPos + 1
Next X
End If
Loop While char <> 256 'EOB
End If
Loop While Not IsLastBlock
Stop_Decompression:
If OutPos > 0 Then
 ReDim Preserve OutStream(OutPos - 1)
Else
 Erase OutStream
End If
Erase InStream
Erase BitMask
Erase Pow2
Erase LC.code
Erase LC.Lenght
Erase dc.code
Erase dc.Lenght
Erase LitLen.code
Erase LitLen.Lenght
Erase Dist.code
Erase Dist.Lenght
Erase LenOrder
bytearray = OutStream
End Function
Private Function Create_Static_Tree()
Dim X As Long
Dim Lenght(287) As Long
If IsStaticBuild = False Then
For X = 0 To 143: Lenght(X) = 8: Next
For X = 144 To 255: Lenght(X) = 9: Next
For X = 256 To 279: Lenght(X) = 7: Next
For X = 280 To 287: Lenght(X) = 8: Next
If Create_Codes(TempLit, Lenght, 287, MaxLLenght, MinLLenght) <> 0 Then
 Create_Static_Tree = -1
 Exit Function
End If
For X = 0 To 31: Lenght(X) = 5: Next
 Create_Static_Tree = Create_Codes(TempDist, Lenght, 31, MaxDLenght, MinDLenght)
 IsStaticBuild = True
Else
 MinLLenght = 7
 MaxLLenght = 9
 MinDLenght = 5
 MaxDLenght = 5
End If
LitLen = TempLit
Dist = TempDist
End Function
Private Function Create_Dynamic_Tree() As Long
Dim Lenght() As Long
Dim Bl_Tree As CodesType
Dim MinBL As Long
Dim MaxBL As Long
Dim NumLen As Long
Dim Numdis As Long
Dim NumCod As Long
Dim char As Long
Dim Nubits As Long
Dim LN As Long
Dim Pos As Long
Dim X As Long
NumLen = GetBits(5) + 257
Numdis = GetBits(5) + 1
NumCod = GetBits(4) + 4
ReDim Lenght(18)
For X = 0 To NumCod - 1
 Lenght(LenOrder(X)) = GetBits(3)
Next
For X = NumCod To 18
 Lenght(LenOrder(X)) = 0
Next
If Create_Codes(Bl_Tree, Lenght, 18, MaxBL, MinBL) <> 0 Then
 Create_Dynamic_Tree = -1
 Exit Function
End If
ReDim Lenght(NumLen + Numdis)
Pos = 0
Do While Pos < NumLen + Numdis
 NeedBits MaxBL
 Nubits = MinBL
Do While Bl_Tree.Lenght(ByteBuff And BitMask(Nubits)) <> Nubits
 Nubits = Nubits + 1
Loop
 char = Bl_Tree.code(ByteBuff And BitMask(Nubits))
 DropBits Nubits
If char < 16 Then
 Lenght(Pos) = char
 Pos = Pos + 1
Else
If char = 16 Then
If Pos = 0 Then
 Create_Dynamic_Tree = -5
 Exit Function
End If
 LN = Lenght(Pos - 1)
 char = 3 + GetBits(2)
ElseIf char = 17 Then
 char = 3 + GetBits(3)
 LN = 0
Else
 char = 11 + GetBits(7)
 LN = 0
End If
If Pos + char > NumLen + Numdis Then
 Create_Dynamic_Tree = -6
 Exit Function
End If
Do While char > 0
 char = char - 1
 Lenght(Pos) = LN
 Pos = Pos + 1
Loop
End If
Loop
If Create_Codes(LitLen, Lenght, NumLen - 1, MaxLLenght, MinLLenght) <> 0 Then
 Create_Dynamic_Tree = -1
 Exit Function
End If
For X = 0 To Numdis
 Lenght(X) = Lenght(X + NumLen)
Next
 Create_Dynamic_Tree = Create_Codes(Dist, Lenght, Numdis - 1, MaxDLenght, MinDLenght)
End Function
Private Function Create_Codes(tree As CodesType, Lenghts() As Long, NumCodes As Long, MaxBits As Long, Minbits As Long) As Long
Dim Bits(16) As Long
Dim next_code(16) As Long
Dim code As Long
Dim LN As Long
Dim X As Long
Minbits = 16
For X = 0 To NumCodes
 Bits(Lenghts(X)) = Bits(Lenghts(X)) + 1
If Lenghts(X) > MaxBits Then MaxBits = Lenghts(X)
If Lenghts(X) < Minbits And Lenghts(X) > 0 Then Minbits = Lenghts(X)
Next
LN = 1
For X = 1 To MaxBits
 LN = LN + LN
 LN = LN - Bits(X)
If LN < 0 Then Create_Codes = LN: Exit Function
Next
Create_Codes = LN
ReDim tree.code(2 ^ MaxBits - 1)
ReDim tree.Lenght(2 ^ MaxBits - 1)
code = 0
Bits(0) = 0
For X = 1 To MaxBits
 code = (code + Bits(X - 1)) * 2
next_code(X) = code
Next
For X = 0 To NumCodes
 LN = Lenghts(X)
If LN <> 0 Then
 code = Bit_Reverse(next_code(LN), LN)
 tree.Lenght(code) = LN
 tree.code(code) = X
next_code(LN) = next_code(LN) + 1
End If
Next
End Function
Private Function Bit_Reverse(ByVal Value As Long, ByVal Numbits As Long)
    Do While Numbits > 0
        Bit_Reverse = Bit_Reverse * 2 + (Value And 1)
        Numbits = Numbits - 1
        Value = Value \ 2
    Loop
End Function
Private Sub Init_Decompress(UncompressedSize As Long)
    Dim temp(), X As Long
    ReDim OutStream(UncompressedSize)
    Erase LitLen.code
    Erase LitLen.Lenght
    Erase Dist.code
    Erase Dist.Lenght
    ReDim LC.code(31)
    ReDim LC.Lenght(31)
    ReDim dc.code(31)
    ReDim dc.Lenght(31)
    temp() = Array(16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15)
    For X = 0 To UBound(temp)
        LenOrder(X) = temp(X)
    Next
    temp() = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258)
    For X = 0 To UBound(temp)
        LC.code(X) = temp(X)
    Next
    temp() = Array(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0)
    For X = 0 To UBound(temp)
        LC.Lenght(X) = temp(X)
    Next
    temp() = Array(1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577, 32769, 49153)
    For X = 0 To UBound(temp)
        dc.code(X) = temp(X)
    Next
    temp() = Array(0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13, 14, 14)
    For X = 0 To UBound(temp)
        dc.Lenght(X) = temp(X)
    Next
    For X = 0 To 16
        BitMask(X) = 2 ^ X - 1
        Pow2(X) = 2 ^ X
    Next
    OutPos = 0
    Inpos = 0
    ByteBuff = 0
    BitNum = 0
End Sub
Private Sub PutByte(char As Byte)
    If OutPos > UBound(OutStream) Then ReDim Preserve OutStream(OutPos + 1000)
    OutStream(OutPos) = char
    OutPos = OutPos + 1
End Sub
Private Sub NeedBits(Numbits As Long)
    While BitNum < Numbits
        If Inpos > UBound(InStream) Then Exit Sub
        ByteBuff = ByteBuff + (InStream(Inpos) * Pow2(BitNum))
        BitNum = BitNum + 8
        Inpos = Inpos + 1
    Wend
End Sub
Private Sub DropBits(Numbits As Long)
    ByteBuff = ByteBuff \ Pow2(Numbits)
    BitNum = BitNum - Numbits
End Sub
Private Function GetBits(Numbits As Long) As Long
    While BitNum < Numbits
        ByteBuff = ByteBuff + (InStream(Inpos) * Pow2(BitNum))
        BitNum = BitNum + 8
        Inpos = Inpos + 1
    Wend
    GetBits = ByteBuff And BitMask(Numbits)
    ByteBuff = ByteBuff \ Pow2(Numbits)
    BitNum = BitNum - Numbits
End Function
Private Function DeFilter(Dat() As Byte) As Byte()
    Dim NewDat() As Byte, Y As Long, iVal As Long, n As Long, StartByte As Long, DestByte As Long
    Dim BPRow As Long, X As Long, RowBytes() As Byte, PrevRowBytes() As Byte, i As Long
    iVal = Interval()
    BPRow = DataPerRow()
    ReDim NewDat(UBound(Dat) - Me.Height)
    ReDim PrevRowBytes(DataPerRow() - 2)
    ReDim RowBytes(DataPerRow() - 2)
    For Y = 0 To Me.Height - 1
        StartByte = BPRow * Y
        DestByte = StartByte - Y
        X = 0
        CopyMemory RowBytes(0), Dat(StartByte + 1), BPRow - 1
        Select Case Dat(StartByte)
            Case 0 'None
            Case 1 'Sub
                ReverseSub RowBytes, iVal
            Case 2 'Up
                ReverseUp RowBytes, PrevRowBytes
            Case 3 'Average
                ReverseAverage RowBytes, PrevRowBytes, iVal
            Case 4 'Paeth
                ReversePaeth RowBytes, PrevRowBytes, iVal
        End Select
        CopyMemory NewDat(DestByte), RowBytes(0), BPRow - 1
        PrevRowBytes = RowBytes
    Next Y
DeFilter = NewDat
End Function
Private Function Interval() As Long
    Interval = BitsPerPixel() \ 8
    If Interval = 0 Then Interval = 1
End Function
Private Function BitsPerPixel() As Long
    Dim Bpp As Long
    If RBD = 0 Then
        Bpp = Me.Bitdepht
    Else
        Bpp = RBD
    End If
    If BPPprivat <> Bpp And BPPprivat <> 0 Then Bpp = BPPprivat
    Select Case Me.ColorType
        Case 0, 3: BitsPerPixel = Bpp
        Case 2: BitsPerPixel = 3 * Bpp
        Case 6: BitsPerPixel = 4 * Bpp
        Case 4: BitsPerPixel = 2 * Bpp
    End Select
End Function
Private Function DataPerRow() As Long
    DataPerRow = (Me.Width * BitsPerPixel() + 7) \ 8 + 1
End Function
Private Sub ReverseAverage(CurRow() As Byte, PrevRow() As Byte, Interval As Long)
    Dim PrevOff As Long, PrevVal As Byte, BPRow As Long, n As Long, X As Integer
    BPRow = UBound(CurRow) + 1
    For n = 0 To BPRow - 1
        PrevOff = n - Interval
        If PrevOff >= 0 Then PrevVal = CurRow(PrevOff)
        X = CurRow(n) + (CInt(PrevRow(n)) + CInt(PrevVal)) \ 2
        CopyMemory CurRow(n), X, 1
    Next n
End Sub
Private Sub ReversePaeth(CurRow() As Byte, PrevRow() As Byte, Interval As Long)
    Dim BPRow As Long, n As Long, X As Integer, LeftPixOff As Long, LeftPix As Byte, UpperLeftPix As Byte
    BPRow = UBound(CurRow) + 1
    For n = 0 To BPRow - 1
        LeftPixOff = n - Interval
        If LeftPixOff >= 0 Then
            LeftPix = CurRow(LeftPixOff)
            UpperLeftPix = PrevRow(LeftPixOff)
        End If
        X = CInt(CurRow(n)) + CInt(PaethPredictor(LeftPix, PrevRow(n), UpperLeftPix))
        CopyMemory CurRow(n), X, 1
    Next n
End Sub
Private Sub ReverseUp(CurRow() As Byte, PrevRow() As Byte)
    Dim PrevVal As Byte, BPRow As Long, n As Long, X As Integer
    BPRow = UBound(CurRow) + 1
    For n = 0 To BPRow - 1
        PrevVal = PrevRow(n)
        X = CInt(CurRow(n)) + CInt(PrevVal)
        CopyMemory CurRow(n), X, 1
    Next n
End Sub
Private Sub ReverseSub(CurRow() As Byte, Interval As Long)
    Dim PrevOff As Long, PrevVal As Byte, BPRow As Long, n As Long, X As Integer
    BPRow = UBound(CurRow) + 1
    For n = 0 To BPRow - 1
        PrevOff = n - Interval
        If PrevOff >= 0 Then PrevVal = CurRow(PrevOff)
        X = CInt(CurRow(n)) + CInt(PrevVal)
        CopyMemory CurRow(n), X, 1
    Next n
End Sub
Private Function PaethPredictor(Left As Byte, Above As Byte, UpperLeft As Byte) As Byte
    Dim pA As Integer, pB As Integer, pC As Integer, p As Integer
    p = CInt(Left) + CInt(Above) - CInt(UpperLeft)
    pA = Abs(p - Left)
    pB = Abs(p - Above)
    pC = Abs(p - UpperLeft)
    If (pA <= pB) And (pA <= pC) Then
        PaethPredictor = Left
    ElseIf pB <= pC Then
        PaethPredictor = Above
    Else
        PaethPredictor = UpperLeft
    End If
End Function
Private Sub ReverseRGB(Dat() As Byte)
    Dim n As Long, Tmp As Byte
    On Error Resume Next
    For n = 0 To UBound(Dat) Step 3
        Tmp = Dat(n)
        Dat(n) = Dat(n + 2)
        Dat(n + 2) = Tmp
    Next n
End Sub
Private Sub Conv16To8(Dat() As Byte)
    Dim n As Long, DestDat() As Byte, DestOff As Long
    ReDim DestDat((UBound(Dat) + 1) \ 2 - 1)
    For n = 0 To UBound(Dat) Step 2
        DestDat(DestOff) = Dat(n)
        DestOff = DestOff + 1
    Next n
    Dat = DestDat
End Sub
Private Sub Align32(BitCount As Integer, Dat() As Byte)
    Dim RowBytes As Long, SrcRowBytes As Long, Y As Long, Dest() As Byte, SrcOff As Long, DestOff As Long
    If BitCount = 32 Then Exit Sub
    RowBytes = LineBytes(Me.Width, BitCount)
    SrcRowBytes = DataPerRow() - 1
    If Me.ColorType = 4 Then SrcRowBytes = SrcRowBytes / 2 'Alpha
    If RowBytes = SrcRowBytes Then Exit Sub
    ReDim Dest(RowBytes * Me.Height - 1)
    For Y = 0 To Me.Height - 1
        SrcOff = Y * SrcRowBytes
        DestOff = Y * RowBytes
        CopyMemory Dest(DestOff), Dat(SrcOff), SrcRowBytes
    Next Y
    Dat = Dest
End Sub
Private Function LineBytes(Width As Long, BitCount As Integer) As Long
    LineBytes = ((Width * BitCount + 31) \ 32) * 4
End Function
Private Sub ReverseRGBA(Dat() As Byte)
    Dim n As Long, Tmp As Byte
    For n = 0 To UBound(Dat) Step 4
        Tmp = Dat(n)
        If n + 2 > UBound(Dat) Then Exit For
        Dat(n) = Dat(n + 2)
        Dat(n + 2) = Tmp
    Next n
End Sub
Private Sub Pal2To8(Width As Long, Height As Long, Dat() As Byte, RowBytes As Long)
    Dim DestDat() As Byte, DestRowBytes As Long, n As Long, Px As Byte, DestOff As Long, X As Long, Y As Long
    DestRowBytes = LineBytes(Width, 8)
    ReDim DestDat(DestRowBytes * Height - 1)
    For Y = 0 To Height - 1
        DestOff = Y * DestRowBytes
        For X = 0 To Width - 1
            n = Y * (RowBytes - 1) + X \ 4
            If (X Mod 4) <> 3 Then
                Px = (Dat(n) \ 4 ^ (3 - (X Mod 4))) And 3
            Else
                Px = Dat(n) And 3
            End If
            DestDat(DestOff) = Px
            DestOff = DestOff + 1
        Next X
    Next Y
    Dat = DestDat
End Sub
Private Sub GrayAToRGBA(Dat() As Byte)
    Dim n As Long, DestDat() As Byte, DestOff As Long
    ReDim DestDat((UBound(Dat) + 1) * 2 - 1)
    For n = 0 To UBound(Dat) Step 2
        DestDat(DestOff) = Dat(n)
        DestDat(DestOff + 1) = Dat(n)
        DestDat(DestOff + 2) = Dat(n)
        DestDat(DestOff + 3) = Dat(n + 1)
        DestOff = DestOff + 4
    Next n
    Dat = DestDat
End Sub
Private Function DeFilterInterlaced(Buffer() As Byte) As Byte()
Dim Stand As String
Dim X As Long
Dim Y As Long
Dim ZL As Long
Dim Bpp As Long
Dim Bufferstand As Long
Dim Zeilenbuffer() As Byte
Dim Height8 As Long
Dim Rest8 As Long
Dim MengeZeilen As Long
Dim i As Long
Dim Filterbyte As Byte
Dim PrevRowBytes() As Byte
Dim ZwischenBuffer() As Byte
Dim Nr As Long
Dim ZZ As Long
Dim BytesPerPixel As Long
Dim ZLBytes As Long
Y = Me.Height
X = Me.Width
Bpp = BitsPerPixel
If Bpp >= 8 Then
BytesPerPixel = Bpp / 8
Else
BytesPerPixel = 1
End If
ReDim ZwischenBuffer((X * Y * BytesPerPixel) - 1)
Rest8 = Y Mod 8
Height8 = (Y - Rest8) / 8
Stand = "1" 'Durchlauf 1
ZL = BerechneZeilenlänge(X, Bpp, Stand)
If ZL > 0 Then
ReDim PrevRowBytes(ZL - 1)
MengeZeilen = Height8
If Rest8 > 0 Then
MengeZeilen = MengeZeilen + 1
End If
For i = 1 To MengeZeilen
ReDim Zeilenbuffer(ZL - 1)
CopyMemory Zeilenbuffer(0), Buffer(Bufferstand + 1), ZL
Filterbyte = Buffer(Bufferstand)
FilterInter Zeilenbuffer, Filterbyte, PrevRowBytes
Bufferstand = Bufferstand + ZL + 1
If Bpp < 8 Then
ZLBytes = BerechneZeilenlänge(X, 8, Stand)
Else
ZLBytes = 0
End If
PutBuffer ZwischenBuffer, Zeilenbuffer, 1, 1, i, ZLBytes
Next i
End If
Stand = "5" 'Durchlauf 2
ZL = BerechneZeilenlänge(X, Bpp, Stand)
If ZL > 0 Then
ReDim PrevRowBytes(ZL - 1)
MengeZeilen = Height8
If Rest8 > 0 Then
MengeZeilen = MengeZeilen + 1
End If
For i = 1 To MengeZeilen
ReDim Zeilenbuffer(ZL - 1)
CopyMemory Zeilenbuffer(0), Buffer(Bufferstand + 1), ZL
Filterbyte = Buffer(Bufferstand)
FilterInter Zeilenbuffer, Filterbyte, PrevRowBytes
Bufferstand = Bufferstand + ZL + 1
If Bpp < 8 Then
ZLBytes = BerechneZeilenlänge(X, 8, Stand)
Else
ZLBytes = 0
End If
PutBuffer ZwischenBuffer, Zeilenbuffer, 2, 1, i, ZLBytes
Next i
End If
Stand = "15" 'Durchlauf 3
ZL = BerechneZeilenlänge(X, Bpp, Stand)
If ZL > 0 Then
ReDim PrevRowBytes(ZL - 1)
MengeZeilen = Height8
If Rest8 > 4 Then
MengeZeilen = MengeZeilen + 1
End If
For i = 1 To MengeZeilen
ReDim Zeilenbuffer(ZL - 1)
CopyMemory Zeilenbuffer(0), Buffer(Bufferstand + 1), ZL
Filterbyte = Buffer(Bufferstand)
FilterInter Zeilenbuffer, Filterbyte, PrevRowBytes
Bufferstand = Bufferstand + ZL + 1
If Bpp < 8 Then
ZLBytes = BerechneZeilenlänge(X, 8, Stand)
Else
ZLBytes = 0
End If
PutBuffer ZwischenBuffer, Zeilenbuffer, 3, 5, i, ZLBytes
Next i
End If
Stand = "37" 'Durchlauf 4 - Zeile 1 - 2
ZZ = 1
ZL = BerechneZeilenlänge(X, Bpp, Stand)
If ZL > 0 Then
ReDim PrevRowBytes(ZL - 1)
MengeZeilen = Height8 * 2
If Rest8 > 0 Then
MengeZeilen = MengeZeilen + 1
End If
If Rest8 > 4 Then
MengeZeilen = MengeZeilen + 1
End If
Nr = 1
For i = 1 To MengeZeilen
ReDim Zeilenbuffer(ZL - 1)
CopyMemory Zeilenbuffer(0), Buffer(Bufferstand + 1), ZL
Filterbyte = Buffer(Bufferstand)
FilterInter Zeilenbuffer, Filterbyte, PrevRowBytes
Bufferstand = Bufferstand + ZL + 1
If Bpp < 8 Then
ZLBytes = BerechneZeilenlänge(X, 8, Stand)
Else
ZLBytes = 0
End If
PutBuffer ZwischenBuffer, Zeilenbuffer, 4, Nr, ZZ, ZLBytes
If Nr = 1 Then
Nr = 5
Else
Nr = 1
ZZ = ZZ + 1
End If
Next i
End If
Stand = "1357" 'Durchlauf 5 - Zeile 1 - 2
ZL = BerechneZeilenlänge(X, Bpp, Stand)
If ZL > 0 Then
ReDim PrevRowBytes(ZL - 1)
MengeZeilen = Height8 * 2
If Rest8 > 2 Then
MengeZeilen = MengeZeilen + 1
End If
If Rest8 > 6 Then
MengeZeilen = MengeZeilen + 1
End If
ZZ = 1
Nr = 3
For i = 1 To MengeZeilen
ReDim Zeilenbuffer(ZL - 1)
CopyMemory Zeilenbuffer(0), Buffer(Bufferstand + 1), ZL
Filterbyte = Buffer(Bufferstand)
FilterInter Zeilenbuffer, Filterbyte, PrevRowBytes
Bufferstand = Bufferstand + ZL + 1
If Bpp < 8 Then
ZLBytes = BerechneZeilenlänge(X, 8, Stand)
Else
ZLBytes = 0
End If
PutBuffer ZwischenBuffer, Zeilenbuffer, 5, Nr, ZZ, ZLBytes
Select Case Nr
Case 3
Nr = 7
Case 7
Nr = 3
ZZ = ZZ + 1
End Select
Next i
End If
Stand = "2468" 'Durchlauf 6 - Zeile 1 - 4
ZL = BerechneZeilenlänge(X, Bpp, Stand)
If ZL > 0 Then
ReDim PrevRowBytes(ZL - 1)
ZZ = 1
Nr = 1
MengeZeilen = Height8 * 4
If Rest8 > 0 Then
MengeZeilen = MengeZeilen + 1
End If
If Rest8 > 2 Then
MengeZeilen = MengeZeilen + 1
End If
If Rest8 > 4 Then
MengeZeilen = MengeZeilen + 1
End If
If Rest8 > 6 Then
MengeZeilen = MengeZeilen + 1
End If
For i = 1 To MengeZeilen
ReDim Zeilenbuffer(ZL - 1)
CopyMemory Zeilenbuffer(0), Buffer(Bufferstand + 1), ZL
Filterbyte = Buffer(Bufferstand)
FilterInter Zeilenbuffer, Filterbyte, PrevRowBytes
Bufferstand = Bufferstand + ZL + 1
If Bpp < 8 Then
ZLBytes = BerechneZeilenlänge(X, 8, Stand)
Else
ZLBytes = 0
End If
PutBuffer ZwischenBuffer, Zeilenbuffer, 6, Nr, ZZ, ZLBytes
Select Case Nr
Case 1
Nr = 3
Case 3
Nr = 5
Case 5
Nr = 7
Case 7
Nr = 1
ZZ = ZZ + 1
End Select
Next i
End If
Stand = "12345678" 'Durchlauf 7 - Zeile 1 - 4
ZL = BerechneZeilenlänge(X, Bpp, Stand)
If ZL > 0 Then
ReDim PrevRowBytes(ZL - 1)
ZZ = 1
Nr = 2
MengeZeilen = Height8 * 4
If Rest8 > 1 Then
MengeZeilen = MengeZeilen + 1
End If
If Rest8 > 3 Then
MengeZeilen = MengeZeilen + 1
End If
If Rest8 > 5 Then
MengeZeilen = MengeZeilen + 1
End If
If Rest8 > 7 Then
MengeZeilen = MengeZeilen + 1
End If
For i = 1 To MengeZeilen
ReDim Zeilenbuffer(ZL - 1)
CopyMemory Zeilenbuffer(0), Buffer(Bufferstand + 1), ZL
Filterbyte = Buffer(Bufferstand)
FilterInter Zeilenbuffer, Filterbyte, PrevRowBytes
Bufferstand = Bufferstand + ZL + 1
If Bpp < 8 Then
ZLBytes = BerechneZeilenlänge(X, 8, Stand)
Else
ZLBytes = 0
End If
PutBuffer ZwischenBuffer, Zeilenbuffer, 7, Nr, ZZ, ZLBytes
Select Case Nr
Case 2
Nr = 4
Case 4
Nr = 6
Case 6
Nr = 8
Case 8
Nr = 2
ZZ = ZZ + 1
End Select
Next i
End If
DeFilterInterlaced = ZwischenBuffer
End Function
Private Function BerechneZeilenlänge(X As Long, Bpp As Long, Stand As String) As Long
Dim Hilfslong As Long
Dim Längenrest As Long
Dim Länge8 As Long
Dim Testlong As Long
Dim Anzahl8 As Long
Dim AnzahlBits As Long
Dim Bytesrest As Long
Dim NBytes As Long
Dim AnzRB As Long
Dim Rest As Long
Dim MengeBits As Long
Dim i As Long
Dim BiggerAs As Long
Dim Menge As Long
MengeBits = Len(Stand)
Längenrest = X Mod 8
BiggerAs = 0
Menge = 0
For i = 1 To MengeBits
If CLng(Mid(Stand, i, 1)) <= Längenrest Then
Menge = Menge + 1
Else
Exit For
End If
Next i
If Bpp < 8 Then
If Längenrest > 0 Then
 Rest = Bpp * Menge
 Else
 Rest = 0
End If
Else
Rest = Menge * (Bpp / 8)
End If
Anzahl8 = (X - Längenrest) / 8
AnzahlBits = Anzahl8 * Bpp * MengeBits
Bytesrest = AnzahlBits Mod 8
NBytes = (AnzahlBits - Bytesrest) / 8
Select Case Bpp
Case Is < 8
Rest = Rest + Bytesrest
Testlong = Rest Mod 8
AnzRB = (Rest - Testlong) / 8
If Testlong <> 0 Then AnzRB = AnzRB + 1
BerechneZeilenlänge = NBytes + AnzRB
Case Else
BerechneZeilenlänge = NBytes + Rest
End Select
End Function
Private Sub FilterInter(RowBytes() As Byte, Filterbyte As Byte, PrevRowBytes() As Byte)
    Dim iVal As Long
    iVal = Interval()
    Select Case Filterbyte
        Case 0 'None
        Case 1 'Sub
            ReverseSub RowBytes, iVal
        Case 2 'Up
            ReverseUp RowBytes, PrevRowBytes
        Case 3 'Average
            ReverseAverage RowBytes, PrevRowBytes, iVal
        Case 4 'Paeth
            ReversePaeth RowBytes, PrevRowBytes, iVal
    End Select
    PrevRowBytes = RowBytes
End Sub
Private Sub PutBuffer(Buffer() As Byte, Zeilenbuffer() As Byte, Zeilentyp As Byte, Zeilennummer As Long, Zeilenzähler As Long, Zeilenlänge As Long)
Dim Anfang As Long
Dim Achtschritt As Long
Dim Zeile As Long
Dim Zeilenanfang As Long
Dim i As Long
Dim Bufferstand As Long
Dim Zeilenstand As Long
Dim Größe As Long
Dim BytesPerPixel As Long
Dim Bpp As Long
Bpp = BitsPerPixel
If Bpp >= 8 Then
BytesPerPixel = Bpp / 8
Else
BytesPerPixel = 1
BytesToBits Zeilenbuffer, Me.Bitdepht, Zeilenlänge
End If
Größe = UBound(Zeilenbuffer) + 1
Zeilenanfang = Me.Width * (Zeilennummer - 1) * BytesPerPixel
Achtschritt = Me.Width * 8 * BytesPerPixel
Anfang = (Achtschritt * (Zeilenzähler - 1)) + Zeilenanfang
'Zeilentyp: 1 = 1; 2 = 5; 3 = 1+5; 4 = 3+7; 5 = 1+3+5+7; 6 = 2+4+6+8; 7 = 1-8;
Bufferstand = Anfang
Select Case Zeilentyp
Case 1
Do While Zeilenstand < Größe
CopyMemory Buffer(Bufferstand), Zeilenbuffer(Zeilenstand), BytesPerPixel
Bufferstand = Bufferstand + (8 * BytesPerPixel)
Zeilenstand = Zeilenstand + BytesPerPixel
Loop
Case 2
Bufferstand = Bufferstand + (4 * BytesPerPixel)
Do While Zeilenstand < Größe
CopyMemory Buffer(Bufferstand), Zeilenbuffer(Zeilenstand), BytesPerPixel
Bufferstand = Bufferstand + (8 * BytesPerPixel)
Zeilenstand = Zeilenstand + BytesPerPixel
Loop
Case 3
Do While Zeilenstand < Größe
CopyMemory Buffer(Bufferstand), Zeilenbuffer(Zeilenstand), BytesPerPixel
If Zeilenstand + BytesPerPixel < Größe Then
CopyMemory Buffer(Bufferstand + (4 * BytesPerPixel)), Zeilenbuffer(Zeilenstand + BytesPerPixel), BytesPerPixel
End If
Bufferstand = Bufferstand + (8 * BytesPerPixel)
Zeilenstand = Zeilenstand + (2 * BytesPerPixel)
Loop
Case 4
Bufferstand = Bufferstand + (2 * BytesPerPixel)
Do While Zeilenstand < Größe
CopyMemory Buffer(Bufferstand), Zeilenbuffer(Zeilenstand), BytesPerPixel
If Zeilenstand + BytesPerPixel < Größe Then
CopyMemory Buffer(Bufferstand + (4 * BytesPerPixel)), Zeilenbuffer(Zeilenstand + BytesPerPixel), BytesPerPixel
End If
Bufferstand = Bufferstand + (8 * BytesPerPixel)
Zeilenstand = Zeilenstand + (2 * BytesPerPixel)
Loop
Case 5
Do While Zeilenstand < Größe
CopyMemory Buffer(Bufferstand), Zeilenbuffer(Zeilenstand), BytesPerPixel
If Zeilenstand + BytesPerPixel < Größe Then
CopyMemory Buffer(Bufferstand + (2 * BytesPerPixel)), Zeilenbuffer(Zeilenstand + BytesPerPixel), BytesPerPixel
End If
If Zeilenstand + (2 * BytesPerPixel) < Größe Then
CopyMemory Buffer(Bufferstand + (4 * BytesPerPixel)), Zeilenbuffer(Zeilenstand + (2 * BytesPerPixel)), BytesPerPixel
End If
If Zeilenstand + (3 * BytesPerPixel) < Größe Then
CopyMemory Buffer(Bufferstand + (6 * BytesPerPixel)), Zeilenbuffer(Zeilenstand + (3 * BytesPerPixel)), BytesPerPixel
End If
Bufferstand = Bufferstand + (8 * BytesPerPixel)
Zeilenstand = Zeilenstand + (4 * BytesPerPixel)
Loop
Case 6
Bufferstand = Bufferstand + BytesPerPixel
Do While Zeilenstand < Größe
CopyMemory Buffer(Bufferstand), Zeilenbuffer(Zeilenstand), BytesPerPixel
If Zeilenstand + BytesPerPixel < Größe Then
CopyMemory Buffer(Bufferstand + (2 * BytesPerPixel)), Zeilenbuffer(Zeilenstand + BytesPerPixel), BytesPerPixel
End If
If Zeilenstand + (2 * BytesPerPixel) < Größe Then
CopyMemory Buffer(Bufferstand + (4 * BytesPerPixel)), Zeilenbuffer(Zeilenstand + (2 * BytesPerPixel)), BytesPerPixel
End If
If Zeilenstand + (3 * BytesPerPixel) < Größe Then
CopyMemory Buffer(Bufferstand + (6 * BytesPerPixel)), Zeilenbuffer(Zeilenstand + (3 * BytesPerPixel)), BytesPerPixel
End If
Bufferstand = Bufferstand + (8 * BytesPerPixel)
Zeilenstand = Zeilenstand + (4 * BytesPerPixel)
Loop
Case 7
CopyMemory Buffer(Bufferstand), Zeilenbuffer(0), UBound(Zeilenbuffer) + 1
End Select
End Sub
Private Sub BytesToBits(Bytefeld() As Byte, Bitanzahl As Byte, Größe As Long)
Dim i As Long
Dim Übergabe() As Byte
Dim Wandeln() As Byte
Dim EinGr As Long
Dim z As Long
EinGr = UBound(Bytefeld) + 1
Select Case Bitanzahl
Case 1
ReDim Übergabe((EinGr * 8) - 1)
For i = 0 To EinGr - 1
ByteToEinBit Bytefeld(i), Wandeln
CopyMemory Übergabe(z), Wandeln(0), 8
z = z + 8
Next i
Case 2
ReDim Übergabe((EinGr * 4) - 1)
For i = 0 To EinGr - 1
ByteToZweiBit Bytefeld(i), Wandeln
CopyMemory Übergabe(z), Wandeln(0), 4
z = z + 4
Next i
Case 4
ReDim Übergabe((EinGr * 2) - 1)
For i = 0 To EinGr - 1
ByteToVierBit Bytefeld(i), Wandeln
CopyMemory Übergabe(z), Wandeln(0), 2
z = z + 2
Next i
End Select
ReDim Preserve Übergabe(Größe - 1)
Bytefeld = Übergabe
End Sub
Private Sub ByteToZweiBit(Number As Byte, Wandeln() As Byte)
    Dim a As Byte
    ReDim Wandeln(3)
    Wandeln(3) = Number And 3
    a = Number And 12
    Wandeln(2) = a / 4
    a = Number And 48
    Wandeln(1) = a / 16
    a = Number And 192
    Wandeln(0) = a / 64
End Sub
Private Sub ByteToEinBit(Number As Byte, Wandeln() As Byte)
    Dim a As Byte
    ReDim Wandeln(7)
    Wandeln(7) = Number And 1
    a = Number And 2
    Wandeln(6) = a / 2
    a = Number And 4
    Wandeln(5) = a / 4
    a = Number And 8
    Wandeln(4) = a / 8
    a = Number And 16
    Wandeln(3) = a / 16
    a = Number And 32
    Wandeln(2) = a / 32
    a = Number And 64
    Wandeln(1) = a / 64
    a = Number And 128
    Wandeln(0) = a / 128
End Sub
Private Sub ByteToVierBit(Number As Byte, Wandeln() As Byte)
    Dim a As Byte
    ReDim Wandeln(1)
    Wandeln(1) = Number And 15
    a = Number And 240
    Wandeln(0) = a / 16
End Sub
Public Property Get ErrorNumber() As Long
    ErrorNumber = m_ErrorNumber
End Property
Private Sub MakeAlpha(PicObject As Object, Buffer() As Byte, Optional X As Long = 0, Optional Y As Long = 0)
Dim Myx As Long, Myy As Long, DatOff As Long
Dim r As Long, g As Long, b As Long, a As Long
Dim sR As Long, sG As Long, sB As Long
Dim dR As Long, dG As Long, dB As Long
Dim DestData() As Byte, bytesperrow As Long
Dim DestOff As Long, DestHdr As BITMAPINFOHEADER
Dim MemDC As Long, hBmp As Long, hOldBmp As Long
Dim SrcData() As Byte
Dim hDC As Long
On Error Resume Next
hDC = PicObject.hDC
If Err.Number = 91 Then
 ReDim SrcData(UBound(Buffer))
  bytesperrow = LineBytes(Me.Width, 24)
If m_OBCol = False Then
  FillColorArray SrcData, Me.BkgdColor, bytesperrow
Else
  FillColorArray SrcData, m_OBCol, bytesperrow
End If
 ReDim DestData(bytesperrow * Me.Height - 1)
Err.Clear
Else
If PicObject.Width < Me.Width * Screen.TwipsPerPixelX Then
 PicObject.Width = Screen.TwipsPerPixelX * Me.Width + 100
End If
If PicObject.Height < Me.Height * Screen.TwipsPerPixelY Then
 PicObject.Height = Screen.TwipsPerPixelY * Me.Height + 100
End If
 hDC = PicObject.hDC
 bytesperrow = LineBytes(Me.Width, 24)
 ReDim DestData(bytesperrow * Me.Height - 1)
 ReDim SrcData(UBound(Buffer))
 DestHdr.BitCount = 24
 DestHdr.Height = Me.Height
 DestHdr.Width = Me.Width
 DestHdr.Planes = 1
 DestHdr.Size = 40
 MemDC = CreateCompatibleDC(hDC)
 hBmp = CreateCompatibleBitmap(hDC, Me.Width, Me.Height)
 hOldBmp = SelectObject(MemDC, hBmp)
 BitBlt MemDC, 0, 0, Me.Width, Me.Height, hDC, X, Y, vbSrcCopy
 GetDIBits MemDC, hBmp, 0, Me.Height, SrcData(0), DestHdr, 0
SelectObject hOldBmp, MemDC
 DeleteObject hBmp
 DeleteDC MemDC
End If
 For Myy = 0 To Me.Height - 1
 For Myx = 0 To Me.Width - 1
 DestOff = Myy * bytesperrow + Myx * 3
 sR = SrcData(DestOff + 2)
 sG = SrcData(DestOff + 1)
 sB = SrcData(DestOff)
 b = Buffer(DatOff)
 g = Buffer(DatOff + 1)
 r = Buffer(DatOff + 2)
 a = Buffer(DatOff + 3)
If a = 255 Then
 DestData(DestOff + 2) = r
 DestData(DestOff + 1) = g
 DestData(DestOff) = b
 ElseIf a = 0 Then
 DestData(DestOff + 2) = sR
 DestData(DestOff + 1) = sG
 DestData(DestOff) = sB
 Else
 dR = r * a + (255 - a) * sR + 255
 dG = g * a + (255 - a) * sG + 255
 dB = b * a + (255 - a) * sB + 255
 CopyMemory DestData(DestOff + 2), ByVal VarPtr(dR) + 1, 1
 CopyMemory DestData(DestOff + 1), ByVal VarPtr(dG) + 1, 1
 CopyMemory DestData(DestOff), ByVal VarPtr(dB) + 1, 1
End If
 DatOff = DatOff + 4
Next Myx
Next Myy
 Buffer = DestData
End Sub
Private Sub MirrorData(Dat() As Byte, RowBytes As Long)
    Dim NewDat() As Byte, Y As Long, Height As Long, StartLine As Long, DestLine As Long
    ReDim NewDat(UBound(Dat))
    Height = (UBound(Dat) + 1) \ RowBytes
    For Y = 0 To Height - 1
        StartLine = Y * RowBytes
        DestLine = (Height - Y - 1) * RowBytes
        CopyMemory NewDat(DestLine), Dat(StartLine), RowBytes
    Next Y
    Dat = NewDat
End Sub
Public Property Get HaveAlpha() As Boolean
    HaveAlpha = m_hAlpha
End Property
Public Property Get HaveTransparence() As Boolean
    HaveTransparence = m_hTrans
End Property
Public Property Let SetTrans(ByVal vNewValue As Boolean)
    m_sTrans = vNewValue
End Property
Public Property Let SetAlpha(ByVal vNewValue As Boolean)
    m_sAlpha = vNewValue
End Property
Private Sub PalToRGBA(Width As Long, Height As Long, BitDepth As Integer, Dat() As Byte)
 Dim DestDat() As Byte, n As Long, PalEntry As Byte
 Dim DestOff As Long, TrnsBnd As Long
 Dim Testint As Integer
 Dim X As Long, Y As Long, WidthBytes As Long
 Dim Pal() As RGBTriple
 Dim IdataLen As Long
 Dim i As Long
 Dim Anzahl As Long
 ReDim DestDat(4 * Width * Height - 1)
 TrnsBnd = UBound(trns)
 WidthBytes = LineBytes(Width, BitDepth)
If Me.ColorType = 0 Then
 Palettenbyte = InitColorTable_Grey(Bitdepht)
 Anzahl = UBound(Palettenbyte) / 3
 Testint = (trns(1))
 ReDim trns(Anzahl - 1)
 For i = 0 To Anzahl - 1
 trns(i) = 255
Next i
 trns(Testint) = 0
 TrnsBnd = UBound(trns)
End If
 ReDim Pal(((UBound(Palettenbyte) + 1) / 3) - 1)
 Colorused = UBound(Pal) + 1
CopyMemory Pal(0), Palettenbyte(0), UBound(Palettenbyte) + 1
Select Case BitDepth
 Case 8
 For Y = 0 To Height - 1
 For X = 0 To Width - 1
 n = Y * WidthBytes + X
 PalEntry = Dat(n)
 With Pal(PalEntry)
 DestDat(DestOff) = .Blue
 DestDat(DestOff + 1) = .Green
 DestDat(DestOff + 2) = .Red
End With
If PalEntry <= TrnsBnd Then
 DestDat(DestOff + 3) = trns(PalEntry)
 Else
 DestDat(DestOff + 3) = 255
End If
 DestOff = DestOff + 4
Next X
Next Y
 Case 4
 For Y = 0 To Height - 1
 For X = 0 To Width - 1
 n = Y * WidthBytes + X \ 2

If (X Mod 2) = 1 Then
 PalEntry = Dat(n) And 15
 Else
 PalEntry = (Dat(n) \ 16) And 15
End If
 With Pal(PalEntry)
 DestDat(DestOff) = .Blue
 DestDat(DestOff + 1) = .Green
 DestDat(DestOff + 2) = .Red
End With
If PalEntry <= TrnsBnd Then
 DestDat(DestOff + 3) = trns(PalEntry)
 Else
 DestDat(DestOff + 3) = 255
End If
 DestOff = DestOff + 4
Next X
Next Y
 Case 1
 For Y = 0 To Height - 1
 For X = 0 To Width - 1
 n = Y * WidthBytes + X \ 8
If (X Mod 8) <> 7 Then
 PalEntry = (Dat(n) \ 2 ^ (7 - X Mod 8)) And 1
 Else
 PalEntry = Dat(n) And 1
End If
 With Pal(PalEntry)
 DestDat(DestOff) = .Blue
 DestDat(DestOff + 1) = .Green
 DestDat(DestOff + 2) = .Red
End With
If PalEntry <= TrnsBnd Then
 DestDat(DestOff + 3) = trns(PalEntry)
 Else
 DestDat(DestOff + 3) = 255
End If
 DestOff = DestOff + 4
Next X
Next Y
End Select
Dat = DestDat
End Sub
Private Sub MakeRGBTransparent(Buffer() As Byte)
Dim i As Long
Dim Wo As Long
Dim Testlong As Long
Dim Testint As Integer
Dim Farblong As Long
Dim Übergabe() As Byte
Dim TestArray(5) As Byte
Dim Farbarray(5) As Byte
Dim l As Byte
Dim Ft As Long
Ft = Me.Bitdepht
Dim Größe As Long
Select Case Me.Bitdepht
Case 8
trns(0) = trns(1)
trns(2) = trns(3)
trns(4) = trns(5)
CopyMemory TestArray(0), trns(0), 6
Case 16
CopyMemory TestArray(0), trns(0), 6
End Select
Größe = (UBound(Buffer) + 1) / 3
ReDim Übergabe((Größe * 4) - 1)
Wo = 0
For i = 0 To UBound(Buffer) - 1 Step 3
CopyMemory Farbarray(0), Buffer(i), 6
CopyMemory Übergabe(Wo), Buffer(i), 3
If Farbarray(0) <> TestArray(0) Or Farbarray(1) <> TestArray(1) Or Farbarray(2) <> TestArray(2) Or Farbarray(3) <> TestArray(3) Or Farbarray(4) <> TestArray(4) Or Farbarray(5) <> TestArray(5) Then
Übergabe(Wo + 3) = 255
End If
Wo = Wo + 4
Next i
Buffer = Übergabe
End Sub
Public Property Get HasBKGDChunk() As Boolean
    HasBKGDChunk = m_hbkgd
End Property
Public Property Get BkgdColor() As Long
    BkgdColor = m_bkgdColor
End Property
Private Function ReadBkgd() As Long
Dim GBc As Long
Dim u As Byte
Dim bkLen As Long
Dim ValR As Integer, ValG As Integer, ValB As Integer
Dim r As Long, g As Long, b As Long
Dim BD As Byte
Dim IntVal As Integer, UInt As Long
Dim Testpal() As Byte
Dim Testcol(2) As Byte
bkLen = UBound(bkgd) + 1
BD = Me.Bitdepht
On Error GoTo Error
Select Case Me.ColorType
Case 3
If bkLen = 1 Then
If bkgd(0) > (UBound(Palettenbyte) - 1) Then
 GoTo Error
Else
 GBc = bkgd(0)
 CopyMemory Testcol(0), Palettenbyte(GBc * 3), 3
 m_bkgdColor = RGB(Testcol(0), Testcol(1), Testcol(2))
End If
Else: GoTo Error
End If
Case 0, 4
If bkLen = 2 Then
 CopyMemory IntVal, bkgd(0), Len(IntVal)
 Swap IntVal
 UInt = UnsignedInt(IntVal)
If UInt > (2 ^ BD - 1) Or (UInt < 0) Then
 GoTo Error
Else
 GBc = UInt
 Testpal = InitColorTable_Grey(Me.Bitdepht)
 CopyMemory Testcol(0), Testpal(GBc * 3), 3
 m_bkgdColor = RGB(Testcol(0), Testcol(1), Testcol(2))
End If
Else: GoTo Error
End If
Case 2, 6
If bkLen = 6 Then
 CopyMemory ValR, bkgd(0), 2
 CopyMemory ValG, bkgd(2), 2
 CopyMemory ValB, bkgd(4), 2
 Swap ValR
 Swap ValG
 Swap ValB
 r = UnsignedInt(ValR)
 g = UnsignedInt(ValG)
 b = UnsignedInt(ValB)
 m_bkgdColor = RGB(r / (2 ^ BD - 1) * 255, g / (2 ^ BD - 1) * 255, b / (2 ^ BD - 1) * 255)
Else: GoTo Error
End If
End Select
Exit Function
Error:
m_bkgdColor = 0
End Function
Private Function UnsignedInt(SignedInt As Integer) As Long
    UnsignedInt = CLng(SignedInt) And &HFFFF&
End Function
Private Sub Swap(Val As Integer)
    Dim Bytef(1) As Byte, u As Byte
    CopyMemory Bytef(0), ByVal VarPtr(Val), 2
    u = Bytef(0)
    Bytef(0) = Bytef(1)
    Bytef(1) = u
    CopyMemory ByVal VarPtr(Val), Bytef(0), 2
End Sub
Public Property Get Text() As String
    Text = m_text
End Property
Public Property Get zText() As String
    zText = m_ztext
End Property
Private Sub DecompressText(Inhalt() As Byte)
    Dim ztxt() As Byte, Ende As Long, Anfang As Long, Teststring As String, StringText As String, Größe As Long, Beendet As Boolean
    Größe = UBound(Inhalt)
    Ende = FindNull(Inhalt, Anfang)
    ReDim ztxt(Ende)
    CopyMemory ztxt(0), Inhalt(0), Ende + 1
    Teststring = StrConv(ztxt, vbUnicode)
    m_ztext = m_ztext & Teststring & Chr(0)
    Anfang = Ende + 5
    Ende = FindNull(Inhalt, Anfang)
    ReDim ztxt(Ende - Anfang)
    CopyMemory ztxt(0), Inhalt(Anfang), Ende - Anfang + 1
    Decompress ztxt, UBound(ztxt) * 12
    Teststring = StrConv(ztxt, vbUnicode)
    m_ztext = m_ztext & Teststring & Chr(0)
End Sub
Private Function FindNull(TestArray() As Byte, Start As Long) As Long
    Dim i As Long, Größe As Long
    Größe = UBound(TestArray)
    FindNull = Größe
    For i = Start To Größe
        If TestArray(i) = 0 Then
            FindNull = i - 1
            Exit For
        End If
    Next i
End Function
Public Property Get ModiTime() As String
    ModiTime = m_Time
End Property
Public Property Get gama() As Double
    gama = m_gama / 100000
End Property
Public Property Let BackgroundPicture(ByVal vNewValue As Object)
    Set m_BGPic = vNewValue
End Property
Private Sub FillColorArray(FArray() As Byte, Color As Long, bytesperrow As Long)
    Dim DA(3) As Byte, i As Long, u As Byte, Zähler As Long
    CopyMemory DA(0), ByVal VarPtr(Color), 3
    If DA(3) = 0 Then
        u = DA(0)
        DA(0) = DA(2)
        DA(2) = u
        u = DA(1)
        If DA(0) = DA(1) And DA(1) = DA(2) Then
            FillMemory FArray(0), UBound(FArray) + 1, DA(0)
        Else
            Zähler = 1
            For i = 0 To UBound(FArray) - 2 Step 3
                CopyMemory FArray(i), DA(0), 3
                If i = ((Zähler * bytesperrow) - 1) Or i = ((Zähler * bytesperrow) - 2) Then
                    i = Zähler * bytesperrow
                    i = bytesperrow * Zähler
                    Zähler = Zähler + 1
                End If
            Next i
        End If
    End If
End Sub
Public Sub SetOwnBkgndColor(OwnBkgndOn As Boolean, Optional ByVal BackColor As Long = 0)
    m_OwnBkgnd = OwnBkgndOn
    m_OBCol = BackColor
End Sub
Public Property Let PicBox(ByVal NewPicBox As Object)
    Set m_PicBox = NewPicBox
End Property
Public Sub SetToBkgrnd(SetToBG As Boolean, Optional X As Long = 0, Optional Y As Long = 0)
    m_Bgx = X
    m_Bgy = Y
    m_settoBG = SetToBG
End Sub

Public Sub InitColorTable_1(Optional Sorting As Integer = 1)
    Dim Fb1 As Byte, Fb2 As Byte
    Select Case Sorting
        Case 0
            Fb1 = 255
            Fb2 = 0
        Case 1
            Fb1 = 0
            Fb2 = 255
    End Select
    bm1.bmiColors(0).rgbRed = Fb1
    bm1.bmiColors(0).rgbGreen = Fb1
    bm1.bmiColors(0).rgbBlue = Fb1
    bm1.bmiColors(0).rgbReserved = 0
    bm1.bmiColors(1).rgbRed = Fb2
    bm1.bmiColors(1).rgbGreen = Fb2
    bm1.bmiColors(1).rgbBlue = Fb2
    bm1.bmiColors(1).rgbReserved = 0
End Sub
Public Sub InitColorTable_1Palette(Palettenbyte() As Byte)
    If UBound(Palettenbyte) = 5 Then
        bm1.bmiColors(0).rgbRed = Palettenbyte(0)
        bm1.bmiColors(0).rgbGreen = Palettenbyte(1)
        bm1.bmiColors(0).rgbBlue = Palettenbyte(2)
        bm1.bmiColors(0).rgbReserved = 0
        bm1.bmiColors(1).rgbRed = Palettenbyte(3)
        bm1.bmiColors(1).rgbGreen = Palettenbyte(4)
        bm1.bmiColors(1).rgbBlue = Palettenbyte(5)
        bm1.bmiColors(1).rgbReserved = 0
    Else
        InitColorTable_1
    End If
End Sub

Public Sub InitColorTable_8(bytearray() As Byte)
    'Construct the palette
    '==================================================
    Dim Palette8() As RGBTriple, nCount As Long
    ReDim Palette8(255)
    CopyMemory Palette8(0), bytearray(0), UBound(bytearray) + 1
    On Error Resume Next
    'Create Palette
    For nCount = 0 To 255
        bm8.bmiColors(nCount).rgbBlue = Palette8(nCount).Blue
        bm8.bmiColors(nCount).rgbGreen = Palette8(nCount).Green
        bm8.bmiColors(nCount).rgbRed = Palette8(nCount).Red
        bm8.bmiColors(nCount).rgbReserved = 0
    Next nCount
End Sub
Public Sub InitColorTable_4(bytearray() As Byte)
    Dim Palette4() As RGBTriple, i As Integer
    ReDim Palette4(15)
    CopyMemory Palette4(0), bytearray(0), UBound(bytearray) + 1
    ' Create a color table
    For i = 0 To 15
        bm4.bmiColors(i).rgbRed = Palette4(i).Red
        bm4.bmiColors(i).rgbGreen = Palette4(i).Green
        bm4.bmiColors(i).rgbBlue = Palette4(i).Blue
        bm4.bmiColors(i).rgbReserved = 0
    Next i
End Sub

Public Sub CreateBitmap_1(bytearray() As Byte, BMPWidth As Long, BMPHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)
    ' Create a 1bit Bitmap
    Dim hDC As Long
    With bm1.bmiHeader
        .biSize = Len(bm1.bmiHeader)
        .biWidth = BMPWidth
        If Orientation = 0 Then
            .biHeight = BMPHeight                    'Bitmap Height, bitmap is top down.
        Else
            .biHeight = -BMPHeight
        End If
        .biPlanes = 1
        .biBitCount = 1
        .biCompression = BI_RGB
        .biSizeImage = 0
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = Colorused
        .biClrImportant = 0
    End With
    ' Get the DC.
    hDC = GetDC(0)
    hBmp = CreateDIBitmap_1(hDC, bm1.bmiHeader, CBM_INIT, bytearray(0), bm1, DIB_RGB_COLORS)
End Sub
Public Sub CreateBitmap_2(bytearray() As Byte, BMPWidth As Long, BMPHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)
' Create a 2bit Bitmap
    Dim hDC As Long
    With bm1.bmiHeader
        .biSize = Len(bm1.bmiHeader)
        .biWidth = BMPWidth
        If Orientation = 0 Then
            .biHeight = BMPHeight                    'Bitmap Height, bitmap is top down.
        Else
            .biHeight = -BMPHeight
        End If
        .biPlanes = 1
        .biBitCount = 2
        .biCompression = BI_RGB
        .biSizeImage = 0
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = Colorused
        .biClrImportant = 0
    End With
    ' Get the DC.
    hDC = GetDC(0)
    hBmp = CreateDIBitmap_2(hDC, bm2.bmiHeader, CBM_INIT, bytearray(0), bm2, DIB_RGB_COLORS)
End Sub

Public Sub CreateBitmap_4(bytearray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)
    ' Creates a device independent bitmap
    ' from the pixel data in Data().
    Dim hDC As Long
    With bm4.bmiHeader
        .biSize = Len(bm1.bmiHeader)
        .biWidth = PicWidth
        If Orientation = 0 Then
            .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
            .biHeight = -PicHeight
        End If
        .biPlanes = 1
        .biBitCount = 4
        .biCompression = BI_RGB
        .biSizeImage = 0
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = Colorused
        .biClrImportant = 0
    End With
    ' Get the DC.
    hDC = GetDC(0)
    hBmp = CreateDIBitmap_4(hDC, bm4.bmiHeader, CBM_INIT, bytearray(0), bm4, DIB_RGB_COLORS)
End Sub

Public Sub CreateBitmap_8(BitmapArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer, Optional Colorused As Long = 0)
    ' Creates a device independent bitmap
    ' from the pixel data in BitmapArry().
    Dim hDC As Long
    With bm8.bmiHeader
        .biSize = Len(bm8.bmiHeader)
        .biWidth = PicWidth
        If Orientation = 0 Then
            .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
            .biHeight = -PicHeight
        End If
        .biPlanes = 1
        .biBitCount = 8
        .biCompression = BI_RGB
        .biSizeImage = 0
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = Colorused
        .biClrImportant = 0
    End With
    ' Get the DC.
    hDC = GetDC(0)
    hBmp = CreateDIBitmap_8(hDC, bm8.bmiHeader, CBM_INIT, BitmapArray(0), bm8, DIB_RGB_COLORS)
End Sub

Public Sub DrawBitmap(PicWidth As Long, PicHeight As Long, PicObject As Object, Scalierung As Boolean, Optional X As Long = 0, Optional Y As Long = 0, Optional DrawToBG As Boolean = False)
    Dim cDC As Long, a As Long, b As Long, Übergabe As ScTw, realheight As Long, realwidth As Long
    If TypeOf PicObject Is Form Then
        'change ScaleMode direct
    Else
        PicObject.Cls
        b = PicObject.Parent.ScaleMode
        PicObject.Parent.ScaleMode = 1
    End If
    a = PicObject.ScaleMode
    PicObject.ScaleMode = 1
    If Scalierung Then
        Übergabe = PixelToTwips(PicWidth, PicHeight)
        If DrawToBG = False Then
            PicObject.Height = Übergabe.Height
            PicObject.Width = Übergabe.Width
        End If
    End If
    If DrawToBG = False Then
        If PicObject.Height <> PicObject.ScaleHeight Then 'with Boarders
            Übergabe = Twipstopixel(PicObject.Width, PicObject.Height)
            realheight = Übergabe.Height
            realwidth = Übergabe.Width
            PicObject.Height = PicObject.Height + (PicObject.Height - PicObject.ScaleHeight)
            PicObject.Width = PicObject.Width + (PicObject.Width - PicObject.ScaleWidth)
        Else
            PicObject.ScaleMode = 3
            realheight = PicObject.ScaleHeight
            realwidth = PicObject.ScaleWidth
        End If
    Else
        realheight = Übergabe.Height
        realwidth = Übergabe.Width
        PicHeight = realheight
        PicWidth = realwidth
    End If
    If hBmp Then
        cDC = CreateCompatibleDC(PicObject.hDC)
        SelectObject cDC, hBmp
        Call StretchBlt(PicObject.hDC, X, Y, realwidth, realheight, cDC, 0, 0, PicWidth, PicHeight, SRCCOPY)
        DeleteDC cDC
        DeleteObject hBmp
        hBmp = 0
    End If
    If TypeOf PicObject Is Form Then
        'change ScaleMode direct
    Else
        PicObject.Parent.ScaleMode = b
    End If
    PicObject.ScaleMode = a
    PicObject.Picture = PicObject.Image
End Sub

Public Sub CreateBitmap_24(bytearray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer, Optional ThreeToOrToFour As Integer = 0)
    ' Creates a device independent bitmap
    ' from the pixel data in BitmapArray().
    Dim hDC As Long, Bits() As RGBQUAD, BitsA() As RGBTriple
    Select Case ThreeToOrToFour
        Case 0
        ReDim Bits((UBound(bytearray) / 4) - 1)
        CopyMemory Bits(0), bytearray(0), UBound(bytearray)
            With bm24.bmiHeader
                .biSize = Len(bm24.bmiHeader)        'SizeOf Struct
                .biWidth = PicWidth        'Bitmap Width
                If Orientation = 0 Then
                    .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
                Else
                    .biHeight = -PicHeight
                End If
                .biBitCount = 32                        '32 bit alignment
                .biPlanes = 1                           'Single plane
                .biCompression = BI_RGB                 'No Compression
                .biSizeImage = 0                        'Default
                .biXPelsPerMeter = 0                    'Default
                .biYPelsPerMeter = 0                    'Default
                .biClrUsed = 0                          'Default
                .biClrImportant = 0                     'Default
            End With
        Case 1
        ReDim BitsA((UBound(bytearray) / 3) - 1)
        CopyMemory BitsA(0), bytearray(0), UBound(bytearray)
        With bm24a.bmiHeader
            .biSize = Len(bm24.bmiHeader)        'SizeOf Struct
            .biWidth = PicWidth        'Bitmap Width
            If Orientation = 0 Then
                .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
            Else
                .biHeight = -PicHeight
            End If
            .biBitCount = 24                        '24 bit alignment
            .biPlanes = 1                           'Single plane
            .biCompression = BI_RGB                 'No Compression
            .biSizeImage = 0                        'Default
            .biXPelsPerMeter = 0                    'Default
            .biYPelsPerMeter = 0                    'Default
            .biClrUsed = 0                          'Default
            .biClrImportant = 0                     'Default
        End With
    End Select
    ' Get the DC.
    hDC = GetDC(0)
    Select Case ThreeToOrToFour
        Case 0
            hBmp = CreateDIBitmap_24(hDC, bm24.bmiHeader, CBM_INIT, Bits(0), bm24, DIB_RGB_COLORS)
        Case 1
            hBmp = CreateDIBitmap_24a(hDC, bm24a.bmiHeader, CBM_INIT, BitsA(0), bm24a, DIB_RGB_COLORS)
    End Select
End Sub
Public Sub CreateBitmap_16(bytearray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer)
    ' Creates a device independent bitmap
    ' from the pixel data in BitmapArray().
    Dim hDC As Long
    With bm16.bmiHeader
        .biSize = Len(bm16.bmiHeader)        'SizeOf Struct
        .biWidth = PicWidth                       'Bitmap Width
        If Orientation = 0 Then
            .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
            .biHeight = -PicHeight
        End If
        .biPlanes = 1                           'Single plane
        .biBitCount = 16                        '32 bit alignment
        .biCompression = BI_RGB                 'No Compression
        .biSizeImage = 0                        'Default
        .biXPelsPerMeter = 0                    'Default
        .biYPelsPerMeter = 0                    'Default
        .biClrUsed = 0                          'Default
        .biClrImportant = 0                     'Default
    End With
    ' Get the DC.
    hDC = GetDC(0)
    hBmp = CreateDIBitmap_16(hDC, bm16.bmiHeader, CBM_INIT, bytearray(0), bm16, DIB_RGB_COLORS)
End Sub

Private Function PixelToTwips(xwert As Long, ywert As Long) As ScTw
    Dim ux As Long, uy As Long, XWert1 As Long, yWert1 As Long
    ux = Screen.TwipsPerPixelX
    PixelToTwips.Width = xwert * ux
    uy = Screen.TwipsPerPixelY
    PixelToTwips.Height = ywert * uy
End Function

Private Function Twipstopixel(xwert As Long, ywert As Long) As ScTw
    Dim ux As Long, uy As Long, XWert1 As Long, yWert1 As Long
    ux = Screen.TwipsPerPixelX
    Twipstopixel.Width = xwert / ux
    uy = Screen.TwipsPerPixelY
    Twipstopixel.Height = ywert / uy
End Function

Public Function InitColorTable_Grey(BitDepth As Integer, Optional To8Bit As Boolean = False) As Byte()
    Dim CurLevel As Integer, Übergabe() As Byte, n As Long, LevelDiff As Byte, Tbl() As RGBQUAD, Table3() As RGBTriple
    Erase bm8.bmiColors
    If BitDepth <> 16 Then
        ReDim Tbl(2 ^ BitDepth - 1)
        ReDim Table3(2 ^ BitDepth - 1)
    Else
        ReDim Tbl(255)
        ReDim Table3(255)
    End If
    LevelDiff = 255 / UBound(Tbl)
    For n = 0 To UBound(Tbl)
        With Tbl(n)
            .rgbRed = CurLevel
            .rgbGreen = CurLevel
            .rgbBlue = CurLevel
        End With
        With Table3(n)
            .Red = CurLevel
            .Green = CurLevel
            .Blue = CurLevel
        End With
        CurLevel = CurLevel + LevelDiff
    Next n
    Select Case BitDepth
        Case 1
            If To8Bit = True Then CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 8
        Case 2
            CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 16
        Case 4
            If To8Bit Then
                CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 64
            Else
                CopyMemory ByVal VarPtr(bm4.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 64
            End If
        Case 8
            CopyMemory ByVal VarPtr(bm8.bmiColors(0).rgbBlue), ByVal VarPtr(Tbl(0).rgbBlue), 1024
    End Select
    ReDim Übergabe(((UBound(Table3) + 1) * 3) - 1)
    CopyMemory Übergabe(0), ByVal VarPtr(Table3(0).Red), ((UBound(Table3) + 1) * 3)
    InitColorTable_Grey = Übergabe
End Function


