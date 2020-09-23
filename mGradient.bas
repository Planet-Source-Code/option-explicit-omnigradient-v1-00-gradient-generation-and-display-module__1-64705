Attribute VB_Name = "mGradient"
'*************************************************************************
'* OmniGradient v1.00 - Gradient generation and display module.          *
'* Compiled by Matthew R. Usner for Planet Source Code.                  *
'*************************************************************************
'* Almost all the code in this module was written by one of the VB gurus *
'* on PSC, Carles P.V.  His original submission can be found on PSC at   *
'* txtCodeId=60580.  In that submission he has two modules.  One is for  *
'* linear gradients that can be displayed at any angle, and the other is *
'* for circular gradients.  My meager contributions to his excellence:   *
'* - The linear and circular gradient generating routines have been      *
'*   merged into one routine that handles both gradient styles.          *
'* - Linear gradients can be generated in a "middle-out" fashion - that  *
'*   is, from Color1 to Color2 back to Color1.  Makes a nice 3D effect.  *
'* - The code that actually displays the gradient has been removed from  *
'*   Carles' gradient generating routines and placed in a separate       *
'*   procedure.  The reason for that is simple - speed.  If you need to  *
'*   refresh the background of a usercontrol often, why recalculate the  *
'*   background gradient information every time?  Calculate it ONCE,     *
'*   cache it, and redraw it as needed with a simple PaintGradient call. *
'*************************************************************************
'* Feedback is welcome but if you wish to vote, please vote not for this *
'* but for Carles' submission at txtCodeId=60580.  Thanks.               *
'*************************************************************************

Option Explicit

Public Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const PI             As Single = 3.14159265358979
Private Const TO_DEG         As Single = 180 / PI
Private Const TO_RAD         As Single = PI / 180
Private Const INT_ROT        As Long = 1000

'  API that actually paints the gradient.
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Sub CalcGradient(Width As Long, Height As Long, _
                        ByVal Color1 As Long, ByVal Color2 As Long, _
                        ByVal Angle As Single, ByVal bMOut As Boolean, _
                        ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long, _
                        Optional ByVal Circular As Boolean = False)

   Dim lGrad()   As Long, lGrad2() As Long

   Dim lClr      As Long
   Dim R1        As Long, G1 As Long, B1 As Long
   Dim R2        As Long, G2 As Long, B2 As Long
   Dim dR        As Long, dG As Long, dB As Long

   Dim Scan      As Long
   Dim i         As Long, j As Long, k As Long
   Dim jIn       As Long
   Dim iEnd      As Long, jEnd As Long
   Dim Offset    As Long

   Dim lQuad     As Long
   Dim AngleDiag As Single
   Dim AngleComp As Single

   Dim g         As Long
   Dim luSin     As Long, luCos As Long

   Dim Offset1   As Long, Offset2 As Long
   Dim iPad      As Long, jPad    As Long

   Dim ia        As Long, iaa     As Long
   Dim ja        As Long, jaa     As Long

   Dim s()       As Long ' squares sequence
   Dim sc        As Long ' squares sequence counter (sequence index -> root)

   If (Width > 0 And Height > 0) Then

      If Circular = False Then

'        when angle is >= 91 and <= 270, the colors
'        invert in MiddleOut mode.  This corrects that.
         If bMOut And Angle >= 91 And Angle <= 270 Then
            g = Color1
            Color1 = Color2
            Color2 = g
         End If

'        -- Right-hand [+] (ox=0º)
         Angle = -Angle + 90

'        -- Normalize to [0º;360º]
         Angle = Angle Mod 360
         If (Angle < 0) Then
            Angle = 360 + Angle
         End If

'        -- Get quadrant (0 - 3)
         lQuad = Angle \ 90

'        -- Normalize to [0º;90º]
           Angle = Angle Mod 90

'        -- Calc. gradient length ('distance')
         If (lQuad Mod 2 = 0) Then
            AngleDiag = Atn(Width / Height) * TO_DEG
         Else
            AngleDiag = Atn(Height / Width) * TO_DEG
         End If
         AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
         Angle = Angle * TO_RAD
         g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem

'        -- Decompose colors
         If (lQuad > 1) Then
            lClr = Color1
            Color1 = Color2
            Color2 = lClr
         End If
         R1 = (Color1 And &HFF&)
         G1 = (Color1 And &HFF00&) \ 256
         B1 = (Color1 And &HFF0000) \ 65536
         R2 = (Color2 And &HFF&)
         G2 = (Color2 And &HFF00&) \ 256
         B2 = (Color2 And &HFF0000) \ 65536

'        -- Get color distances
         dR = R2 - R1
         dG = G2 - G1
         dB = B2 - B1

'        -- Size gradient-colors array
         ReDim lGrad(0 To g - 1)
         ReDim lGrad2(0 To g - 1)

'        -- Calculate gradient-colors
         iEnd = g - 1
         If (iEnd = 0) Then
'           -- Special case (1-pixel wide gradient)
            lGrad2(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
         Else
            For i = 0 To iEnd
               lGrad2(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
            Next i
         End If

'        'if' block added by Matthew R. Usner - accounts for possible MiddleOut gradient draw.
         If bMOut Then
            k = 0
            For i = 0 To iEnd Step 2
               lGrad(k) = lGrad2(i)
               k = k + 1
            Next i
            For i = iEnd - 1 To 1 Step -2
               lGrad(k) = lGrad2(i)
               k = k + 1
            Next i
         Else
            For i = 0 To iEnd
               lGrad(i) = lGrad2(i)
            Next i
         End If

'        -- Size DIB array
         ReDim lBits(Width * Height - 1) As Long
         iEnd = Width - 1
         jEnd = Height - 1
         Scan = Width

'        -- Render gradient DIB
         Select Case lQuad

            Case 0, 2
               luSin = Sin(Angle) * INT_ROT
               luCos = Cos(Angle) * INT_ROT
               Offset = 0
               jIn = 0
               For j = 0 To jEnd
                  For i = 0 To iEnd
                     lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
                  Next i
                  jIn = jIn + luCos
                  Offset = Offset + Scan
               Next j

            Case 1, 3
               luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
               luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
               Offset = jEnd * Scan
               jIn = 0
               For j = 0 To jEnd
                  For i = 0 To iEnd
                     lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
                  Next i
                  jIn = jIn + luCos
                  Offset = Offset - Scan
               Next j

         End Select

'        -- Define DIB header
         With uBIH
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = Width
            .biHeight = Height
         End With

      Else

         '-- Calc. gradient length ('diagonal')
         g = Sqr(Width * Width + Height * Height) \ 2

         '-- Decompose colors
         R1 = (Color1 And &HFF&)
         G1 = (Color1 And &HFF00&) \ 256
         B1 = (Color1 And &HFF0000) \ 65536
         R2 = (Color2 And &HFF&)
         G2 = (Color2 And &HFF00&) \ 256
         B2 = (Color2 And &HFF0000) \ 65536

         '-- Get color distances
         dR = R2 - R1
         dG = G2 - G1
         dB = B2 - B1

         '-- Size gradient-colors array
         ReDim lGrad(0 To g)

         '-- Build squares sequence LUT
         ReDim s(0 To g)
         For i = 1 To g
            s(i) = s(i - 1) + i + i - 1
         Next i

         '-- Calculate gradient-colors
         If (g = 0) Then
            '-- Special case (1-pixel wide gradient)
            lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
         Else
            For i = 0 To g
               lGrad(i) = B1 + (dB * i) \ g + 256 * (G1 + (dG * i) \ g) + 65536 * (R1 + (dR * i) \ g)
            Next i
         End If

         '-- Size DIB array
         ReDim lBits(Width * Height - 1) As Long

         '== Render gradient DIB

         '-- First "quadrant"...

         Scan = Width
         iPad = Width Mod 2
         jPad = Height Mod 2

         iEnd = Scan \ 2 + iPad - 1
         jEnd = Height \ 2 + jPad - 1
         Offset1 = jEnd * Scan + Scan \ 2

         ja = 1
         jaa = -1
         For j = 0 To jEnd
            sc = j
            ja = ja + jaa
            jaa = jaa + 2
            ia = ja + 1
            iaa = -1
            For i = Offset1 To Offset1 + iEnd
               ia = ia + iaa
               iaa = iaa + 2
               lBits(i) = lGrad(sc)
               If (ia >= s(sc) - sc) Then
                  sc = sc + 1
               End If
            Next i
            Offset1 = Offset1 - Scan
         Next j

         '-- Mirror first "quadrant"

         iEnd = iEnd - iPad
         Offset1 = 0
         Offset2 = Scan - 1

         For j = 0 To jEnd
            For i = 0 To iEnd
               lBits(Offset1 + i) = lBits(Offset2 - i)
            Next i
            Offset1 = Offset1 + Scan
            Offset2 = Offset2 + Scan
         Next j

         '-- Mirror first "half"

         iEnd = Scan - 1
         jEnd = jEnd - jPad
         Offset1 = (Height - 1) * Scan
         Offset2 = 0

         For j = 0 To jEnd
            For i = 0 To iEnd
               lBits(Offset1 + i) = lBits(Offset2 + i)
            Next i
            Offset1 = Offset1 - Scan
            Offset2 = Offset2 + Scan
         Next j

         '-- Define DIB header
         With uBIH
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = Width
            .biHeight = Height
         End With

      End If

   End If

End Sub

Public Sub PaintGradient(DestDC As Long, _
                         ByVal XDest As Long, ByVal YDest As Long, _
                         ByVal DestWidth As Long, ByVal DestHeight As Long, _
                         ByVal XSource As Long, ByVal YSource As Long, _
                         ByVal SourceWidth As Long, ByVal SourceHeight As Long, _
                         ByRef lBits() As Long, ByRef uBIH As BITMAPINFOHEADER)

'*************************************************************************
'* displays the gradient to the destination DC.  This is almost the same *
'* as the StretchDIBits API call but I wanted to have the StretchDIBits  *
'* declare in this module to keep it as fully self-contained as possible.*
'* User only needs to declare lBits() color array and uBIH bitmap info   *
'* UDT variable in appropriate form or public module.                    *
'*************************************************************************

   Call StretchDIBits(DestDC, _
                      XDest, YDest, _
                      DestWidth, DestHeight, _
                      XSource, YSource, _
                      SourceWidth, SourceHeight, _
                      lBits(0), uBIH, _
                      DIB_RGB_COLORS, vbSrcCopy)

End Sub
