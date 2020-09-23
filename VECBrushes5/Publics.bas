Attribute VB_Name = "Publics"
' Publics.bas

Option Explicit

'============================================
Public PaintColor As Long
Public aPMRed As Boolean      ' +/- Red
Public LCCount As Long        ' Left button press count
Public aDRAW As Boolean
Public Xpic As Single   ' = picCanvas X
Public Ypic As Single   ' = picCanvas Y
Public picMem() As Long

Public aFast As Boolean
Public FillCul As Long
Public BackCul As Long


' X fading parameters
Public bPaintRed As Byte
Public bPaintGreen As Byte
Public bPaintBlue As Byte
Public CanvasColor As Long
Public bCanvasRed As Byte
Public bCanvasGreen As Byte
Public bCanvasBlue As Byte
Public bR As Byte
Public bG As Byte
Public bB As Byte
Public zalpha As Single

' Stroke tracking
Public bBack() As Byte

' Brushes
Public BrushSize As Integer
Public BrushAngle As Integer

' Saved values
Public svPaintColor As Long
Public svzalpha As Single
Public svBrushSize As Integer
Public svBrushAngle As Integer
Public svBrushSizeAtStretch As Integer

' VEC Delete last, Undo, Redo
Public NumVectors As Integer
Public NumVisVectors As Integer
Public VEC() As Integer
Public VECSIZE As Long
Public aRUNVEC As Boolean
Public aStart As Boolean

' VEC array stretching
Public VECORG() As Integer
Public aDupVECArray  As Boolean
Public aStretch As Boolean

' Drawing parameter
Public Xprev As Single, Yprev As Single
Public aHairs As Boolean

' Parked colors
Public NumParkedColors As Long
Public ParkedColors() As Long

' General frame position & picResizer position
Public fraX As Single, fraY As Single

' Twips/pixel
Public STX As Long, STY As Long

' Cmd Vec Buttons Undo etc
Public aHiLite As Boolean

' Files
Public aFileOps As Boolean
Public PathSpec$, CurrentPath$

' Magnifier
Public aMagON As Boolean

' Help
Public aHelp As Boolean

' General publics
Public response As Long
Public Cul As Long
Public i As Long
Public j As Long
Public k As Long

Public Sub START_BRUSH(PIC As PictureBox, ByVal X As Single, ByVal Y As Single)
   Xprev = X
   Yprev = Y
   ReDim bBack(PIC.Width, PIC.Height)  ' Stroke copy to = 0
   If Not aRUNVEC Then
      CanvasColor = PIC.Point(X, Y)
   Else    'DrawVectors being used
      If aFast Then
         CanvasColor = picMem(X + 1, Y + 1)
      Else
         CanvasColor = PIC.Point(X, Y)
      End If
   End If
   bCanvasRed = CanvasColor And &HFF&
   bCanvasGreen = (CanvasColor And &HFF00&) / &H100&
   bCanvasBlue = (CanvasColor And &HFF0000) / &H10000
   bR = zalpha * (1& * bPaintRed - bCanvasRed) + bCanvasRed
   bG = zalpha * (1& * bPaintGreen - bCanvasGreen) + bCanvasGreen
   bB = zalpha * (1& * bPaintBlue - bCanvasBlue) + bCanvasBlue
   If Not aRUNVEC Then
      PIC.PSet (X, Y), RGB(bR, bG, bB)
   Else  'DrawVectors being used
      If aFast Then
         picMem(X + 1, Y + 1) = RGB(bB, bG, bR)
      Else
         PIC.PSet (X, Y), RGB(bR, bG, bB)
      End If
   End If
End Sub

Public Sub DO_SLOW_BRUSH(PIC As PictureBox, ByVal X As Single, ByVal Y As Single)
' Called from picCanvas_MouseMove, picScratch(Index)_MouseMove & DrawVectors
Dim idx As Integer
Dim idy As Integer
Dim ilines As Integer
Dim zstepx As Single
Dim zstepy As Single
Dim xx As Single
Dim yy As Single
Dim rs As Integer
Dim L As Long
   
   idx = X - Xprev
   idy = Y - Yprev
   ilines = Abs(idx)
   If ilines < Abs(idy) Then ilines = Abs(idy)
   
   If ilines <> 0 Then
      zstepx = idx / ilines
      zstepy = idy / ilines
      xx = Xprev
      yy = Yprev
      
      rs = BrushSize \ 2
      
      For L = 0 To 2 * ilines
         
         If BrushSize = 1 Then
            Slow_BresLine PIC, xx, yy, xx, yy
         Else  ' BrushSize > 1
            Select Case BrushAngle      ' Depends on Brush angle
            Case 0      ' |
               Slow_BresLine PIC, xx, yy - rs, xx, yy + rs
               xx = xx + 0.5
               Slow_BresLine PIC, xx, yy - rs, xx, yy + rs
               xx = xx - 0.5
               yy = yy + 0.5
               Slow_BresLine PIC, xx, yy - rs, xx, yy + rs
               yy = yy - 0.5
            Case 1      ' \
               Slow_BresLine PIC, xx - rs, yy - rs, xx + rs, yy + rs
               xx = xx + 0.5
               Slow_BresLine PIC, xx - rs, yy - rs, xx + rs, yy + rs
               xx = xx - 0.5
               yy = yy + 0.5
               Slow_BresLine PIC, xx - rs, yy - rs, xx + rs, yy + rs
               yy = yy - 0.5
            Case 2      ' /
               Slow_BresLine PIC, xx - rs, yy + rs, xx + rs, yy - rs
               xx = xx + 0.5
               Slow_BresLine PIC, xx - rs, yy + rs, xx + rs, yy - rs
               xx = xx - 0.5
               yy = yy + 0.5
               Slow_BresLine PIC, xx - rs, yy + rs, xx + rs, yy - rs
               yy = yy - 0.5
            Case 3      ' --
               Slow_BresLine PIC, xx - rs, yy, xx + rs, yy
               xx = xx + 0.5
               Slow_BresLine PIC, xx - rs, yy, xx + rs, yy
               xx = xx - 0.5
               yy = yy + 0.5
               Slow_BresLine PIC, xx - rs, yy, xx + rs, yy
               yy = yy - 0.5
            End Select
         End If   ' If BrushSize = 1 Then Else
         
         xx = xx + zstepx / 2
         yy = yy + zstepy / 2
      
      Next L   ' For L = 0 To ilines
      
      Xprev = xx - zstepx / 2
      Yprev = yy - zstepy / 2
   
   End If   ' If ilines <> 0 Then
End Sub

Public Sub DO_FAST_BRUSH(ByVal X As Single, ByVal Y As Single)
' Called from DrawVectors
Dim idx As Integer
Dim idy As Integer
Dim ilines As Integer
Dim zstepx As Single
Dim zstepy As Single
Dim xx As Single
Dim yy As Single
Dim rs As Integer
Dim L As Long
   
   idx = X - Xprev
   idy = Y - Yprev
   ilines = Abs(idx)
   If ilines < Abs(idy) Then ilines = Abs(idy)
   
   If ilines <> 0 Then
      zstepx = idx / ilines
      zstepy = idy / ilines
      xx = Xprev
      yy = Yprev
      
      rs = BrushSize \ 2
      
      For L = 0 To 2 * ilines
         
         If BrushSize = 1 Then
            Fast_BresLine xx, yy, xx, yy
         Else  ' BrushSize > 1
            Select Case BrushAngle      ' Depends on Brush angle
            Case 0      ' |
               Fast_BresLine xx, yy - rs, xx, yy + rs
               xx = xx + 0.5
               Fast_BresLine xx, yy - rs, xx, yy + rs
               xx = xx - 0.5
               yy = yy + 0.5
               Fast_BresLine xx, yy - rs, xx, yy + rs
               yy = yy - 0.5
            Case 1      ' \
               Fast_BresLine xx - rs, yy - rs, xx + rs, yy + rs
               xx = xx + 0.5
               Fast_BresLine xx - rs, yy - rs, xx + rs, yy + rs
               xx = xx - 0.5
               yy = yy + 0.5
               Fast_BresLine xx - rs, yy - rs, xx + rs, yy + rs
               yy = yy - 0.5
            Case 2      ' /
               Fast_BresLine xx - rs, yy + rs, xx + rs, yy - rs
               xx = xx + 0.5
               Fast_BresLine xx - rs, yy + rs, xx + rs, yy - rs
               xx = xx - 0.5
               yy = yy + 0.5
               Fast_BresLine xx - rs, yy + rs, xx + rs, yy - rs
               yy = yy - 0.5
            Case 3      ' --
               Fast_BresLine xx - rs, yy, xx + rs, yy
               xx = xx + 0.5
               Fast_BresLine xx - rs, yy, xx + rs, yy
               xx = xx - 0.5
               yy = yy + 0.5
               Fast_BresLine xx - rs, yy, xx + rs, yy
               yy = yy - 0.5
            End Select
         End If   ' If BrushSize = 1 Then Else
         
         xx = xx + zstepx / 2
         yy = yy + zstepy / 2
      
      Next L   ' For L = 0 To ilines
      
      Xprev = xx - zstepx / 2
      Yprev = yy - zstepy / 2
   
   End If   ' If ilines <> 0 Then
End Sub


Public Sub Slow_BresLine(PIC As PictureBox, ByVal ix1 As Long, ByVal iy1 As Long, _
   ByVal ix2 As Long, ByVal iy2 As Long)
Dim idx As Integer, idy As Integer
Dim jkstep As Integer
Dim incx As Integer
Dim id As Integer
Dim IX As Long, IY As Long
Dim ainc As Integer, binc As Integer

' Test
'Dim s1 As Long, s2 As Long

   idx = Abs(ix2 - ix1 + 0.5)
   idy = Abs(iy2 - iy1 + 0.5)
   jkstep = 1
   incx = 1
   If idx < idy Then   '-- Steep slope
      If iy1 > iy2 Then jkstep = -1
      If ix2 < ix1 Then incx = -1
      id = 2 * idx - idy
      ainc = 2 * (idx - idy)   '-ve
      binc = 2 * idx
      j = iy1: k = iy2: IX = ix1
   Else                '-- Shallow slope
      If ix1 > ix2 Then jkstep = -1
      If iy2 < iy1 Then incx = -1
      id = 2 * idy - idx
      ainc = 2 * (idy - idx)   '-ve
      binc = 2 * idy
      j = ix1: k = ix2: IX = iy1
   End If
   
   For IY = j To k Step jkstep
      
      If idx < idy Then   '-- Steep slope
         If IX >= 0 Then
         If IX <= PIC.Width - 1 Then
         If IY >= 0 Then
         If IY <= PIC.Height - 1 Then
         
            If bBack(IX, IY) = 0 Then
               CanvasColor = PIC.Point(IX, IY)
               bCanvasRed = CanvasColor And &HFF&
               bCanvasGreen = (CanvasColor And &HFF00&) / &H100&
               bCanvasBlue = (CanvasColor And &HFF0000) / &H10000
               bR = zalpha * (1& * bPaintRed - bCanvasRed) + bCanvasRed
               bG = zalpha * (1& * bPaintGreen - bCanvasGreen) + bCanvasGreen
               bB = zalpha * (1& * bPaintBlue - bCanvasBlue) + bCanvasBlue
               SetPixelV PIC.hdc, IX, IY, RGB(bR, bG, bB)
               bBack(IX, IY) = 1 ' Copy stroke
            End If
      
         End If
         End If
         End If
         End If
      
      Else                '-- Shallow slope
         If IY >= 0 Then
         If IY <= PIC.Width - 1 Then
         If IX >= 0 Then
         If IX <= PIC.Height - 1 Then
         
            If bBack(IY, IX) = 0 Then
               CanvasColor = PIC.Point(IY, IX)
               bCanvasRed = CanvasColor And &HFF&
               bCanvasGreen = (CanvasColor And &HFF00&) / &H100&
               bCanvasBlue = (CanvasColor And &HFF0000) / &H10000
               bR = zalpha * (1& * bPaintRed - bCanvasRed) + bCanvasRed
               bG = zalpha * (1& * bPaintGreen - bCanvasGreen) + bCanvasGreen
               bB = zalpha * (1& * bPaintBlue - bCanvasBlue) + bCanvasBlue
               SetPixelV PIC.hdc, IY, IX, RGB(bR, bG, bB)
               bBack(IY, IX) = 1 ' Copy stroke
            End If
      
         End If
         End If
         End If
         End If
         
      End If
      
      If id > 0 Then
          id = id + ainc
          IX = IX + incx
      Else
          id = id + binc
      End If
   
   Next IY

   End Sub

Public Sub Fast_BresLine(ByVal ix1 As Long, ByVal iy1 As Long, _
   ByVal ix2 As Long, ByVal iy2 As Long)
Dim idx As Integer, idy As Integer
Dim jkstep As Integer
Dim incx As Integer
Dim id As Integer
Dim IX As Long, IY As Long
Dim ainc As Integer, binc As Integer

' Test
'Dim s1 As Long, s2 As Long

   idx = Abs(ix2 - ix1 + 0.5)
   idy = Abs(iy2 - iy1 + 0.5)
   jkstep = 1
   incx = 1
   If idx < idy Then   '-- Steep slope
      If iy1 > iy2 Then jkstep = -1
      If ix2 < ix1 Then incx = -1
      id = 2 * idx - idy
      ainc = 2 * (idx - idy)   '-ve
      binc = 2 * idx
      j = iy1: k = iy2: IX = ix1
   Else                '-- Shallow slope
      If ix1 > ix2 Then jkstep = -1
      If iy2 < iy1 Then incx = -1
      id = 2 * idy - idx
      ainc = 2 * (idy - idx)   '-ve
      binc = 2 * idy
      j = ix1: k = ix2: IX = iy1
   End If
   
   For IY = j To k Step jkstep
      
      If idx < idy Then   '-- Steep slope
         If IX >= 0 Then
         If IX <= UBound(picMem, 1) - 1 Then
         If IY >= 0 Then
         If IY <= UBound(picMem, 2) - 1 Then
         
            If bBack(IX, IY) = 0 Then
               CanvasColor = picMem(IX + 1, IY + 1)
               bCanvasRed = CanvasColor And &HFF&
               bCanvasGreen = (CanvasColor And &HFF00&) / &H100&
               bCanvasBlue = (CanvasColor And &HFF0000) / &H10000
               bR = zalpha * (1& * bPaintRed - bCanvasRed) + bCanvasRed
               bG = zalpha * (1& * bPaintGreen - bCanvasGreen) + bCanvasGreen
               bB = zalpha * (1& * bPaintBlue - bCanvasBlue) + bCanvasBlue
               picMem(IX + 1, IY + 1) = RGB(bB, bG, bR)
               bBack(IX, IY) = 1 ' Copy stroke
            End If
      
         End If
         End If
         End If
         End If
      
      Else                '-- Shallow slope
         If IY >= 0 Then
         If IY <= UBound(picMem, 1) - 1 Then
         If IX >= 0 Then
         If IX <= UBound(picMem, 2) - 1 Then
         
            If bBack(IY, IX) = 0 Then
               CanvasColor = picMem(IY + 1, IX + 1)
               bCanvasRed = CanvasColor And &HFF&
               bCanvasGreen = (CanvasColor And &HFF00&) / &H100&
               bCanvasBlue = (CanvasColor And &HFF0000) / &H10000
               bR = zalpha * (1& * bPaintRed - bCanvasRed) + bCanvasRed
               bG = zalpha * (1& * bPaintGreen - bCanvasGreen) + bCanvasGreen
               bB = zalpha * (1& * bPaintBlue - bCanvasBlue) + bCanvasBlue
               picMem(IY + 1, IX + 1) = RGB(bB, bG, bR)
               bBack(IY, IX) = 1 ' Copy stroke
            End If
      
         End If
         End If
         End If
         End If
         
      End If
      
      If id > 0 Then
          id = id + ainc
          IX = IX + incx
      Else
          id = id + binc
      End If
   
   Next IY
   End Sub


'### GENERAL FRAME MOVER #####################################
Public Sub fraMOVER(frm As Form, fra As Frame, Button As Integer, X As Single, Y As Single)
Dim fraLeft As Long
Dim fraTop As Long

   If Button = vbLeftButton Then
      
      fraLeft = fra.Left + (X - fraX) \ STX
      If fraLeft < 0 Then fraLeft = 0
      If fraLeft + fra.Width > frm.Width \ STX + fra.Width \ 2 Then
         fraLeft = frm.Width \ STX - fra.Width \ 2
      End If
      fra.Left = fraLeft
      
      fraTop = fra.Top + (Y - fraY) \ STY
      If fraTop < 8 Then fraTop = 8
      If fraTop + fra.Height > frm.Height \ STY + fra.Height \ 2 Then
         fraTop = frm.Height \ STY - fra.Height \ 2
      End If
      fra.Top = fraTop
      
   End If
End Sub
'### END GENERAL FRAME MOVER #####################################

Public Sub FixScrollbars(picC As PictureBox, picP As PictureBox, HS As HScrollBar, VS As VScrollBar)
   ' picC = Container = picFrame
   ' picP = Picture   = picDisplay
      HS.Max = picP.Width - picC.Width + 12   ' +4 to allow for border
      VS.Max = picP.Height - picC.Height + 12 ' +4 to allow for border
      HS.LargeChange = picC.Width \ 10
      HS.SmallChange = 1
      VS.LargeChange = picC.Height \ 10
      VS.SmallChange = 1
      HS.Top = picC.Top + picC.Height + 1
      HS.Left = picC.Left
      HS.Width = picC.Width
      If picP.Width < picC.Width Then
         HS.Visible = False
         'HS.Enabled = False
      Else
         HS.Visible = True
         'HS.Enabled = True
      End If
      VS.Top = picC.Top
      VS.Left = picC.Left - VS.Width - 1
      VS.Height = picC.Height
      If picP.Height < picC.Height Then
         VS.Visible = False
         'VS.Enabled = False
      Else
         VS.Visible = True
         'VS.Enabled = True
      End If
End Sub

Public Sub FixExtension(FSpec$, Ext$)
Dim p As Long

If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   p = InStr(1, FSpec$, ".")
   
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      'Ext$ = LCase$(Mid$(FSpec$, p))
      If LCase$(Mid$(FSpec$, p)) <> Ext$ Then FSpec$ = Mid$(FSpec$, 1, p) & Ext$
   End If

End Sub

Public Sub READ_VEC_FILE(FSpec$)
Dim W As Long
   
   VECSIZE = 0
   ReDim VEC(0)
   Open FSpec$ For Input As #1
   Input #1, W    ' VEC(0)
   If W = 0 Then
      Close
      Exit Sub
   End If
   NumVectors = W
   NumVisVectors = NumVectors
   
   VECSIZE = 4
   ReDim VEC(0 To VECSIZE)
   VEC(0) = NumVectors
   Input #1, VEC(1)  ' picCanvas.Width
   Input #1, VEC(2)  ' picCanvas.Height
   Input #1, VEC(VECSIZE - 1), VEC(VECSIZE) ' -1, -1
   
   Do
      VECSIZE = VECSIZE + 6
      ReDim Preserve VEC(0 To VECSIZE)
      Input #1, VEC(VECSIZE - 5), VEC(VECSIZE - 4) ' BrushSize, BrushAngle
      Input #1, VEC(VECSIZE - 3), VEC(VECSIZE - 2) ' bPaintRed, bPaintGreen
      Input #1, VEC(VECSIZE - 1), VEC(VECSIZE)     ' bPaintBlue, zalpha*100
   
      Do
         VECSIZE = VECSIZE + 2
         ReDim Preserve VEC(0 To VECSIZE)
         Input #1, VEC(VECSIZE - 1), VEC(VECSIZE) ' 1st X,Y or -1,-1 or -2,-2
         If VEC(VECSIZE) = -1 Then
            Exit Do      ' Next stroke
         ElseIf VEC(VECSIZE) = -2 Then
            Exit Do  ' END
         End If
      Loop
      If VEC(VECSIZE) = -2 Then Exit Do
   Loop
   Close
End Sub

Public Sub SAVE_VEC_FILE(FSpec$)
Dim W As Long
   
   Open FSpec$ For Output As #1
   Print #1, Trim$(Str$(CInt(VEC(0))))
   Print #1, Trim$(Str$(CInt(VEC(1))))
   Print #1, Trim$(Str$(CInt(VEC(2))))
   W = 3
   Do
      If VEC(W) = -2 Then Exit Do  ' No strokes
      
      If VEC(W) = -1 Then  ' Start of stroke
         Print #1, Trim$(Str$(CInt(VEC(W)))); ","; Trim$(Str$(CInt(VEC(W + 1))))  ' -1, vecnum
         W = W + 2: Print #1, Trim$(Str$(CByte(VEC(W)))); ",";     ' BrushSize,
         W = W + 1: Print #1, Trim$(Str$(CByte(VEC(W))))           ' BrushAngle
         W = W + 1: Print #1, Trim$(Str$(CByte(VEC(W)))); ",";     ' bPaintRed,
         W = W + 1: Print #1, Trim$(Str$(CByte(VEC(W))))           ' bPaintGreen
         W = W + 1: Print #1, Trim$(Str$(CByte(VEC(W)))); ",";     ' bPaintBlue,
         W = W + 1: Print #1, Trim$(Str$(CByte(VEC(W))))           ' zalpha*100
         W = W + 1
         Do
            If VEC(W) = -1 Then
               Exit Do      ' Next stroke
            ElseIf VEC(W) = -2 Then
               Exit Do  ' END
            Else
               Print #1, Trim$(Str$(CInt(VEC(W)))); ","; Trim$(Str$(CInt(VEC(W + 1)))) ' X,Y
            End If
            W = W + 2
         Loop
         If VEC(W) = -2 Then Exit Do   ' END
      End If
   
   Loop
   Print #1, Trim$(Str$(CInt(VEC(W)))); ","; Trim$(Str$(CInt(VEC(W + 1)))) ' -2,-2
   Close
End Sub

Public Sub SAVE_VEB_FILE(FSpec$)
   If Dir$(FSpec$) <> "" Then Kill FSpec$
   Open FSpec$ For Binary As #1
   Put #1, , VEC()
   Close
End Sub

Public Sub READ_VEB_FILE(FSpec$)
   Open FSpec$ For Binary As #1
   VECSIZE = LOF(1) \ 2 - 1
   ReDim VEC(0 To VECSIZE)
   Get #1, , VEC()
   Close
   
   For i = 0 To VECSIZE
      If VEC(i) = -2 Then Exit For  ' i @ 1st -2
   Next i

   If i < VECSIZE + 1 Then
      ' Correct for any added zero on veb binary save
      VECSIZE = i + 1
      ReDim Preserve VEC(0 To VECSIZE)
   End If
   
   NumVectors = VEC(0)
   NumVisVectors = NumVectors
End Sub

Public Sub READ_PRK_FILE(FSpec$)
Dim v As Long
   Open FSpec$ For Input As #1
   Input #1, v
   If v = 0 Then
      Close
      MsgBox "PRK file corrupt", vbCritical, "Import prk file"
      Exit Sub
   End If
   NumParkedColors = v
   ReDim ParkedColors(0 To NumParkedColors)
   ParkedColors(0) = NumParkedColors
   For i = 1 To NumParkedColors Step 3
      Input #1, ParkedColors(i), ParkedColors(i + 1), ParkedColors(i + 2)
   Next i
   Close
End Sub

Public Sub SAVE_PRK_FILE(FSpec$)
   Open FSpec$ For Output As #1
   Print #1, ParkedColors(0)  ' NumParkedColors
   For i = 1 To NumParkedColors Step 3
      Print #1, Trim$(Str$(CInt(ParkedColors(i)))); ",";       ' X
      Print #1, Trim$(Str$(CInt(ParkedColors(i + 1)))); ",";   ' Y
      Print #1, ParkedColors(i + 2)    ' Parked Long Color
   Next i
   Close
End Sub

'### FILL #####################################################

Public Sub Fill(PIC As PictureBox, X As Single, Y As Single)
   ' Fill with FillColor = DrawColor at X,Y
   PIC.DrawStyle = vbSolid
   PIC.DrawMode = 13
   PIC.DrawWidth = 1
   PIC.FillColor = PaintColor
   PIC.FillStyle = vbFSSolid
   
   ' FLOODFILLSURFACE = 1
   ' Fills with FillColor so long as point surrounded by
   ' color = PIC.Point(X, Y)
   
   ExtFloodFill PIC.hdc, X, Y, PIC.Point(X, Y), 1    'FLOODFILLSURFACE
   
   PIC.FillStyle = vbFSTransparent  'Default (Transparent)
'   PIC.DrawWidth = TheDrawWidth
'   PIC.ForeColor = DrawColor
   PIC.Refresh
End Sub
'### END FILL #####################################################


'Public Sub SingleFill(IX As Integer, IY As Integer)
'' Replaces every BackCul with FillCul
'' for the whole picture - can only be
'' once else last Fill overwrites previous
'' fills
'
''Public FillCul As Long
''Public BackCul As Long
'
'Dim ixp As Integer, iyp As Integer
'Dim ixx As Integer, iyy As Integer
'Dim px As Integer, py As Integer
'
'Dim dz As Integer, sign As Integer, T As Integer
'Dim ddz As Integer, spi As Integer, n As Integer
'
'ixp = IX: iyp = IY
'
'   picMem(ixp, iyp) = FillCul
'   bBack(ixp, iyp) = 1
'
'   dz = 0: sign = 1
'   ixx = ixp: iyy = iyp
'   Do
'      T = 0
'      ddz = 1
'      spi = 1
'      GoSub SPIN
'      GoSub SPIN
'      If spi = 1 Then Exit Do
'   Loop
'Exit Sub
''=============
'SPIN:
'   dz = dz + ddz
'   sign = -sign
'   T = 1 - T
'   GoSub XYPLUSMINUS
'   T = 1 - T
'   GoSub XYPLUSMINUS
'Return
'
'XYPLUSMINUS:
'   For n = 1 To dz
'      If T = 0 Then
'         ixx = ixx + sign
'      Else
'         iyy = iyy + sign
'      End If
'      If ixx > 0 And ixx <= UBound(picMem, 1) Then
'      If iyy > 0 And iyy <= UBound(picMem, 2) Then
'         GoSub TestCross
'      End If
'      End If
'   Next n
'Return
'
'TestCross:
'   spi = 0
'   Cul = picMem(ixx, iyy)
'   If Cul = BackCul Then picMem(ixx, iyy) = FillCul Else Return
'
'   py = iyy - 1
'   If py > 0 And py <= UBound(picMem, 2) Then
'      If picMem(ixx, py) = BackCul Then picMem(ixx, py) = FillCul Else Return
'   End If
'   py = iyy + 1
'   If py > 0 And py <= UBound(picMem, 2) Then
'      If picMem(ixx, py) = BackCul Then picMem(ixx, py) = FillCul Else Return
'   End If
'
'   px = ixx - 1
'   If px > 0 And px <= UBound(picMem, 1) Then
'      If picMem(px, iyy) = BackCul Then picMem(px, iyy) = FillCul Else Return
'   End If
'   px = ixx + 1
'   If px > 0 And px <= UBound(picMem, 1) Then
'      If picMem(px, iyy) = BackCul Then picMem(px, iyy) = FillCul Else Return
'   End If
'Return
'End Sub
'
'Public Sub FastFillRecurs(IX As Integer, IY As Integer)
'' Recursive for very small areas only
'' else 'Out of Stack Space'
'' On Error Resume Next however just breaks out of the sub
'' giving a partial fill
'
'' Public FillCul As Long, BackCul As Long)
'' ReDim bBack(1 to picCanvas.Width, 1 to picCanvas.Height)
'On Error Resume Next
'
'Dim ii As Long
'Dim jj As Long
'
'   If IX > 0 And IX <= UBound(picMem, 1) Then
'   If IY > 0 And IY <= UBound(picMem, 2) Then
'
'      If picMem(IX, IY) = BackCul Then
'         picMem(IX, IY) = FillCul
'         bBack(IX, IY) = 1
'         For jj = -1 To 1
'         For ii = -1 To 1
'            If bBack(IX + ii, IY + jj) = 0 Then
'               FastFill IX + ii, IY + jj   ', FillCul, BackCul
'               bBack(IX + ii, IY + jj) = 1
'            End If
'         Next ii
'         Next jj
'
'      End If
'
'   End If
'   End If
'End Sub
'
