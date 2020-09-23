Attribute VB_Name = "VBFill"
'VBFill module by Msi, 2003

Private Type ErrorType
   ErrorX As Long
   ErrorY As Long
   Direction As Byte '0=Left,1=Right,'3=Vertical
   StopStage As Byte '0=before up,1=before down
   CurrentLeft As Long
   CurrentRight As Long
   MaxLeft As Long
   MaxRight As Long
End Type

Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const HORZRES = 8
Private Const VERTRES = 10

Dim FillhDC As Long, FillCol As Long, AreaColor As Long
Dim FillErrors() As ErrorType
Dim FillAccelTable() As Byte

Dim FillImage() As Long, GlobalFillWithImage As Byte, X As Long, Y As Long
Dim FImageWidth As Long, FImageHeight As Long

'This sub reads a PictureBox's contents into an array.
'This speeds up thing a lot, but creates a little pause
'at the start of the prog.
'It's half a sec with a 256*256 bitmap. Anyways, i
'don't think You want to fill with something bigger...
Public Sub ParseFillImage(PB As PictureBox)
FImageWidth = PB.Width
FImageHeight = PB.Height
ReDim FillImage(FImageWidth - 1, FImageHeight - 1)
For X = 0 To FImageWidth - 1
   For Y = 0 To FImageHeight - 1
      FillImage(X, Y) = PB.Point(X, Y)
   Next 'Y
Next 'X

End Sub

Public Sub PaintErrorPlaces(TargetForm As Form)
For X = 0 To UBound(FillErrors) - 1
   TargetForm.PSet (FillErrors(X).ErrorX, FillErrors(X).ErrorY), 2 ^ 24 - 1
Next X
TargetForm.Caption = UBound(FillErrors) + 1

End Sub

'Look for comments at FillLine().
'After that, You will easily understand this.

Public Sub VBFloodFill(hdc As Long, StartX As Long, StartY As Long, ImageFill As Byte, FloodFillColor As Long, FillAreaColor As Long)
Dim TempX As Long, TempY As Long, Z As Long, TempFE As ErrorType

FillhDC = hdc
GlobalFillWithImage = ImageFill
FillCol = FloodFillColor
AreaColor = FillAreaColor

ReDim FillAccelTable(-1 To GetDeviceCaps(FillhDC, HORZRES) + 1, -1 To GetDeviceCaps(FillhDC, VERTRES) + 1)

ReDim FillErrors(0)
FillErrors(0).ErrorX = -10000
FillErrors(0).ErrorY = -10000

FillLine StartX, StartY
TempFE = FillErrors(Z)
Do Until (TempFE.ErrorX = -10000) And (TempFE.ErrorY = -10000)
   If TempFE.Direction = 0 Then 'Left
      If TempFE.StopStage = 0 Then
         For X = TempFE.CurrentLeft To TempFE.MaxLeft
            FillLine TempFE.ErrorX - X, TempFE.ErrorY - 1
            FillLine TempFE.ErrorX - X, TempFE.ErrorY + 1
         Next 'x
      Else
         FillLine TempFE.ErrorX - TempFE.CurrentLeft, TempFE.ErrorY + 1
         For X = TempFE.CurrentLeft + 1 To TempFE.MaxLeft
            FillLine TempFE.ErrorX - X, TempFE.ErrorY - 1
            FillLine TempFE.ErrorX - X, TempFE.ErrorY + 1
         Next 'x
      End If
      
      For X = TempFE.CurrentRight To TempFE.MaxRight
         FillLine TempFE.ErrorX + X, TempFE.ErrorY - 1
         FillLine TempFE.ErrorX + X, TempFE.ErrorY + 1
      Next 'x
   ElseIf TempFE.Direction = 1 Then 'Right
      If TempFE.StopStage = 0 Then
         For X = TempFE.CurrentRight To TempFE.MaxRight
            FillLine TempFE.ErrorX + X, TempFE.ErrorY - 1
            FillLine TempFE.ErrorX + X, TempFE.ErrorY + 1
         Next 'x
      Else
         FillLine TempFE.ErrorX + TempFE.CurrentRight, TempFE.ErrorY + 1
         For X = TempFE.CurrentRight + 1 To TempFE.MaxRight
            FillLine TempFE.ErrorX + X, TempFE.ErrorY - 1
            FillLine TempFE.ErrorX + X, TempFE.ErrorY + 1
         Next 'x
      End If
   Else 'Vertical
      If TempFE.StopStage = 0 Then
         FillLine TempFE.ErrorX, TempFE.ErrorY - 1
         FillLine TempFE.ErrorX, TempFE.ErrorY + 1
      Else
         FillLine TempFE.ErrorX, TempFE.ErrorY + 1
      End If
   End If
   Z = Z + 1
   TempFE = FillErrors(Z)
Loop

Erase FillAccelTable ', FillErrors 'Commented out for debugging; Erase FillErrors for real use, too.

End Sub

Private Sub PutPixel(tohDC As Long, toX As Long, toY As Long, PixelColor As Long)
If GlobalFillWithImage = 0 Then
   SetPixelV FillhDC, toX, toY, FillCol
Else
   SetPixelV FillhDC, toX, toY, FillImage(toX Mod FImageWidth, toY Mod FImageHeight)
End If

End Sub

'I decided to keep the procedure clean of comments,
'because it's a little messy without comments, too... -:-)

'So, here's how this thing works:
'First, it puts a pixel to (StartX,StartY) if there
'is AreaColor.
'Then it starts to draw a line to the left and to the right,
'as long as it finds AreaColor in the new pixel's place.
'If a line cannot be drawn furter, it sets RightDead or
'LeftDead (obviously LeftDead is for the left line).
'After this, the program tries to call itself for every
'pixel below and above every pixel in the lines.
'I wrote "tries" 'cause sometimes it will bump into
'the recursive programs biggest obstacle:
'The Stack Overflow Error.
'This is a serious problem, but there is a solution:
'You can save the progress data of the sub,
'and when the "parent" sub and it's recursive "childs"
'finish, you recreate the previously erroneous situation
'(which was erroneous because the fullness of the stack)
'with a nearly empty stack. Then You repeat this until
'there are no other errors generated during run.
'The result: a perfectly working recursive algorithm,
'with a nearly infinite stack size...
'Lovely, isn't it? -:-))

Private Sub FillLine(StartX As Long, StartY As Long)
Dim rX As Long, lX As Long, K As Long, MadeThrough As Byte
Dim LeftDead As Byte, RightDead As Byte, LeftMax As Long, RightMax As Long

If (StartX < -1) Or (StartY < -1) Or (StartX > UBound(FillAccelTable, 1)) Or (StartY > UBound(FillAccelTable, 2)) Then Exit Sub

If FillAccelTable(StartX, StartY) = 0 Then
   If GetPixel(FillhDC, StartX, StartY) = AreaColor Then
      If GlobalFillWithImage = 0 Then
         SetPixelV FillhDC, StartX, StartY, FillCol
      Else
         SetPixelV FillhDC, StartX, StartY, FillImage(StartX Mod FImageWidth, StartY Mod FImageHeight)
      End If
      
      FillAccelTable(StartX, StartY) = 1

      rX = StartX: lX = StartX
      Do
         rX = rX + 1
         If (RightDead = 0) Then
            If (FillAccelTable(rX, StartY) = 0) Then
               If (GetPixel(FillhDC, rX, StartY) = AreaColor) Then
                  If GlobalFillWithImage = 0 Then
                     SetPixelV FillhDC, rX, StartY, FillCol
                  Else
                     SetPixelV FillhDC, rX, StartY, FillImage(rX Mod FImageWidth, StartY Mod FImageHeight)
                  End If

                  FillAccelTable(rX, StartY) = 1
                  RightMax = RightMax + 1
               Else
                  RightDead = 1
               End If
            End If
         End If

         lX = lX - 1
         If (LeftDead = 0) Then
            If (FillAccelTable(lX, StartY) = 0) Then
               If (GetPixel(FillhDC, lX, StartY) = AreaColor) Then
                  If GlobalFillWithImage = 0 Then
                     SetPixelV FillhDC, lX, StartY, FillCol
                  Else
                     SetPixelV FillhDC, lX, StartY, FillImage(lX Mod FImageWidth, StartY Mod FImageHeight)
                  End If
                  FillAccelTable(lX, StartY) = 1
                  LeftMax = LeftMax + 1
               Else
                  LeftDead = 1
               End If
            End If
         End If
      Loop Until (LeftDead = 1) And (RightDead = 1)

      If LeftMax > 0 Then
         On Error GoTo LeftM
         For K = 0 To LeftMax
            MadeThrough = 0
            If FillAccelTable(StartX - K, StartY - 1) = 0 Then
               FillLine StartX - K, StartY - 1
            End If
            MadeThrough = 1
            If FillAccelTable(StartX - K, StartY + 1) = 0 Then
               FillLine StartX - K, StartY + 1
            End If
            MadeThrough = 2
LeftM:
            If MadeThrough = 0 Then
               FillErrors(UBound(FillErrors)).ErrorX = StartX - K
               FillErrors(UBound(FillErrors)).ErrorY = StartY
               FillErrors(UBound(FillErrors)).Direction = 0
               FillErrors(UBound(FillErrors)).StopStage = 0
               FillErrors(UBound(FillErrors)).CurrentLeft = K
               FillErrors(UBound(FillErrors)).MaxLeft = LeftMax
               
               ReDim Preserve FillErrors(UBound(FillErrors) + 1)
               FillErrors(UBound(FillErrors)).ErrorX = StartX - K
               FillErrors(UBound(FillErrors)).ErrorY = StartY
               FillErrors(UBound(FillErrors)).Direction = 0
               FillErrors(UBound(FillErrors)).StopStage = 1
               FillErrors(UBound(FillErrors)).CurrentLeft = K
               FillErrors(UBound(FillErrors)).MaxLeft = LeftMax
               
               ReDim Preserve FillErrors(UBound(FillErrors) + 1)
               FillErrors(UBound(FillErrors)).ErrorX = -10000
               FillErrors(UBound(FillErrors)).ErrorY = -10000
               Exit Sub
            ElseIf MadeThrough = 1 Then
               FillErrors(UBound(FillErrors)).ErrorX = StartX - K
               FillErrors(UBound(FillErrors)).ErrorY = StartY
               FillErrors(UBound(FillErrors)).Direction = 0
               FillErrors(UBound(FillErrors)).StopStage = 1
               FillErrors(UBound(FillErrors)).CurrentLeft = K
               FillErrors(UBound(FillErrors)).MaxLeft = LeftMax
               
               ReDim Preserve FillErrors(UBound(FillErrors) + 1)
               FillErrors(UBound(FillErrors)).ErrorX = -10000
               FillErrors(UBound(FillErrors)).ErrorY = -10000
               Exit Sub
            End If
         Next 'K
      End If
      
      If RightMax > 0 Then
         On Error GoTo RightM
         For K = 0 To RightMax
            MadeThrough = 0
            If FillAccelTable(StartX + K, StartY - 1) = 0 Then
               FillLine StartX + K, StartY - 1
            End If
            MadeThrough = 1
            If FillAccelTable(StartX + K, StartY + 1) = 0 Then
               FillLine StartX + K, StartY + 1
            End If
            MadeThrough = 2
RightM:
            If MadeThrough = 0 Then
               FillErrors(UBound(FillErrors)).ErrorX = StartX + K
               FillErrors(UBound(FillErrors)).ErrorY = StartY
               FillErrors(UBound(FillErrors)).Direction = 1
               FillErrors(UBound(FillErrors)).StopStage = 0
               FillErrors(UBound(FillErrors)).CurrentRight = K
               FillErrors(UBound(FillErrors)).MaxRight = RightMax

               ReDim Preserve FillErrors(UBound(FillErrors) + 1)
               FillErrors(UBound(FillErrors)).ErrorX = StartX + K
               FillErrors(UBound(FillErrors)).ErrorY = StartY
               FillErrors(UBound(FillErrors)).Direction = 1
               FillErrors(UBound(FillErrors)).StopStage = 1
               FillErrors(UBound(FillErrors)).CurrentRight = K
               FillErrors(UBound(FillErrors)).MaxRight = RightMax

               ReDim Preserve FillErrors(UBound(FillErrors) + 1)
               FillErrors(UBound(FillErrors)).ErrorX = -10000
               FillErrors(UBound(FillErrors)).ErrorY = -10000
               Exit Sub
            ElseIf MadeThrough = 1 Then
               FillErrors(UBound(FillErrors)).ErrorX = StartX + K
               FillErrors(UBound(FillErrors)).ErrorY = StartY
               FillErrors(UBound(FillErrors)).Direction = 1
               FillErrors(UBound(FillErrors)).StopStage = 1
               FillErrors(UBound(FillErrors)).CurrentRight = K
               FillErrors(UBound(FillErrors)).MaxRight = RightMax

               ReDim Preserve FillErrors(UBound(FillErrors) + 1)
               FillErrors(UBound(FillErrors)).ErrorX = -10000
               FillErrors(UBound(FillErrors)).ErrorY = -10000
               Exit Sub
            End If
         Next 'K
      End If
      'Left-Right end
      
      If (LeftMax = 0) And (RightMax = 0) Then
         On Error GoTo VertM
      'We can only fill up and down this time... a 1-pixel wide vertical line.
      'This part isn't well optimized, but things like
      'a 1-pixel wide vertical line doesn't exists so
      'often in hand-drawn images, so i kept this as is.
         MadeThrough = 0
         FillLine StartX, StartY - 1
         MadeThrough = 1
         FillLine StartX, StartY + 1
         MadeThrough = 2
VertM:
         If MadeThrough = 0 Then
            FillErrors(UBound(FillErrors)).ErrorX = StartX
            FillErrors(UBound(FillErrors)).ErrorY = StartY
            FillErrors(UBound(FillErrors)).Direction = 3
            FillErrors(UBound(FillErrors)).StopStage = 0
            
            ReDim Preserve FillErrors(UBound(FillErrors) + 1)
            FillErrors(UBound(FillErrors)).ErrorX = StartX
            FillErrors(UBound(FillErrors)).ErrorY = StartY
            FillErrors(UBound(FillErrors)).Direction = 3
            FillErrors(UBound(FillErrors)).StopStage = 1
            
            ReDim Preserve FillErrors(UBound(FillErrors) + 1)
            FillErrors(UBound(FillErrors)).ErrorX = -10000
            FillErrors(UBound(FillErrors)).ErrorY = -10000
         ElseIf MadeThrough = 1 Then
            FillErrors(UBound(FillErrors)).ErrorX = StartX
            FillErrors(UBound(FillErrors)).ErrorY = StartY
            FillErrors(UBound(FillErrors)).Direction = 3
            FillErrors(UBound(FillErrors)).StopStage = 1

            ReDim Preserve FillErrors(UBound(FillErrors) + 1)
            FillErrors(UBound(FillErrors)).ErrorX = -10000
            FillErrors(UBound(FillErrors)).ErrorY = -10000
         End If
         
      End If
   End If
End If
End Sub
