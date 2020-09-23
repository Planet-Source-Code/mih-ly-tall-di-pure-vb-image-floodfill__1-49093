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

Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Const HORZRES = 8
Private Const VERTRES = 10

Dim FillhDC As Long
Dim FillErrors() As ErrorType
Dim FillAccelTable() As Byte

Public Sub PaintErrorPlaces(TargetForm As Form)
For x = 0 To UBound(FillErrors) - 1
   TargetForm.PSet (FillErrors(x).ErrorX, FillErrors(x).ErrorY), 2 ^ 24 - 1
Next x
TargetForm.Caption = UBound(FillErrors) + 1

End Sub

Public Sub VBFloodFill(hdc As Long, StartX As Long, StartY As Long, FillCol As Long, AreaColor As Long)
Dim TempX As Long, TempY As Long, Z As Long, TempFE As ErrorType
FillhDC = hdc

ReDim FillAccelTable(-1 To GetDeviceCaps(FillhDC, HORZRES) + 1, -1 To GetDeviceCaps(FillhDC, VERTRES) + 1)

ReDim FillErrors(0)
FillErrors(0).ErrorX = -10000
FillErrors(0).ErrorY = -10000

FillLine StartX, StartY, FillCol, AreaColor
TempFE = FillErrors(Z)
Do Until (TempFE.ErrorX = -10000) And (TempFE.ErrorY = -10000)
   If TempFE.Direction = 0 Then 'Left
      If TempFE.StopStage = 0 Then
         For x = TempFE.CurrentLeft To TempFE.MaxLeft
            FillLine TempFE.ErrorX - x, TempFE.ErrorY - 1, FillCol, AreaColor
            FillLine TempFE.ErrorX - x, TempFE.ErrorY + 1, FillCol, AreaColor
         Next 'x
      Else
         FillLine TempFE.ErrorX - TempFE.CurrentLeft, TempFE.ErrorY + 1, FillCol, AreaColor
         For x = TempFE.CurrentLeft + 1 To TempFE.MaxLeft
            FillLine TempFE.ErrorX - x, TempFE.ErrorY - 1, FillCol, AreaColor
            FillLine TempFE.ErrorX - x, TempFE.ErrorY + 1, FillCol, AreaColor
         Next 'x
      End If
      
      For x = TempFE.CurrentRight To TempFE.MaxRight
         FillLine TempFE.ErrorX + x, TempFE.ErrorY - 1, FillCol, AreaColor
         FillLine TempFE.ErrorX + x, TempFE.ErrorY + 1, FillCol, AreaColor
      Next 'x
   ElseIf TempFE.Direction = 1 Then 'Right
      If TempFE.StopStage = 0 Then
         For x = TempFE.CurrentRight To TempFE.MaxRight
            FillLine TempFE.ErrorX + x, TempFE.ErrorY - 1, FillCol, AreaColor
            FillLine TempFE.ErrorX + x, TempFE.ErrorY + 1, FillCol, AreaColor
         Next 'x
      Else
         FillLine TempFE.ErrorX + TempFE.CurrentRight, TempFE.ErrorY + 1, FillCol, AreaColor
         For x = TempFE.CurrentRight + 1 To TempFE.MaxRight
            FillLine TempFE.ErrorX + x, TempFE.ErrorY - 1, FillCol, AreaColor
            FillLine TempFE.ErrorX + x, TempFE.ErrorY + 1, FillCol, AreaColor
         Next 'x
      End If
   Else 'Vertical
      If TempFE.StopStage = 0 Then
         FillLine TempFE.ErrorX, TempFE.ErrorY - 1, FillCol, AreaColor
         FillLine TempFE.ErrorX, TempFE.ErrorY + 1, FillCol, AreaColor
      Else
         FillLine TempFE.ErrorX, TempFE.ErrorY + 1, FillCol, AreaColor
      End If
   End If
   Z = Z + 1
   TempFE = FillErrors(Z)
Loop

Erase FillAccelTable ', FillErrors 'Commented out for debugging

End Sub

Private Sub FillLine(StartX As Long, StartY As Long, FillCol As Long, AreaColor As Long)
Dim rX As Long, lX As Long, K As Long, MadeThrough As Byte
Dim LeftDead As Byte, RightDead As Byte, LeftMax As Long, RightMax As Long

If (StartX < -1) Or (StartY < -1) Or (StartX > UBound(FillAccelTable, 1)) Or (StartY > UBound(FillAccelTable, 2)) Then Exit Sub

If FillAccelTable(StartX, StartY) = 0 Then
   If GetPixel(FillhDC, StartX, StartY) = AreaColor Then
      SetPixelV FillhDC, StartX, StartY, FillCol
      FillAccelTable(StartX, StartY) = 1

      rX = StartX: lX = StartX
      Do
         rX = rX + 1
         If (RightDead = 0) Then
            If (FillAccelTable(rX, StartY) = 0) Then
               If (GetPixel(FillhDC, rX, StartY) = AreaColor) Then
                  SetPixelV FillhDC, rX, StartY, FillCol
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
                  SetPixelV FillhDC, lX, StartY, FillCol
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
               FillLine StartX - K, StartY - 1, FillCol, AreaColor
            End If
            MadeThrough = 1
            If FillAccelTable(StartX - K, StartY + 1) = 0 Then
               FillLine StartX - K, StartY + 1, FillCol, AreaColor
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
               Exit Sub 'EZ AZ!!!!!!!!!!!!!!!
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
               Exit Sub 'EZ AZ!!!!!!!!!!!!!!!
            End If
         Next 'K
      End If
      
      If RightMax > 0 Then
         On Error GoTo RightM
         For K = 0 To RightMax
            MadeThrough = 0
            If FillAccelTable(StartX + K, StartY - 1) = 0 Then
               FillLine StartX + K, StartY - 1, FillCol, AreaColor
            End If
            MadeThrough = 1
            If FillAccelTable(StartX + K, StartY + 1) = 0 Then
               FillLine StartX + K, StartY + 1, FillCol, AreaColor
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
               Exit Sub 'EZ AZ!!!!!!!!!!!!!!!
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
               Exit Sub 'EZ AZ!!!!!!!!!!!!!!!
            End If
         Next 'K
      End If
      'Left-Right end
      
      If (LeftMax = 0) And (RightMax = 0) Then
         On Error GoTo VertM
      'We can only fill up and down this time... a 1-pixel wide vertical line.
         MadeThrough = 0
         FillLine StartX, StartY - 1, FillCol, AreaColor
         MadeThrough = 1
         FillLine StartX, StartY + 1, FillCol, AreaColor
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
