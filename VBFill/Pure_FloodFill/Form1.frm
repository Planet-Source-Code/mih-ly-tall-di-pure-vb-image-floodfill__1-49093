VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   DrawWidth       =   2
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   347
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Private Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Private Const FILLALG = "Line"

Dim LastX As Single, LastY As Single, Dist As Single
Dim sX As Long, sY As Long, timS As Single

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim r As Single
Dim f As Single
Dim LineNum As Double
Dim Par As Byte

If KeyCode = vbKeyF1 Then
   LineNum = 40
   For y = Me.ScaleHeight / LineNum To Me.ScaleHeight Step Me.ScaleHeight / LineNum
      Me.Line (Me.ScaleWidth / LineNum, y)-(Me.ScaleWidth - Me.ScaleWidth / LineNum, y)
   Next 'y
   For x = Me.ScaleWidth / LineNum To Me.ScaleWidth Step Me.ScaleWidth / LineNum
      Me.Line (x, Me.ScaleHeight / LineNum)-(x, Me.ScaleHeight - Me.ScaleHeight / LineNum)
   Next 'x
ElseIf KeyCode = vbKeyF2 Then
   LineNum = 40
   For y = Me.ScaleHeight / LineNum To Me.ScaleHeight Step Me.ScaleHeight / LineNum
      Me.Line (10, y)-(Me.ScaleWidth - 10, y)
      Me.PSet (Me.ScaleWidth - 10, y)
      If y + Me.ScaleHeight / LineNum < Me.ScaleHeight Then
         If Par = 1 Then
            Me.Line (10, y)-(10, y + Me.ScaleHeight / LineNum)
            Par = 0
         Else
            Me.Line (Me.ScaleWidth - 10, y)-(Me.ScaleWidth - 10, y + Me.ScaleHeight / LineNum)
            Par = 1
         End If
      End If
   Next 'y
ElseIf KeyCode = vbKeyF3 Then
   LastX = Me.ScaleWidth / 2
   LastY = Me.ScaleHeight / 2
   For r = 0 To Me.ScaleHeight / 2 Step 0.05
      f = f + 1
      sX = r * Cos(f * (3.14159265358979 / 180)) + Me.ScaleWidth / 2
      sY = r * Sin(f * (3.14159265358979 / 180)) + Me.ScaleHeight / 2
      
      Me.Line (LastX, LastY)-(sX, sY)
      LastX = sX
      LastY = sY
   Next 'r
ElseIf KeyCode = vbKeyF4 Then
   r = Me.ScaleHeight / 2
   For f = 0 To 180 Step 5
      sX = r * Cos(f * (3.14159265358979 / 180)) + Me.ScaleWidth / 2
      sY = r * Sin(f * (3.14159265358979 / 180)) + Me.ScaleHeight / 2
      
      LastX = r * Cos((f + 180) * (3.14159265358979 / 180)) + Me.ScaleWidth / 2
      LastY = r * Sin((f + 180) * (3.14159265358979 / 180)) + Me.ScaleHeight / 2
      
      Me.Line (LastX, LastY)-(sX, sY)
   Next 'f
ElseIf KeyCode = vbKeyF5 Then
   LineNum = 150
   Randomize Timer
   LastX = Int((Me.ScaleWidth + 1) * Rnd)
   LastY = Int((Me.ScaleHeight + 1) * Rnd)
   For x = 0 To LineNum - 1
      sX = Int((Me.ScaleWidth + 1) * Rnd)
      sY = Int((Me.ScaleHeight + 1) * Rnd)
      Me.Line (LastX, LastY)-(sX, sY)
      LastX = sX
      LastY = sY
   Next 'x
ElseIf KeyCode = vbKeyF6 Then
   LineNum = 0.04 'Now this means frequency.
   LastX = 0
   LastY = Me.ScaleHeight / 4
   For x = 0 To Me.ScaleWidth
      Line (LastX, LastY)-(x, Sin(x * LineNum) * Me.ScaleHeight / 4.4 + Me.ScaleHeight / 4)
      Line (LastX, LastY + 100)-(x, Sin(x * LineNum) * Me.ScaleHeight / 4.4 + Me.ScaleHeight / 4 + 100)
      
      Line (LastX, LastY + Me.ScaleHeight / 3)-(x, Sin(x * LineNum) * Me.ScaleHeight / 4.4 + Me.ScaleHeight / 4 + Me.ScaleHeight / 3)
      Line (LastX, LastY + Me.ScaleHeight / 3 + 100)-(x, Sin(x * LineNum) * Me.ScaleHeight / 4.4 + Me.ScaleHeight / 4 + Me.ScaleHeight / 3 + 100)
      
      LastX = x
      LastY = Sin(x * LineNum) * Me.ScaleHeight / 4.4 + Me.ScaleHeight / 4
   Next 'x
ElseIf KeyCode = vbKeyF11 Then
   Me.Cls
ElseIf KeyCode = vbKeyF12 Then
   PaintErrorPlaces Me
End If

End Sub

Private Sub Form_Load()
InitPCTimer
Me.WindowState = vbMaximized
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   PSet (x, y), 0
   LastX = x
   LastY = y
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
bInt = 0
If Button = 1 Then
   'See if we need to interpolate...
   Dist = Sqr((x - LastX) ^ 2 + (y - LastY) ^ 2)
   If Dist > Me.DrawWidth / 2 Then
      bInt = 1
   End If
   
   Me.PSet (x, y)
   If bInt = 1 Then
      Me.Line (LastX, LastY)-(x, y)
   End If
LastX = x
LastY = y
End If
If (Button = 2) And (Shift = 1) Then Me.Caption = x & ";" & y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If (Button = 2) And (Shift <> 1) Then
   Select Case FILLALG
   Case "LINE"
      timS = GetPCTime
      VBFloodFill Me.hdc, CLng(x), CLng(y), Me.FillColor, Me.Point(x, y)
      Me.Caption = Round(GetPCTime - timS, 3)
   Case "WIN"
      timS = GetPCTime
      ExtFloodFill Me.hdc, x, y, Me.Point(x, y), 1
      Me.Caption = Round(GetPCTime - timS, 3)
   End Select
ElseIf Button = 4 Then
   Me.FillStyle = 1
   Me.Circle (x, y), Me.ScaleHeight / 5
   Me.FillStyle = 0
End If
End Sub

