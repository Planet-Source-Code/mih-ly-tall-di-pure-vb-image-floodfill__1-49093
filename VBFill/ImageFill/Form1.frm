VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   DrawWidth       =   2
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PatternPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   1200
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Private Declare Function ExtFloodFill Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

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
   For Y = Me.ScaleHeight / LineNum To Me.ScaleHeight Step Me.ScaleHeight / LineNum
      Me.Line (Me.ScaleWidth / LineNum, Y)-(Me.ScaleWidth - Me.ScaleWidth / LineNum, Y)
   Next 'y
   For X = Me.ScaleWidth / LineNum To Me.ScaleWidth Step Me.ScaleWidth / LineNum
      Me.Line (X, Me.ScaleHeight / LineNum)-(X, Me.ScaleHeight - Me.ScaleHeight / LineNum)
   Next 'x
ElseIf KeyCode = vbKeyF2 Then
   LineNum = 40
   For Y = Me.ScaleHeight / LineNum To Me.ScaleHeight Step Me.ScaleHeight / LineNum
      Me.Line (10, Y)-(Me.ScaleWidth - 10, Y)
      Me.PSet (Me.ScaleWidth - 10, Y)
      If Y + Me.ScaleHeight / LineNum < Me.ScaleHeight Then
         If Par = 1 Then
            Me.Line (10, Y)-(10, Y + Me.ScaleHeight / LineNum)
            Par = 0
         Else
            Me.Line (Me.ScaleWidth - 10, Y)-(Me.ScaleWidth - 10, Y + Me.ScaleHeight / LineNum)
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
   For X = 0 To LineNum - 1
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
   For X = 0 To Me.ScaleWidth
      Line (LastX, LastY)-(X, Sin(X * LineNum) * Me.ScaleHeight / 4.4 + Me.ScaleHeight / 4)
      Line (LastX, LastY + 100)-(X, Sin(X * LineNum) * Me.ScaleHeight / 4.4 + Me.ScaleHeight / 4 + 100)
      
      Line (LastX, LastY + Me.ScaleHeight / 3)-(X, Sin(X * LineNum) * Me.ScaleHeight / 4.4 + Me.ScaleHeight / 4 + Me.ScaleHeight / 3)
      Line (LastX, LastY + Me.ScaleHeight / 3 + 100)-(X, Sin(X * LineNum) * Me.ScaleHeight / 4.4 + Me.ScaleHeight / 4 + Me.ScaleHeight / 3 + 100)
      
      LastX = X
      LastY = Sin(X * LineNum) * Me.ScaleHeight / 4.4 + Me.ScaleHeight / 4
   Next 'x
ElseIf KeyCode = vbKeyF11 Then
   Me.Cls
ElseIf KeyCode = vbKeyF12 Then
   PaintErrorPlaces Me
End If

End Sub

Private Sub Form_Load()
InitPCTimer
ParseFillImage PatternPic

MsgBox "Well, here's a short walktrough of the controls:" & vbNewLine _
      & "Use F1 to F6 for some drawing tests. You should try the fill with these, too." & vbNewLine _
      & "Press F11 to clear the screen, F12 to paint the last erroneous places." & vbNewLine _
      & "On the mouse: left button to draw, right button to fill, Shift-right button for ImageFill." & vbNewLine _
      & "Press your middle mouse button to draw a circle." & vbNewLine _
      & "Have Fun!! - Msi"
      

Me.WindowState = vbMaximized
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   PSet (X, Y), 0
   LastX = X
   LastY = Y
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bInt = 0
If Button = 1 Then
   'See if we need to interpolate...
   Dist = Sqr((X - LastX) ^ 2 + (Y - LastY) ^ 2)
   If Dist > Me.DrawWidth / 2 Then
      bInt = 1
   End If
   
   Me.PSet (X, Y)
   If bInt = 1 Then
      Me.Line (LastX, LastY)-(X, Y)
   End If
LastX = X
LastY = Y
End If
If (Button = 2) And (Shift = 1) Then Me.Caption = X & ";" & Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   Select Case FILLALG
   Case "LINE"
      timS = GetPCTime
      If Shift = 1 Then
         VBFloodFill Me.hdc, CLng(X), CLng(Y), 1, Me.FillColor, Me.Point(X, Y)
      Else
         VBFloodFill Me.hdc, CLng(X), CLng(Y), 0, Me.FillColor, Me.Point(X, Y)
      End If
      Me.Caption = Round(GetPCTime - timS, 3)
   Case "WIN"
      timS = GetPCTime
      ExtFloodFill Me.hdc, X, Y, Me.Point(X, Y), 1
      Me.Caption = Round(GetPCTime - timS, 3)
   End Select
ElseIf Button = 4 Then
   Me.FillStyle = 1
   Me.Circle (X, Y), Me.ScaleHeight / 5
   Me.FillStyle = 0
End If

End Sub

