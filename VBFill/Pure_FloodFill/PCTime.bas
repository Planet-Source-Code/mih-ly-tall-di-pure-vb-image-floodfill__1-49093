Attribute VB_Name = "PCTime"
Public Type LARGE_INTEGER
   lowpart As Long
   highpart As Long
End Type

Public Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (lpFrequency As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32.dll" (lpPerformanceCount As LARGE_INTEGER) As Long

Dim PCFreq As Double, TempLint As LARGE_INTEGER

Public Sub InitPCTimer()
QueryPerformanceFrequency TempLint
PCFreq = DoubleFromLInt(TempLint.highpart, TempLint.lowpart)
End Sub

Public Function GetPCTime() As Double
QueryPerformanceCounter TempLint
GetPCTime = DoubleFromLInt(TempLint.highpart, TempLint.lowpart) / PCFreq
End Function

Private Function CLargeInt(Lo As Long, Hi As Long) As Double
Dim dblLo As Double, dblHi As Double
If Lo < 0 Then
    dblLo = 2 ^ 32 + Lo
Else
    dblLo = Lo
End If
If Hi < 0 Then
    dblHi = 2 ^ 32 + Hi
Else
    dblHi = Hi
End If

CLargeInt = dblLo + dblHi * 2 ^ 32
End Function

Private Function DoubleFromLInt(highpart As Long, lowpart As Long) As Double
   DoubleFromLInt = highpart * 2 ^ 32 + lowpart
End Function

