Attribute VB_Name = "Module1"
'Outlying Pixels - II
'by Mehmet Hakan Satman
'Istanbul University - Departments of Econometrics
'mhsatman@yahoo.com

'This program gets mouses coordinates and captures screen.
'Then calculates main statistics (mean and variance) from captured screen
'and defines a volume of area as an outlier pixels set where pixel color values
'exceed a critical value.
'Finally calculates minimum volume rectangle and shows outlying area

'Note: For more robust outlying applications use median instead of mean.

'For basic operation of this small algorithm look "Outlying Pixels" in Source
'code planet web site by me.

'enjoy.


Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type


Public Type OutlyingObject
    Left As Double
    Top As Double
    Rigth As Double
    Bottom As Double
End Type




Sub GetScreen(ByRef target As PictureBox, Optional ByVal w As Integer, Optional ByVal h As Integer)
Dim mydc As Long, TargetDC As Long
Dim mousex As Long, mousey As Long
Dim mypoint As POINTAPI

If w = 0 Or IsEmpty(w) Then w = target.Width
If h = 0 Or IsEmpty(h) Then h = target.Height

target.Cls
mydc = GetDC(0)
TargetDC = target.hdc
GetCursorPos mypoint

mousex = mypoint.x - 20
mousey = mypoint.y - 20

BitBlt TargetDC, 0, 0, w, h, mydc, mousex, mousey, SRCCOPY
target.Refresh
DoEvents

ReleaseDC 0, mydc
End Sub

Sub CalculateDescriptives(ByRef target As PictureBox, ByRef mean As Double, ByRef variance As Double)
Dim x, y, x1, y1
Dim k

k = 0
x1 = target.Width
y1 = target.Height
For x = 0 To x1
    For y = 0 To y1
        k = k + 1
        mean = mean + target.Point(x, y)
    Next
Next
mean = mean / k

For x = 0 To x1
    For y = 0 To y1
        variance = variance + (target.Point(x, y) - mean) ^ 2
    Next
Next

variance = variance / k
End Sub

Sub DrawZ(source As PictureBox, target As PictureBox, mean As Double, variance As Double, criticalvalue As Double)
Dim x, y, x1, y1
Dim s, z

s = Sqr(variance)

x1 = source.Width
y1 = source.Height
target.Cls

For x = 0 To x1
    For y = 0 To y1
        k = source.Point(x, y)
        z = Abs((k - mean) / s)
        If z >= criticalvalue Then target.PSet (x, y), k
    Next
Next
target.Refresh
End Sub

Sub Stats(source As PictureBox, MainPicture As PictureBox)
'If there are some outlying pixels, take them into a box
Dim x1 As Double, y1 As Double, x As Double, y As Double
Dim Fexit As Boolean
Dim myout As OutlyingObject



x1 = source.Width
y1 = source.Height
Fexit = False

'Getting upper left coordinate
For x = 0 To x1
For y = 0 To y1
If source.Point(x, y) > 0 Then
Fexit = True
Exit For
End If
Next
If Fexit Then Exit For
Next

myout.Left = x

Fexit = False

'Getting upper top coordinate
For y = 0 To y1
For x = 0 To x1
If source.Point(x, y) > 0 Then
Fexit = True
Exit For
End If
Next
If Fexit Then Exit For
Next

myout.Top = y

Fexit = False

'Getting bottom coordinate
For y = y1 To 0 Step -1
For x = x1 To 0 Step -1
If source.Point(x, y) > 0 Then
Fexit = True
Exit For
End If
Next
If Fexit Then Exit For
Next

myout.Bottom = y

Fexit = False

'Getting bottom coordinate
For x = x1 To 0 Step -1
For y = 0 To y1
If source.Point(x, y) > 0 Then
Fexit = True
Exit For
End If
Next
If Fexit Then Exit For
Next

myout.Rigth = x


MainPicture.Line (myout.Left, myout.Top)-(myout.Rigth, myout.Bottom), QBColor(15), B
MainPicture.Line (myout.Left - 1, myout.Top - 1)-(myout.Rigth + 1, myout.Bottom + 1), QBColor(0), B
MainPicture.Line (myout.Left - 2, myout.Top - 2)-(myout.Rigth + 2, myout.Bottom + 2), QBColor(15), B

End Sub


