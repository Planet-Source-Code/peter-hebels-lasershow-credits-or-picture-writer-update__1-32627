Attribute VB_Name = "Module1"
'******************************************************************************************
'LaserShow Credits Writer project by Peter Hebels, Website "www.phsoft.cjb.net"           *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

'Some API calls
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Declares used for drawing the text
Public X, Y, I As Integer
Public PicArray As Integer
Public NumPics As Integer
Public DrawSpeed As Integer

'Constants for the BitBlt API call
Public Const MERGEPAINT = &HBB0226
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020

Public PicTop As Long
Public PicHalf As Long

'This is the main text-draw sub
Public Function StartTextDraw()
'The color variable stored from GetPixel
Dim PixCol As Long
'For the color values
Dim r, g, b As Integer
'Resume after every error
On Error Resume Next
                 
'are we at the end of our show?
If PicArray >= NumPics - 1 Then
 MsgBox "This was the show, thanx for trying this app!", vbInformation, "See ya next time!"
  'Stop the timer, otherwise it will keep trying to call the Function
  Form1.Timer1.Enabled = False
  'End the app
  Unload Form1
  End
 Exit Function
End If

'I use ControlArrays in this app for easy coding!
'Every call to this Function will go to the next picturebox in the array
PicArray = PicArray + 1

'Start some loops, so the text-draw routine will begin
'Calculate the X value
For X = Form1.Picture1(PicArray).Width - 1 To 0 Step -1
'Calculate the Y value
For Y = 0 To Form1.Picture1(PicArray).Height - 1
    
    'Move the Laser together with X
    Form1.Picture2.Top = X
    'Get the pixel color from the picture
    PixCol = GetPixel(Form1.Picture1(PicArray).hDC, X, Y)
    
    'Make the image 16Bits color
    r = PixCol Mod 256
    b = Int(PixCol / 65536)
    g = (PixCol - (b * 65536) - r) / 256
     
    'Draw the lines (Laser Beams :))
    PicTop = Form1.Picture2.Top
    PicHalf = Form1.Picture2.Height / 2
    Form1.Line (35, PicTop + PicHalf)-(X + 100, Y + 50), RGB(r, g, b)

Next Y
'Let windows do its events
DoEvents
'This changes the speed, with can be set in FormLoad
'This also reduces the CPU usage a lot ~90%
Sleep DrawSpeed
Next X

'Clear the picture
Form1.Picture = Form1.Picture
'Enable the timer, so the picture goes down and disapears
Form1.Timer1.Enabled = True
End Function

