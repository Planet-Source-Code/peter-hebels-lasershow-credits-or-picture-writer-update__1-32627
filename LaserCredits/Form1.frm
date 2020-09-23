VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LaserShow credits writer V-II"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12705
   DrawWidth       =   3
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   847
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   105
      TabIndex        =   8
      Top             =   0
      Width           =   105
      Begin VB.Image Image2 
         Height          =   5490
         Left            =   0
         Picture         =   "Form1.frx":32594
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   6
      Left            =   7920
      Picture         =   "Form1.frx":33546
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   7
      Top             =   4320
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   5
      Left            =   7920
      Picture         =   "Form1.frx":37408
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   4
      Left            =   7920
      Picture         =   "Form1.frx":3B2CA
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   3
      Left            =   7920
      Picture         =   "Form1.frx":3F18C
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   2
      Left            =   7920
      Picture         =   "Form1.frx":4304E
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   1
      Left            =   7920
      Picture         =   "Form1.frx":46F10
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   450
      TabIndex        =   1
      Top             =   600
      Width           =   450
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         Picture         =   "Form1.frx":4ADD2
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   0
      Left            =   7920
      Picture         =   "Form1.frx":4B634
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   4635
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   8
      Y1              =   0
      Y2              =   368
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'LaserShow Credits Writer V-II project by Peter Hebels, Website "www.phsoft.cjb.net"           *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

Dim I As Integer 'Used for controlling the ScaleHeight of the form
Dim I2 As Integer

Private Sub Form_Load()
 
 'Form control code:
 Form1.Show 'Show the form
 Me.ScaleMode = 3 'Set the scalemode to pixels because we going to use bitblt
 Me.Width = 7140 'Don't show the pictureboxes
  
 'set the values:
 NumPics = 7 'How many pict's have we used
 DrawSpeed = 10 'How fast are the letters drawn, Higher value = Slower speed
 PicArray = -1 'We begin at 0
  
 Form1.AutoRedraw = True 'Set AutoRedraw to true, otherwise nothing will appear.
  
 StartTextDraw 'Start the text-Draw function
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Timer1.Enabled = False 'Make sure the timer is stopped
Unload Form1 'Unload the form
End 'End the program
End Sub

Private Sub Timer1_Timer()
'This timersub does the downgoing text job
On Error Resume Next 'Resume on errors

If I > Form1.ScaleHeight - Picture1(PicArray).Height - 100 Then
  I2 = I 'Keep the I value for the waiting picture
  Timer1.Enabled = False 'Dissable the timer so the moving picture stops
  I = 0 'Reset I to 0 because we have writed a new text
  StartTextDraw 'Restart the drawing function
 Exit Sub
End If

I = I + 2 'For the image going down
I2 = I2 + 1 'For the waiting image

Picture2.Top = I 'Move the laser down

'BitBlt the image using I as the Y value, so its going down
BitBlt Form1.hDC, X + 102, Y - 2 + I, Form1.Picture1(PicArray).Width, Form1.Picture1(PicArray).Height, Form1.Picture1(PicArray).hDC, 0, 0, SRCCOPY
BitBlt Form1.hDC, X + 102, Y - 2 + I2, Form1.Picture1(PicArray - 1).Width, Form1.Picture1(PicArray - 1).Height, Form1.Picture1(PicArray - 1).hDC, 0, 0, SRCCOPY
'Refresh the form every single frame.
Form1.Refresh

End Sub
