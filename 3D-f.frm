VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   270
   ClientTop       =   990
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   Begin VB.Timer TIMERofDIRECTIONS 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6165
      Top             =   2295
   End
   Begin VB.CommandButton btnHideDots 
      Caption         =   "Hide All Dots"
      Height          =   330
      Left            =   4680
      TabIndex        =   9
      Top             =   4410
      Width           =   1455
   End
   Begin VB.TextBox txtCord 
      Height          =   330
      Index           =   7
      Left            =   5580
      TabIndex        =   8
      Top             =   1575
      Width           =   915
   End
   Begin VB.TextBox txtCord 
      Height          =   330
      Index           =   6
      Left            =   5580
      TabIndex        =   7
      Top             =   1215
      Width           =   915
   End
   Begin VB.TextBox txtCord 
      Height          =   330
      Index           =   5
      Left            =   4635
      TabIndex        =   6
      Top             =   1215
      Width           =   915
   End
   Begin VB.TextBox txtCord 
      Height          =   330
      Index           =   4
      Left            =   4590
      TabIndex        =   5
      Top             =   495
      Width           =   915
   End
   Begin VB.TextBox txtCord 
      Height          =   330
      Index           =   3
      Left            =   5535
      TabIndex        =   4
      Top             =   495
      Width           =   915
   End
   Begin VB.TextBox txtCord 
      Height          =   330
      Index           =   2
      Left            =   5535
      TabIndex        =   3
      Top             =   135
      Width           =   915
   End
   Begin VB.TextBox txtCord 
      Height          =   330
      Index           =   1
      Left            =   4590
      TabIndex        =   2
      Top             =   135
      Width           =   915
   End
   Begin VB.TextBox txtCord 
      Height          =   330
      Index           =   8
      Left            =   4635
      TabIndex        =   1
      Top             =   1575
      Width           =   915
   End
   Begin VB.PictureBox picFIELD 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DrawWidth       =   2
      ForeColor       =   &H0000FF00&
      Height          =   4830
      Left            =   90
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   0
      Top             =   0
      Width           =   4065
      Begin VB.Shape shSquare 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   90
         Index           =   8
         Left            =   1530
         Top             =   2160
         Width           =   90
      End
      Begin VB.Shape shSquare 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   90
         Index           =   7
         Left            =   2205
         Top             =   2160
         Width           =   90
      End
      Begin VB.Shape shSquare 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   90
         Index           =   6
         Left            =   2340
         Top             =   1575
         Width           =   90
      End
      Begin VB.Shape shSquare 
         BackColor       =   &H00FFFF00&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   90
         Index           =   5
         Left            =   1665
         Top             =   1575
         Width           =   90
      End
      Begin VB.Shape shSquare 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   5  'Dash-Dot-Dot
         BorderWidth     =   2
         Height          =   90
         Index           =   4
         Left            =   1305
         Top             =   2520
         Width           =   90
      End
      Begin VB.Shape shSquare 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   5  'Dash-Dot-Dot
         BorderWidth     =   2
         Height          =   90
         Index           =   3
         Left            =   1980
         Top             =   2520
         Width           =   90
      End
      Begin VB.Shape shSquare 
         BackColor       =   &H0000FFFF&
         BorderColor     =   &H00000000&
         BorderStyle     =   5  'Dash-Dot-Dot
         BorderWidth     =   2
         Height          =   90
         Index           =   0
         Left            =   2970
         Top             =   90
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Shape shSquare 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   90
         Index           =   1
         Left            =   1440
         Top             =   1980
         Width           =   90
      End
      Begin VB.Shape shSquare 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         BorderStyle     =   5  'Dash-Dot-Dot
         BorderWidth     =   2
         Height          =   90
         Index           =   2
         Left            =   2115
         Top             =   1980
         Width           =   90
      End
   End
   Begin VB.Image imgDirection 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   4
      Left            =   4905
      Picture         =   "3D-f.frx":0000
      Stretch         =   -1  'True
      Top             =   2655
      Width           =   300
   End
   Begin VB.Image imgDirection 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   3
      Left            =   5265
      Picture         =   "3D-f.frx":0442
      Stretch         =   -1  'True
      Top             =   2925
      Width           =   300
   End
   Begin VB.Image imgDirection 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   2
      Left            =   5625
      Picture         =   "3D-f.frx":0884
      Stretch         =   -1  'True
      Top             =   2655
      Width           =   300
   End
   Begin VB.Image imgDirection 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   5265
      Picture         =   "3D-f.frx":0CC6
      Stretch         =   -1  'True
      Top             =   2385
      Width           =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    =======================================
'     3-D rotating cube (Incomplete)
'    =======================================
'
'    I just tried to rotate a 3-D cube,
'    this is what I got, pretty interesting... :)
'
'    Visit my Homepage:
'    http://www.geocities.com/emu8086/vb/
'
'
'    Last Update: Thursday, July 11, 2002
'
'
'    Copyright 2002 Alexander Popov Emulation Soft.
'               All rights reserved.
'        http://www.geocities.com/emu8086/


Option Explicit

Dim n As Integer
Dim iDistanceX As Integer
Dim iDistanceY As Integer
Dim iDots As Integer     ' Number of All Dots.
Dim iLayer1Dots As Integer
Dim iLayer2Dots As Integer
Dim iDirection As Integer

Private Sub btnHideDots_Click()
 Dim iX As Integer
 
 For iX = 1 To iDots
    shSquare(iX).Visible = Not (shSquare(iX).Visible)
 Next iX

 DRAW_LINES

End Sub

Private Sub Form_Load()
  iLayer1Dots = 4
  iLayer2Dots = 4
  iDots = iLayer1Dots + iLayer2Dots
End Sub

Private Sub imgDirection_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 iDirection = Index
 TIMERofDIRECTIONS.Enabled = True
End Sub

Private Sub imgDirection_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 TIMERofDIRECTIONS.Enabled = False
End Sub

Private Sub picFIELD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

DESELECT_ALL ' Just a Reset.

For n = 1 To iDots

  iDistanceX = X - shSquare(n).Left
  iDistanceY = Y - shSquare(n).Top

   If iDistanceX < shSquare(n).Width And iDistanceY < shSquare(n).Height And iDistanceX > 0 And iDistanceY > 0 Then
      shSquare(n).BorderColor = RGB(0, 0, 255)  ' shSquare(n) was clicked! , so let's color him.
      Exit Sub
   End If
   
Next n

DESELECT_ALL
End Sub

Private Sub picFIELD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 And n <> 0 Then
   shSquare(n).Top = Y - iDistanceY
   shSquare(n).Left = X - iDistanceX
   DRAW_LINES
 End If
 
 OUTPUT_CORDINATS
 
End Sub

Private Sub DESELECT_ALL()
      For n = 1 To iDots
      shSquare(n).BorderColor = RGB(0, 0, 0)
     Next n
      n = 0 ' nothing selected.
End Sub

Private Sub DRAW_LINES()
Dim iX As Integer ' used in for - next.

 picFIELD.Cls ' Clear the picFIELD!
 
' The fareest Layers must be Drawn first.
'__________________________________Layer2______________________________________
picFIELD.ForeColor = RGB(255, 0, 0)
 picFIELD.Line (shSquare(5).Left + shSquare(5).Width / 2, shSquare(5).Top + shSquare(5).Height / 2)-(shSquare(6).Left + shSquare(6).Width / 2, shSquare(6).Top + shSquare(6).Height / 2)
 
If iLayer1Dots < 3 Then Exit Sub   ' Nothing more to Draw.
For iX = 7 To 8 ' Dots 5 and 6 were drawn before, now we are drawing last ones.
   picFIELD.Line -(shSquare(iX).Left + shSquare(iX).Width / 2, shSquare(iX).Top + shSquare(iX).Height / 2)
Next iX

  picFIELD.Line -(shSquare(5).Left + shSquare(5).Width / 2, shSquare(5).Top + shSquare(5).Height / 2) ' Draw back to shSquare(1)
'------------------------------------------------------------------------------------------------------------------------------------------------------------

'__________________________________Between Layer2 & Layer1__________________________
picFIELD.ForeColor = RGB(0, 0, 255)
  picFIELD.Line (shSquare(1).Left + shSquare(1).Width / 2, shSquare(1).Top + shSquare(1).Height / 2)-(shSquare(5).Left + shSquare(5).Width / 2, shSquare(5).Top + shSquare(5).Height / 2)
  picFIELD.Line (shSquare(2).Left + shSquare(2).Width / 2, shSquare(2).Top + shSquare(2).Height / 2)-(shSquare(6).Left + shSquare(6).Width / 2, shSquare(6).Top + shSquare(6).Height / 2)
  picFIELD.Line (shSquare(3).Left + shSquare(3).Width / 2, shSquare(3).Top + shSquare(3).Height / 2)-(shSquare(7).Left + shSquare(7).Width / 2, shSquare(7).Top + shSquare(7).Height / 2)
  picFIELD.Line (shSquare(4).Left + shSquare(4).Width / 2, shSquare(4).Top + shSquare(4).Height / 2)-(shSquare(8).Left + shSquare(8).Width / 2, shSquare(8).Top + shSquare(8).Height / 2)




'__________________________________Layer1______________________________________
picFIELD.ForeColor = RGB(0, 255, 0)
 picFIELD.Line (shSquare(1).Left + shSquare(1).Width / 2, shSquare(1).Top + shSquare(1).Height / 2)-(shSquare(2).Left + shSquare(2).Width / 2, shSquare(2).Top + shSquare(2).Height / 2)
 
If iLayer1Dots < 3 Then Exit Sub   ' Nothing more to Draw.
For iX = 3 To iLayer1Dots ' Dots 1 and 2 were drawn before, now we are drawing last ones.
   picFIELD.Line -(shSquare(iX).Left + shSquare(iX).Width / 2, shSquare(iX).Top + shSquare(iX).Height / 2)
Next iX

  picFIELD.Line -(shSquare(1).Left + shSquare(1).Width / 2, shSquare(1).Top + shSquare(1).Height / 2) ' Draw back to shSquare(1)
'------------------------------------------------------------------------------------------------------------------------------------------------------------

OUTPUT_CORDINATS

End Sub

Private Sub picFIELD_Paint()
 DRAW_LINES
End Sub

Private Sub OUTPUT_CORDINATS()
Dim iX As Integer

   For iX = 1 To iDots
    txtCord(iX).Text = shSquare(iX).Left & "-" & shSquare(iX).Top
   Next iX
   
End Sub

Private Sub TIMERofDIRECTIONS_Timer()
' Used to Generate Move when an Arrow is pressed.
  Dim iX As Integer
  
   Select Case iDirection
      Case 1 ' up
            For iX = 1 To 2
                   shSquare(iX).Top = shSquare(iX).Top + 1
            Next iX
  
            For iX = 3 To 4
                  shSquare(iX).Top = shSquare(iX).Top - 1
            Next iX
            
            For iX = 5 To 6
                   shSquare(iX).Top = shSquare(iX).Top + 1
            Next iX
  
            For iX = 7 To 8
                  shSquare(iX).Top = shSquare(iX).Top - 1
            Next iX
            
      Case 3 ' down
            For iX = 1 To 2
                   shSquare(iX).Top = shSquare(iX).Top - 1
            Next iX
  
            For iX = 3 To 4
                  shSquare(iX).Top = shSquare(iX).Top + 1
            Next iX
            
            For iX = 5 To 6
                   shSquare(iX).Top = shSquare(iX).Top - 1
            Next iX
  
            For iX = 7 To 8
                  shSquare(iX).Top = shSquare(iX).Top + 1
            Next iX
            
   End Select

 DRAW_LINES  ' Repainting.

End Sub
