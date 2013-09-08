VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vroom 1.1"
   ClientHeight    =   8640
   ClientLeft      =   2430
   ClientTop       =   1725
   ClientWidth     =   11280
   Icon            =   "vroom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   11280
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Top             =   8160
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&About Us"
      Height          =   375
      Left            =   9480
      TabIndex        =   4
      Top             =   8160
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Restart"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   8160
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10320
      Top             =   8040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10800
      Top             =   8040
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7815
      Left            =   0
      ScaleHeight     =   7755
      ScaleWidth      =   11235
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   165
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   5040
         Shape           =   3  'Circle
         Top             =   3120
         Width           =   165
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   6
      Top             =   8160
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   8040
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   8160
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------License---------------------------------------------

'Note: This resource has been created by Ananda Prakash Verma (http://www.apverma.com)

'TERMS OF USE:

'You may freely download, play, redistribute and use it into your software project/college project or you can extend it without removing credits of the original authors.

'If you use/modify the resources in your projects please linkback to the resource page (https://vroom.apverma.com). (Please don’t link directly to the .zip files, please link to the resource page.)

'If you should have any questions please contact me here: http://www.apverma.com, alternatively you can mail me at: apverma[at]apverma.com

'-----------------------------------License----------------------------------------------

Dim a As Integer
Dim b As Integer
Dim dir1 As Integer
Dim dir2 As Integer
Dim coll1, coll2 As Long
Dim Tim As Integer
Dim sensitivity As Integer
Private Declare Function Beep Lib "kernel32" _
 (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Dim gameover1 As Boolean
Dim gameover2 As Boolean
Private Sub Command1_Click()
initialize
Picture1.SetFocus
End Sub
Private Sub Command2_Click()
frmAbout.Show
End Sub
Private Sub Command3_Click()
End
End Sub
Private Sub Form_Load()
initialize
End Sub
Private Sub Picture1_KeyPress(KeyAscii As Integer)
'Debug.Print Tim
'debug.Print a
'If Tim = 5 Then
'a = a + 5
'b = b + 5
'Tim = 0
'Else
'Tim = Tim + 1
'End If

If KeyAscii = 119 Or KeyAscii = 115 Or KeyAscii = 97 Or KeyAscii = 100 Then
If gameover1 = False Then
Label2.Caption = "Player1 changed his direction"
End If
ElseIf KeyAscii = 52 Or KeyAscii = 54 Or KeyAscii = 53 Or KeyAscii = 56 Then
If gameover2 = False Then
Label2.Caption = "Player2 changed his direction"
End If
Else
If gameover1 = False And gameover2 = False Then
Label2.Caption = "You fools, press the right keys WASD(p1) or IJKL(p2)"
End If
End If

Debug.Print KeyAscii
If gameover1 = False Then
Timer1.Enabled = True
End If
If gameover2 = False Then
Timer2.Enabled = True
End If

Select Case KeyAscii

'-------------------for yellow ball
Case 119
'Shape3.Top = Shape1.Top - Shape1.Height + Shape3.Height
'Shape3.Left = Shape1.Lefts
dir1 = 0
Case 115
'Shape3.Top = Shape1.Top + Shape1.Height
'Shape3.Left = Shape1.Left
dir1 = 1
Case 97
'Shape3.Top = Shape1.Top + Shape1.Height / 2 - Shape3.Height / 2
'Shape3.Left = Shape1.Left - Shape1.Width / 2 - Shape3.Width / 4
dir1 = 2
Case 100
'Shape3.Top = Shape1.Top + Shape1.Height / 2 - Shape3.Height / 2
'Shape3.Left = Shape1.Left + Shape1.Width / 2 + Shape3.Width / 4
dir1 = 3

'-----------------------for green ball
Case 105
'Shape4.Top = Shape2.Top - Shape2.Height + Shape4.Height
''Shape4.Left = Shape2.Left
dir2 = 0
Case 107
'Shape4.Top = Shape2.Top + Shape2.Height
'Shape4.Left = Shape2.Left
dir2 = 1
Case 106
'Shape4.Top = Shape2.Top + Shape2.Height / 2 - Shape4.Height / 2
'Shape4.Left = Shape2.Left - Shape2.Width / 2 - Shape4.Width / 4
dir2 = 2
Case 108
'Shape4.Top = Shape2.Top + Shape2.Height / 2 - Shape4.Height / 2
'Shape4.Left = Shape2.Left + Shape2.Width / 2 + Shape4.Width / 4
dir2 = 3
Case Default
dir1 = 4
dir2 = 4
End Select
End Sub

Private Sub Timer1_Timer()
Picture1.DrawWidth = 2
Select Case dir1
Case 0 'goin up
sensorx = Shape1.Left
sensory = Shape1.Top - sensitivity
coll2 = Picture1.Point(sensorx, sensory)
If coll1 <> coll2 Then
Shape1.Visible = False
playsound1
Label2.Caption = "player1 hit the wall"
Label4.Caption = "player2 wins"
End If
Shape1.Top = Shape1.Top - a
'Shape3.Top = Shape3.Top - a
Picture1.PSet (Shape1.Left + Shape1.Width / 2, Shape1.Top + Shape1.Height / 2), vbYellow
Case 1 'goin down
sensorx = Shape1.Left
sensory = Shape1.Top + Shape1.Height + sensitivity
coll2 = Picture1.Point(sensorx, sensory)
If coll1 <> coll2 Then
Shape1.Visible = False
playsound1
Label2.Caption = "player1 hit the wall"
Label4.Caption = "player2 wins"
End If

Shape1.Top = Shape1.Top + a
'Shape3.Top = Shape3.Top + a
Picture1.PSet (Shape1.Left + Shape1.Width / 2, Shape1.Top + Shape1.Height / 2), vbYellow
Case 2

sensorx = Shape1.Left - sensitivity
sensory = Shape1.Top
coll2 = Picture1.Point(sensorx, sensory)
If coll1 <> coll2 Then
Shape1.Visible = False
playsound1
Label2.Caption = "player1 hit the wall"
Label4.Caption = "player2 wins"
End If

Shape1.Left = Shape1.Left - a
'Shape3.Left = Shape3.Left - a
Picture1.PSet (Shape1.Left + Shape1.Width / 2, Shape1.Top + Shape1.Height / 2), vbYellow
Case 3

sensorx = Shape1.Left + Shape1.Width + sensitivity
sensory = Shape1.Top
coll2 = Picture1.Point(sensorx, sensory)
If coll1 <> coll2 Then
Shape1.Visible = False
playsound1
Label2.Caption = "player1 hit the wall"
Label4.Caption = "player2 wins"
End If

Shape1.Left = Shape1.Left + a
'Shape3.Left = Shape3.Left + a
Picture1.PSet (Shape1.Left + Shape1.Width / 2, Shape1.Top + Shape1.Height / 2), vbYellow
Case 4
Timer1.Enabled = False
End Select
End Sub

Private Sub Timer2_Timer()
Select Case dir2

Case 0
sensorx = Shape2.Left
sensory = Shape2.Top - sensitivity
coll2 = Picture1.Point(sensorx, sensory)
If coll1 <> coll2 Then
Shape2.Visible = False
playsound2

Label2.Caption = "player2 hit the wall"
Label4.Caption = "player1 wins"
End If

Shape2.Top = Shape2.Top - b
'Shape4.Top = Shape4.Top - a
Picture1.PSet (Shape2.Left + Shape2.Width / 2, Shape2.Top + Shape2.Height / 2), vbGreen

Case 1
sensorx = Shape2.Left
sensory = Shape2.Top + Shape2.Height + sensitivity
coll2 = Picture1.Point(sensorx, sensory)
If coll1 <> coll2 Then
Shape2.Visible = False
playsound2

Label2.Caption = "player2 hit the wall"
Label4.Caption = "player1 wins"
End If

Shape2.Top = Shape2.Top + b
'Shape4.Top = Shape4.Top + a
Picture1.PSet (Shape2.Left + Shape2.Width / 2, Shape2.Top + Shape2.Height / 2), vbGreen

Case 2
sensorx = Shape2.Left - sensitivity
sensory = Shape2.Top
coll2 = Picture1.Point(sensorx, sensory)
If coll1 <> coll2 Then
Shape2.Visible = False
playsound2

Label2.Caption = "player2 hit the wall"
Label4.Caption = "player1 wins"
End If

Shape2.Left = Shape2.Left - b
'Shape4.Left = Shape4.Left - a
Picture1.PSet (Shape2.Left + Shape2.Width / 2, Shape2.Top + Shape2.Height / 2), vbGreen

Case 3
sensorx = Shape2.Left + Shape2.Width + sensitivity
sensory = Shape2.Top
coll2 = Picture1.Point(sensorx, sensory)
If coll1 <> coll2 Then
Shape2.Visible = False
playsound2

Label2.Caption = "player2 hit the wall"
Label4.Caption = "player1 wins"
End If

Shape2.Left = Shape2.Left + b
'Shape4.Left = Shape4.Left + a
Picture1.PSet (Shape2.Left + Shape2.Width / 2, Shape2.Top + Shape2.Height / 2), vbGreen

Case 4
Timer2.Enabled = False
End Select
End Sub
Public Sub initialize()
gameover1 = False
gameover2 = False
If Shape1.Visible = False Then
Shape1.Visible = True
End If
If Shape2.Visible = False Then
Shape2.Visible = True
End If
'Picture1.Top = Form1.Top
'Picture1.Height = Form1.Height - 1000
Picture1.Width = Form1.Width
Picture1.Cls
Shape1.Top = Picture1.Top + 300
Shape1.Left = Picture1.Width / 2
Shape2.Top = Picture1.Top + Picture1.Height - 600
Shape2.Left = Picture1.Width / 2
b = 30
a = 30
Tim = 0
sensitivity = 5
dir1 = 4
dir2 = 4
coll1 = Picture1.Point(0, 0)
Dim sensorx As Integer
Dim sensory As Integer
x = 0
y = 0
Label2.Caption = "Welcome to Vroom, P1(Red, key - WSAD), P2(Blue, key - IJKL), have fun"
Label4.Caption = "Start the game"
End Sub

Public Sub playsound1()
gameover1 = True
Beep 400, 400
Timer1.Enabled = False
End Sub

Public Sub playsound2()
gameover2 = True
Beep 400, 400
Timer2.Enabled = False
End Sub


'-----------------------------------License---------------------------------------------

'Note: This resource has been created by Ananda Prakash Verma (http://www.apverma.com)

'TERMS OF USE:

'You may freely download, play, redistribute and use it into your software project/college project or you can extend it without removing credits of the original authors.

'If you use/modify the resources in your projects please linkback to the resource page (https://vroom.apverma.com). (Please don’t link directly to the .zip files, please link to the resource page.)

'If you should have any questions please contact me here: http://www.apverma.com, alternatively you can mail me at: apverma[at]apverma.com

'-----------------------------------License----------------------------------------------
