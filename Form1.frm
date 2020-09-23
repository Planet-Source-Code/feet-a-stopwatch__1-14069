VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H80000012&
   Caption         =   "StopWatch"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1605
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   1605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pause For:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pause"
      Height          =   255
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   950
      Left            =   840
      Top             =   1320
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[[[[[[[[[[||||||]]]]]]]]]]
'[[[[StopWatch By Feet]]]]]
'[[[[[[[[[[||||||]]]]]]]]]]
'this is my first example i have posted on PSC, prepare to see many more!
'if u like this please vote for it at psc
'u can use this module/zip file for anythhing, just don't sell it or make a profit off it please!
'i could have added much much much much more to this but i wanted to keep it simple for a user that has practically never dealt with vb
'if u post enough feedback i may post a better one
'any feedback is appreciated good or bad
'to vote or put feedback, goto www.planetsourcecode.com and goto visual basic section
'then search for me, feet and look for this project

Private Sub Command1_Click()
Timer1.Enabled = False
'Pauses the timer(StopWatch)
End Sub

Private Sub Command2_Click()
On Error GoTo away
Dim HowLong As Integer
Timer1.Enabled = False
HowLong = InputBox("Please enter how long you could like to pause for ***IN SECONDS***", "Pause For How Long?", "10")
Pause (HowLong)
Timer1.Enabled = True
away:
'asks long to pause with with an input box then pauses
'for that amount of time with the timer not enabled then it turns the timer(StopWatch) back on
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True
'start the stopwatch
End Sub

Private Sub Command4_Click()
Call Restart(Label1)
Timer1.Enabled = False
'calls to restart the label and disable the timer(StopWatch)
End Sub

Private Sub Command5_Click()
MsgBox "Programmed By Feet On Wednesday, January 3rd, 2001 On Visual Basic 6.0 Happy New Year. Happy REAL millenium.", vbInformation, "About This"
End Sub

Private Sub Timer1_Timer()
Call StopWatch.StopWatch(Label1)
End Sub

Private Sub Timer2_Timer()
Label2.Caption = Time
End Sub

Private Sub Timer3_Timer()
Label3.Caption = Label3.Caption + 1
End Sub
