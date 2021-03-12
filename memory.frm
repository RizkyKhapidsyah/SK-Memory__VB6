VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memory"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptVery 
      Caption         =   "Very good (17-22)"
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   2640
      Width           =   1575
   End
   Begin VB.OptionButton OptExpert 
      Caption         =   "Expert (23+)"
      Height          =   195
      Left            =   1560
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.OptionButton OptGood 
      Caption         =   "Good (11-16)"
      Height          =   195
      Left            =   1560
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.OptionButton OptKeep 
      Caption         =   "Keep going (0-10)"
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   4080
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1200
      Top             =   4080
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   1680
      Top             =   4080
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   720
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   4080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Memory Level"
      Height          =   195
      Left            =   1800
      TabIndex        =   9
      Top             =   1920
      Width           =   990
   End
   Begin VB.Label Label2 
      Caption         =   "You can use keyboard for 1 - 4 or click it with the mouse for response"
      Height          =   435
      Left            =   840
      TabIndex        =   8
      Top             =   240
      Width           =   2805
   End
   Begin VB.Label LblScore 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Score"
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label Lbl4 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Lbl3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Lbl2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Lbl1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public x
Public Amount As Integer
Public Lit As Integer
Public Score As Integer
Public Answer As String
Public Entry As String
Public OnOff As Boolean
Public LightsDone As Boolean
Public NumOfCalls As Integer
Public GameOver As Boolean
Public BeepYesNo As Boolean
Public AddScore As Boolean

Private Sub CmdExit_Click()
End
End Sub

Private Sub CmdExit_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 49
        Call Lbl1_Click
    Case 50
        Call Lbl2_Click
    Case 51
        Call Lbl3_Click
    Case 52
        Call Lbl4_Click
End Select
End Sub

Private Sub CmdStart_Click()
Score = 0
LblScore = Score
Answer = ""
Entry = ""
Amount = 1
NumOfCalls = 0
CmdStart.Enabled = False
GameOver = False
OnOff = True
Timer2.Enabled = False
LightsDone = False
Call Main
End Sub

Private Sub CmdStart_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 49
        Call Lbl1_Click
    Case 50
        Call Lbl2_Click
    Case 51
        Call Lbl3_Click
    Case 52
        Call Lbl4_Click
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 49
        Call Lbl1_Click
    Case 50
        Call Lbl2_Click
    Case 51
        Call Lbl3_Click
    Case 52
        Call Lbl4_Click
End Select
End Sub

Private Sub Form_Load()
Amount = 1
Lit = 0
LightsDone = False
NumOfCalls = 0
GameOver = True
End Sub

Private Sub Lbl1_Click()
Entry = Entry & "1"
Timer2.Enabled = False
'set up score

Dim a5, b5, c5 As Integer
a5 = Len(Entry)
b5 = Len(Answer)
If a5 > b5 Then
AddScore = False
c5 = a5 - b5
a5 = a5 - c5
Entry2 = Left(Entry, a5)
Entry = Entry2
Else
AddScore = True
End If

If AddScore = True Then
a = Len(Entry)
d = Left(Answer, a)
If d <> Entry Then
MsgBox "Incorrect."
BeepYesNo = False
Call Wrong
Else
BeepYesNo = True
Timer2.Enabled = True
Score = Score + 1
LblScore = Score
End If
End If
Timer2.Enabled = True
End Sub

Private Sub Lbl1_DblClick()
Call Lbl1_Click
End Sub

Private Sub Lbl2_Click()
Entry = Entry & "2"
Timer2.Enabled = False
'set up score

Dim a5, b5, c5 As Integer
a5 = Len(Entry)
b5 = Len(Answer)
If a5 > b5 Then
AddScore = False
c5 = a5 - b5
a5 = a5 - c5
Entry2 = Left(Entry, a5)
Entry = Entry2
Else
AddScore = True
End If

If AddScore = True Then
a = Len(Entry)
d = Left(Answer, a)
If d <> Entry Then
MsgBox "Incorrect."
BeepYesNo = False
Call Wrong
Else
BeepYesNo = True
Timer2.Enabled = True
Score = Score + 1
LblScore = Score
End If
End If
Timer2.Enabled = True
End Sub

Private Sub Lbl2_DblClick()
Call Lbl2_Click
End Sub

Private Sub Lbl3_Click()
Entry = Entry & "3"
Timer2.Enabled = False
'set up score

Dim a5, b5, c5 As Integer
a5 = Len(Entry)
b5 = Len(Answer)
If a5 > b5 Then
AddScore = False
c5 = a5 - b5
a5 = a5 - c5
Entry2 = Left(Entry, a5)
Entry = Entry2
Else
AddScore = True
End If

If AddScore = True Then
a = Len(Entry)
d = Left(Answer, a)
If d <> Entry Then
MsgBox "Incorrect."
BeepYesNo = False
Call Wrong
Else
BeepYesNo = True
Timer2.Enabled = True
Score = Score + 1
LblScore = Score
End If
End If
Timer2.Enabled = True
End Sub

Private Sub Lbl3_DblClick()
Call Lbl3_Click
End Sub

Private Sub Lbl4_Click()
Entry = Entry & "4"
Timer2.Enabled = False
'set up score

Dim a5, b5, c5 As Integer
a5 = Len(Entry)
b5 = Len(Answer)
If a5 > b5 Then
AddScore = False
c5 = a5 - b5
a5 = a5 - c5
Entry2 = Left(Entry, a5)
Entry = Entry2
Else
AddScore = True
End If

If AddScore = True Then
a = Len(Entry)
d = Left(Answer, a)
If d <> Entry Then
MsgBox "Incorrect."
BeepYesNo = False
Call Wrong
Else
BeepYesNo = True
Timer2.Enabled = True
Score = Score + 1
LblScore = Score
End If
End If
Timer2.Enabled = True
End Sub

Private Sub Lbl4_DblClick()
Call Lbl4_Click
End Sub

Private Sub Timer1_Timer()
If OnOff = True Then
If LightsDone = False Then
If Lit > 0 Then
Timer1.Enabled = False

Else
Lit = Lit + 1
Randomize Timer
num = Int(Rnd * 4) + 1
Answer = Answer & num
Select Case num
    Case 1
        Lbl1.BackColor = vbRed
        Lbl2.BackColor = vbBlack
        Lbl3.BackColor = vbBlack
        Lbl4.BackColor = vbBlack
    Case 2
        Lbl1.BackColor = vbBlack
        Lbl2.BackColor = vbYellow
        Lbl3.BackColor = vbBlack
        Lbl4.BackColor = vbBlack
    Case 3
        Lbl1.BackColor = vbBlack
        Lbl2.BackColor = vbBlack
        Lbl3.BackColor = vbBlue
        Lbl4.BackColor = vbBlack
    Case 4
        Lbl1.BackColor = vbBlack
        Lbl2.BackColor = vbBlack
        Lbl3.BackColor = vbBlack
        Lbl4.BackColor = vbGreen
End Select
End If
End If
End If
Timer3.Enabled = True
OnOff = False
Timer6.Enabled = False
End Sub

Private Sub Timer2_Timer()
'this is to analyze response

Dim a, b, c As Integer
a = Len(Entry)
b = Len(Answer)
If a > b Then
c = a - b
a = a - c
Entry2 = Left(Entry, a)
Entry = Entry2
End If

If LightsDone = True And GameOver = False Then
If Entry = Answer Then
Amount = Amount + 1
NumOfCalls = 0
LightsDone = False
Entry = ""
Answer = ""
Call Main
Else
Call Wrong
End If
End If
Entry = ""
Answer = ""
Timer2.Enabled = False
End Sub

Private Sub Wrong()
GameOver = True
Entry = ""
Answer = ""
If BeepYesNo <> False Then
MsgBox "Incorrect."
End If
x = 1
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer6.Enabled = False
CmdStart.Enabled = True
NumOfCalls = 0
Lit = 0
Amount = 1
LightsDone = True
Entry = ""
Answer = ""
End Sub

Private Sub ColorOff()
Lbl1.BackColor = vbBlack
Lbl2.BackColor = vbBlack
Lbl3.BackColor = vbBlack
Lbl4.BackColor = vbBlack
Lit = 0
Timer1.Enabled = True
If NumOfCalls <= Amount Then
Timer6.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
Call ColorOff
End Sub

Private Sub Timer4_Timer()
If NumOfCalls > Amount Then
LightsDone = True
Timer1.Enabled = False
Timer6.Enabled = False
End If
If LightsDone = True And GameOver = False Then
Timer2.Enabled = True
End If
Score2 = Val(LblScore)
If Score2 <= 10 Then
OptKeep.Value = True
ElseIf Score2 >= 11 And Score2 <= 16 Then
OptGood.Value = True
ElseIf Score2 >= 17 And Score2 <= 22 Then
OptVery.Value = True
ElseIf Score2 > 22 Then
OptExpert.Value = True
End If
End Sub

Private Sub Main()
Timer6.Enabled = True
End Sub

Private Sub Timer6_Timer()
NumOfCalls = NumOfCalls + 1
If NumOfCalls <= Amount And GameOver = False Then
LightsDone = False
End If
If NumOfCalls <= Amount And LightsDone = False And GameOver = False Then
OnOff = True
Timer1.Enabled = True
Timer2.Enabled = False
LightsDone = False
Else
Timer2.Enabled = True
End If
Timer6.Enabled = False
End Sub
