VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Singing Reindeer mk 2 (Right-Click sky to stop music & enjoy the singing)"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7740
   Icon            =   "reindeer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7200
      Top             =   120
   End
   Begin Project1.dance dance3 
      Height          =   3570
      Left            =   960
      TabIndex        =   12
      Top             =   1680
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   6297
   End
   Begin Project1.dance ctlDancer2 
      Height          =   3570
      Left            =   2640
      TabIndex        =   11
      Top             =   1680
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   6297
   End
   Begin Project1.XmasTree ctlXTree 
      Height          =   3540
      Index           =   5
      Left            =   3360
      TabIndex        =   10
      Top             =   1920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6244
      LMode           =   2
   End
   Begin Project1.XmasTree ctlXTree 
      Height          =   3540
      Index           =   2
      Left            =   6000
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6244
   End
   Begin Project1.XmasTree ctlXTree 
      Height          =   3540
      Index           =   1
      Left            =   -600
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6244
      LMode           =   2
   End
   Begin Project1.fence ctlFence 
      Height          =   2175
      Left            =   -120
      TabIndex        =   0
      Top             =   1800
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3836
   End
   Begin Project1.dixon ctlDixon1 
      Height          =   3570
      Left            =   4680
      TabIndex        =   9
      Top             =   480
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   6297
   End
   Begin Project1.XmasTree ctlXTree 
      Height          =   3540
      Index           =   4
      Left            =   4560
      TabIndex        =   6
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6244
   End
   Begin Project1.rudolph ctlRudolph 
      Height          =   3585
      Left            =   3600
      TabIndex        =   8
      Top             =   0
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   6324
   End
   Begin Project1.dance ctlDancer1 
      Height          =   3570
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   6297
   End
   Begin Project1.XmasTree ctlXTree 
      Height          =   3540
      Index           =   3
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6244
   End
   Begin Project1.blitz blitz1 
      Height          =   3570
      Left            =   2280
      TabIndex        =   1
      Top             =   -240
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   6297
   End
   Begin Project1.XmasTree ctlXTree 
      Height          =   3540
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6244
      LMode           =   2
   End
   Begin Project1.blitz blitz2 
      Height          =   3570
      Left            =   240
      TabIndex        =   13
      Top             =   -840
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   6297
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Singing Reindeer   - Santa 2004 -
'
'Click on each individual reindeer to start them singing,
'click on them again, to quiet them.
'
'A special thankyou to Mr. Paul Turkson.  ;)
'
'Peace
'
Dim killTimer As Long
Private midiOn                   As Boolean
Private Const cSoundBuffers      As Long = 5
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, _
                                                                               ByVal lpstrReturnString As String, _
                                                                               ByVal uReturnLength As Long, _
                                                                               ByVal hwndCallback As Long) As Long

Private Sub Form_Load()

  Dim I As Long
Randomize Timer
  'put most reindeer off stage
   ctlRudolph.Left = ctlRudolph.Left + 10000 + Int(Rnd * 2000)
  blitz2.Left = blitz2.Left + -10000 + Int(Rnd * 4000)
  ctlDancer1.Left = ctlDancer1.Left - 10000 - Int(Rnd * 3000)
  ctlDancer2.Left = ctlDancer2.Left - 10000 - Int(Rnd * 2000)
  dance3.Left = dance3.Left - 10000 - Int(Rnd * 10000)
  ctlDixon1.Left = ctlDixon1.Left + 10000 + Int(Rnd * 4500)
  Me.Show
  DoEvents
  SetupDX7Sound Me
  'sound file path
  SoundDir App.Path
  'default file
  CreateBuffers cSoundBuffers, "rudolph.wav"
  'start the music
  midiOn = True
  mciSendString "play " & "mid005.mid", 0&, 0, 0
  '
  'wake-up the reindeer
  ctlRudolph.StartMe Rnd > 0.5
  blitz1.StartMe True
  blitz2.StartMe Rnd > 0.5
  ctlDixon1.StartMe Rnd > 0.5
  ctlDancer1.StartMe Rnd > 0.5
  ctlDancer2.StartMe Rnd > 0.5
  dance3.StartMe Rnd > 0.5
  For I = 0 To ctlXTree.Count - 1
    ctlXTree(I).LightUp = True
    ctlXTree(I).LightMode = Int(Rnd * 3)
  Next I

End Sub

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

  If Button = vbRightButton Then
    If midiOn Then
      mciSendString "close " & "mid005.mid", 0&, 0, 0
      midiOn = False
     Else
      mciSendString "play " & "mid005.mid", 0&, 0, 0
      midiOn = True
    End If
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

  mciSendString "close " & "mid005.mid", 0&, 0, 0
  End

End Sub

Private Sub MoveDeer(ctrl As Variant, _
                     ByVal stopPoint As Long, _
                     ByVal lngDir As Long, _
                     Finished As Long)



  If lngDir > 0 Then
    If ctrl.Left < stopPoint Then
      ctrl.Left = stopPoint
      Finished = Finished + 1
     ElseIf ctrl.Left > stopPoint Then
      ctrl.Left = ctrl.Left - 100
    End If
   Else
    If ctrl.Left > stopPoint Then
      ctrl.Left = stopPoint
      Finished = Finished + 1
     ElseIf ctrl.Left < stopPoint Then
      ctrl.Left = ctrl.Left + 100
    End If
  End If

End Sub

Private Sub Timer1_Timer()

  

  MoveDeer ctlDixon1, 4680, 1, killTimer
  MoveDeer ctlRudolph, 3600, 1, killTimer
  MoveDeer blitz2, 240, -1, killTimer
  MoveDeer ctlDancer1, 240, -1, killTimer
  MoveDeer ctlDancer2, 2640, -1, killTimer
  MoveDeer dance3, 960, -1, killTimer
   If killTimer = 6 Then
    Timer1.Enabled = False
  End If

End Sub

':)Code Fixer V2.8.0 (22/12/2004 2:17:45 AM) 13 + 106 = 119 Lines Thanks Ulli for inspiration and lots of code.

