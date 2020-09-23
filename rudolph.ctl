VERSION 5.00
Begin VB.UserControl rudolph 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   MaskColor       =   &H000000FF&
   MaskPicture     =   "rudolph.ctx":0000
   Picture         =   "rudolph.ctx":97C2
   ScaleHeight     =   3585
   ScaleWidth      =   2385
   Windowless      =   -1  'True
   Begin VB.Timer singTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   2160
   End
   Begin VB.Timer eyeTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   2160
   End
   Begin Project1.EyeLids EyeLids1 
      Height          =   315
      Left            =   1050
      TabIndex        =   1
      Top             =   1250
      Width           =   405
      _extentx        =   847
      _extenty        =   661
      useeye          =   22
   End
   Begin Project1.Mouth Mouth1 
      Height          =   825
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   1065
      _extentx        =   1879
      _extenty        =   1455
      usemouth        =   8
   End
End
Attribute VB_Name = "rudolph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private bSing             As Boolean
Private m_Song            As String
Private lngEyeState       As Long
Private lngMouthState     As Long

Private Sub eyeTimer_Timer()

  Select Case lngEyeState
   Case 1
    EyeLids1.State = Int(Rnd * 4) + 1
    eyeTimer.Interval = Int(10001 * Rnd) + 200
   Case 2, 3, 4, 5
    EyeLids1.State = Int(Rnd * 5) + 1
    eyeTimer.Interval = Int(201 * Rnd) + 200
   Case 6
    EyeLids1.State = eClose
    eyeTimer.Interval = 500
  End Select
  DoEvents
  lngEyeState = lngEyeState + 1
  If lngEyeState > 6 Then
    lngEyeState = 1
  End If

End Sub

Private Sub singTimer_Timer()

  Select Case lngMouthState
   Case 0
    Mouth1.Visible = True
    Mouth1.State = 0
    singTimer.Interval = 1000
    PlaySoundAnyBuffer m_Song, 100
   Case 1
    Mouth1.Visible = True
    singTimer.Interval = 500
   Case 2
    Mouth1.Visible = True
    singTimer.Interval = 2000
   Case 3
    Mouth1.Visible = False
    If bSing Then
      singTimer.Interval = Int(2001 * Rnd + 2000)
     Else
      singTimer.Interval = 10
      singTimer.Enabled = False
    End If
  End Select
  lngMouthState = lngMouthState + 1
  If lngMouthState > 3 Then
    lngMouthState = 0
  End If

End Sub

Public Property Get Song() As String

  Song = m_Song

End Property

Public Property Let Song(ByVal fileName As String)

  m_Song = fileName
  PropertyChanged "Song"

End Property

Public Sub StartMe(Optional ByVal bStartSinging As Boolean = False)


  eyeTimer.Enabled = True
  Mouth1.Visible = False
  Mouth1.State = 0
  If bStartSinging Then
    bSing = Not bSing
    If bSing Then
      singTimer.Enabled = True
    End If
  End If

End Sub

Private Sub UserControl_InitProperties()

  EyeLids1.Visible = True
  m_Song = "rudolph.wav"
  lngEyeState = 1

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

  bSing = Not bSing
  If bSing Then
    singTimer.Enabled = True
  End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  m_Song = PropBag.ReadProperty("Song", "rudolph.wav")

End Sub

Private Sub UserControl_Resize()

  Width = 2385
  Height = 3585

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  PropBag.WriteProperty "Song", m_Song, "rudolph.wav"

End Sub

':)Code Fixer V2.8.0 (22/12/2004 2:17:46 AM) 5 + 119 = 124 Lines Thanks Ulli for inspiration and lots of code.

