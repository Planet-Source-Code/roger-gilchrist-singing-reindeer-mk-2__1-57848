VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl XmasTree 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   MaskColor       =   &H00000000&
   MaskPicture     =   "XmasTree.ctx":0000
   PaletteMode     =   4  'None
   Picture         =   "XmasTree.ctx":1F6AA
   ScaleHeight     =   3585
   ScaleWidth      =   2835
   ToolboxBitmap   =   "XmasTree.ctx":3ED54
   Windowless      =   -1  'True
   Begin VB.Timer tmrXmasTree 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2520
      Top             =   0
   End
   Begin VB.Image imgLight 
      Height          =   135
      Index           =   0
      Left            =   360
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin ComctlLib.ImageList imlXmasTree 
      Left            =   2520
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   5
      ImageHeight     =   6
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":3F066
            Key             =   "blue"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":3F118
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":3F2F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":3F4D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":3F6B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":3F890
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":3FA6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":3FC4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":3FE2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":40008
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":401E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":403C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":405A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":40780
            Key             =   "red"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":409A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "XmasTree.ctx":40B84
            Key             =   "yellow"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "XmasTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum eLight
  Simple
  Standard
  Variable
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Simple, Standard, Variable
#End If
Private m_Test        As Boolean
Private m_Lites       As eLight
Private m_LightUp     As Boolean
Private LiteCount     As Long

Private Sub ColourSet()

  Dim I As Long

  Select Case m_Lites
   Case Simple
    For I = 0 To LiteCount
      imgLight(I).Picture = imlXmasTree.ListImages(Int(Rnd * 3) + 2).Picture
    Next I
   Case Standard
    For I = 0 To LiteCount
      imgLight(I).Picture = imlXmasTree.ListImages(Int(Rnd * imlXmasTree.ListImages.Count) + 1).Picture
    Next I
    'Case Variable
    'Nothing here the Timer does it
  End Select

End Sub

Public Property Get LightMode() As eLight

  LightMode = m_Lites

End Property

Public Property Let LightMode(Lights As eLight)

  m_Lites = Lights
  ColourSet
  PropertyChanged "LMode"

End Property

Public Property Get LightUp() As Boolean

  LightUp = m_LightUp

End Property

Public Property Let LightUp(ByVal vNewValue As Boolean)

  m_LightUp = vNewValue
  tmrXmasTree.Enabled = m_LightUp
  RndLightPos

End Property

Private Sub RndLightPos()

  Dim I  As Long
  Dim Rx As Long
  Dim ry As Long

  For I = 0 To imgLight.Count - 1
    '  With UserControl
    With UserControl
      Do
        Do
          Rx = Rnd * (.Height - imgLight(I).Height)
          ry = Rnd * (.Width - imgLight(I).Width)
        Loop While .Point(ry, Rx) < 100
      Loop While .Point(ry + imgLight(I).Height, Rx) < 100
      '   End With 'UserControl
    End With 'UserControl
    imgLight(I).Top = Rx
    imgLight(I).Left = ry
  Next I

End Sub

Private Property Get Test() As Boolean

  Test = m_Test

End Property

Private Property Let Test(ByVal X As Boolean)

  m_Test = X
  UserControl_Show
  LightUp = m_Test

End Property

Private Sub tmrXmasTree_Timer()

  Dim I As Long

  For I = 0 To imgLight.Count - 1
    imgLight(I).Visible = Rnd > 0.5
    If imgLight(I).Visible Then
      If m_Lites = Variable Then
        imgLight(I).Picture = imlXmasTree.ListImages(Int(Rnd * imlXmasTree.ListImages.Count) + 1).Picture
      End If
    End If
  Next I
  tmrXmasTree.Interval = 200 + Int(Rnd * 400)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  m_Lites = PropBag.ReadProperty("LMode", Simple)

End Sub

Private Sub UserControl_Resize()

  UserControl.Height = 3540
  UserControl.Width = 2895

End Sub

Private Sub UserControl_Show()

  Dim I As Long

  On Error Resume Next
  If Ambient.UserMode Or m_Test Then
    For I = 1 To LiteCount
      Unload imgLight(I)
    Next I
    LiteCount = 15 * Int(Rnd * 5)
    For I = 1 To LiteCount
      Load imgLight(I)
      imgLight(I).Width = 80 + Int(Rnd * 50)
      imgLight(I).Height = 80 + Int(Rnd * 50)
    Next I
    ColourSet
    DoEvents
  End If
  On Error GoTo 0

End Sub

Private Sub UserControl_Terminate()

  tmrXmasTree.Enabled = False

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  PropBag.WriteProperty "LMode", m_Lites, Simple

End Sub

':)Code Fixer V2.8.0 (22/12/2004 2:17:54 AM) 13 + 150 = 163 Lines Thanks Ulli for inspiration and lots of code.
