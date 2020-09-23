VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl Mouth 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1035
   HasDC           =   0   'False
   MaskColor       =   &H00000000&
   MaskPicture     =   "Mouth.ctx":0000
   Picture         =   "Mouth.ctx":08A2
   ScaleHeight     =   900
   ScaleWidth      =   1035
   Windowless      =   -1  'True
   Begin ComctlLib.ImageList imlMouth 
      Left            =   120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   31
      ImageHeight     =   35
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":1144
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":1CA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":2800
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":30B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":3964
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":428E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":4BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":5A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":696C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":6EFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":7490
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Mouth.ctx":7AC2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Mouth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum Rname
  rBlitzen = 0
  rRDixon = 4
  rRoudolph = 8
  rDancer = 10
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private rBlitzen, rRDixon, rRoudolph, rDancer
#End If
Private curMouth       As Long
Private m_MouthSet     As Rname

Public Property Get MouthSet() As Rname

  MouthSet = m_MouthSet

End Property

Public Property Let MouthSet(Reindeer As Rname)

  m_MouthSet = Reindeer
  PropertyChanged "useMouth"
  State = 0

End Property

Public Property Get State() As Long

  State = curMouth

End Property

Public Property Let State(ByVal mouthNumber As Long)

  curMouth = mouthNumber
  PropertyChanged "MouthState"
  Select Case curMouth
   Case 0
    UserControl.Picture = imlMouth.ListImages(m_MouthSet + 2).Picture
    UserControl.MaskPicture = imlMouth.ListImages(m_MouthSet + 1).Picture
   Case 1
    UserControl.Picture = imlMouth.ListImages(m_MouthSet + 4).Picture
    UserControl.MaskPicture = imlMouth.ListImages(m_MouthSet + 3).Picture
  End Select

End Property

Private Sub UserControl_InitProperties()

  State = 0
  MouthSet = rBlitzen

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  MouthSet = PropBag.ReadProperty("useMouth", rBlitzen)
  State = PropBag.ReadProperty("MouthState", 0)

End Sub

Private Sub UserControl_Resize()

  If m_MouthSet + curMouth = 0 Then
    Width = imlMouth.ListImages(1).Picture.Width
    Height = imlMouth.ListImages(1).Picture.Height
   Else
    Width = imlMouth.ListImages(m_MouthSet + curMouth).Picture.Width
    Height = imlMouth.ListImages(m_MouthSet + curMouth).Picture.Height
  End If
  DoEvents

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  PropBag.WriteProperty "useMouth", m_MouthSet, rBlitzen
  PropBag.WriteProperty "MouthState", curMouth, 0

End Sub

':)Code Fixer V2.8.0 (22/12/2004 2:17:51 AM) 12 + 72 = 84 Lines Thanks Ulli for inspiration and lots of code.
