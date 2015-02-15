VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form TestForm 
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15015
   Icon            =   "TestForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   15015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   7575
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "TestForm.frx":11FF2
      Top             =   720
      Width           =   8775
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   10320
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   1095
      Left            =   9960
      TabIndex        =   3
      Top             =   -120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   7200
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hsh As String
Private Sub Command1_Click()
Dim log() As String
Dim hashS() As String
Dim mov() As String
hsh = find_Hash("D:\Movies\ForTest\Alien - Movie Sequel\Alien Directors Cut [1979]\Alien Director's Cut 1979.720p.BrRip.x264.YIFY.mp4")
log = loginXml
hashS = checkMovieHashXml(log(3), log(1), hsh)
mov = getIMDBdetailsXml(log(3), log(1), hashS(1))
Dim mmm
For Each mmm In mov
Text1 = Text1 & vbCrLf & mmm
Next
End Sub
