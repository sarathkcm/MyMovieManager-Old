VERSION 5.00
Begin VB.Form aboutFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3600
   ClientLeft      =   2760
   ClientTop       =   3420
   ClientWidth     =   5475
   ControlBox      =   0   'False
   Icon            =   "aboutFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton okB 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "I really appreciate the help offered by my friend, Nikhil SB in testing the software..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   5115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Software uses www.opensubtitles.org API to identify and download Movie Details..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   5115
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by Sarath KCM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   120
      Picture         =   "aboutFrm.frx":C84A
      Top             =   240
      Width           =   4500
   End
   Begin VB.Label vInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   600
   End
End
Attribute VB_Name = "aboutFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Initialize()
vInfo = " Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Image1_Click()
ShellExecute Me.hWnd, "Open", "http://experimenter.x10.mx", 0&, 0&, 0&
End Sub

Private Sub Label1_Click()
ShellExecute Me.hWnd, "Open", "http://www.opensubtitles.org", 0&, 0&, 0&
End Sub

Private Sub okB_Click()
Unload Me
End Sub
