VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form updateFromNetForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000002&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Updating Movie Details From Internet - My Movie Manager"
   ClientHeight    =   2160
   ClientLeft      =   5445
   ClientTop       =   5265
   ClientWidth     =   10590
   ControlBox      =   0   'False
   Icon            =   "updateFromNetForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar prgs 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.CommandButton stpButton 
      Caption         =   "Stop"
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
      Left            =   8880
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5640
      Top             =   1560
   End
   Begin VB.Label errorMSG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   2
      Top             =   240
      Width           =   780
   End
   Begin VB.Label stts 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Top             =   1320
      Width           =   780
   End
End
Attribute VB_Name = "updateFromNetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit




Private Sub stpButton_Click()
   On Error GoTo stpButton_Click_Error

stopSignalNet = True
Unload Me

   On Error GoTo 0
   Exit Sub

stpButton_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure stpButton_Click of Form updateFromNetForm" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub Timer1_Timer()
   On Error GoTo Timer1_Timer_Error

Timer1.Interval = 0
Call updateFromNet(True, stts, Label1, prgs, errorMSG)
Unload Me

   On Error GoTo 0
   Exit Sub

Timer1_Timer_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer1_Timer of Form updateFromNetForm" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub
