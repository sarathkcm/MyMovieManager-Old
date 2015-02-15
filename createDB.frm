VERSION 5.00
Begin VB.Form createDB 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creating Database..."
   ClientHeight    =   2580
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7500
   ControlBox      =   0   'False
   Icon            =   "createDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   7215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6840
      Top             =   240
   End
   Begin VB.Label st 
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "createDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim ms As String

Private Sub Form_Load()
ms = ""

End Sub



Private Sub st_Change()
   On Error GoTo st_Change_Error

ms = ms & st.Caption & vbCrLf
Text1 = ms

   On Error GoTo 0
   Exit Sub

st_Change_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure st_Change of Form createDB" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub Text1_Change()
   On Error GoTo Text1_Change_Error

Text1.SelStart = Len(Text1)

   On Error GoTo 0
   Exit Sub

Text1_Change_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Text1_Change of Form createDB" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
   On Error GoTo Timer1_Timer_Error

Timer1.Interval = 0
Call updateFileList(st)
Call mainF.init(st)
MsgBox "Database Created at Said Locations. Select Menu > Update Details From IMDb to Fetch Movie Details From Internet", vbInformation
mainF.Show
Unload Me



   On Error GoTo 0
   Exit Sub

Timer1_Timer_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer1_Timer of Form createDB" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub


