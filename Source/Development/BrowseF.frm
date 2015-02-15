VERSION 5.00
Begin VB.Form BrowseF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select A Folder - My Movie Manager"
   ClientHeight    =   5235
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4455
   Icon            =   "BrowseF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox flDr 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox flPath 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton bOK 
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
      Left            =   1800
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton bCancel 
      Caption         =   "Cancel"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.DirListBox flBr 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3150
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin VB.CheckBox fR 
      Caption         =   "Look in the Subfolders, their Subfolders etc."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   5055
   End
End
Attribute VB_Name = "BrowseF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub bCancel_Click()
Unload Me
End Sub

Private Sub bOK_Click()
Dim ii
Dim dimStatus
   On Error GoTo bOK_Click_Error

dimStatus = 0
For Each ii In Fwizard.LocData.ListItems
If UCase(ii.Text) = UCase(flBr.path) Then dimStatus = 1
Next

If dimStatus = 0 Then
Fwizard.fPath = flBr.path
Fwizard.fR.Value = fR.Value
Fwizard.fAdd_Click
Else
GoTo msg
End If


dimStatus = 0
For Each ii In addFolders.LocData.ListItems
If UCase(ii.Text) = UCase(flBr.path) Then dimStatus = 1
Next

If dimStatus = 0 Then
addFolders.fPath = flBr.path
addFolders.fR.Value = fR.Value
addFolders.fAdd_Click
Else
GoTo msg
End If

Unload Me
Exit Sub

msg:
MsgBox "This Folder is Already Added. Please Choose another.", vbInformation

   On Error GoTo 0
   Exit Sub

bOK_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure bOK_Click of Form BrowseF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub flBr_Change()
   On Error GoTo flBr_Change_Error

flPath.Text = UCase(flBr.path)

   On Error GoTo 0
   Exit Sub

flBr_Change_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure flBr_Change of Form BrowseF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub

Private Sub flDr_Change()
   On Error GoTo flDr_Change_Error

On Error Resume Next
flBr.path = flDr.Drive

   On Error GoTo 0
   Exit Sub

flDr_Change_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure flDr_Change of Form BrowseF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

flPath = flBr.path
flPath.ForeColor = vbBlack
   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form BrowseF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub


