VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Fwizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movie Manager Setup Wizard"
   ClientHeight    =   6315
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7155
   Icon            =   "Fwizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton nextB 
      Caption         =   "Next >"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cancelB 
      Caption         =   "Cancel"
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
      Left            =   5880
      TabIndex        =   1
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton backB 
      Caption         =   "< Back"
      Enabled         =   0   'False
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
      Left            =   3480
      TabIndex        =   0
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame frWz 
      BackColor       =   &H8000000E&
      Caption         =   "Add the Locations Where You Keep Movie Files"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7095
      Begin VB.PictureBox lFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   120
         ScaleHeight     =   5145
         ScaleWidth      =   6825
         TabIndex        =   4
         Top             =   360
         Width           =   6855
         Begin ComctlLib.ListView LocData 
            Height          =   4455
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   7858
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.CommandButton delList 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            TabIndex        =   6
            Top             =   4680
            Width           =   735
         End
         Begin VB.CommandButton fBrowse 
            Caption         =   "Add location"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   5
            Top             =   4680
            UseMaskColor    =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.TextBox fPath 
         Enabled         =   0   'False
         Height          =   405
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.CheckBox fR 
         BackColor       =   &H8000000E&
         Caption         =   "Look for Movie Files in the Subfolders, their Subfolders etc."
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
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CommandButton fAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5520
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Location :"
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.Frame frWz 
      BackColor       =   &H8000000E&
      Caption         =   "Finished Configuring"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Index           =   2
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7095
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Fwizard.frx":C84A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2400
         TabIndex        =   14
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Image Image2 
         Height          =   4500
         Left            =   120
         Picture         =   "Fwizard.frx":C8E0
         Top             =   600
         Width           =   2220
      End
   End
   Begin VB.Frame frWz 
      BackColor       =   &H8000000E&
      Caption         =   "Welcome To Movie Manger"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7095
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Before getting Started, You will need to Configure Movie manager. Press The Next Button To Continue."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   12
         Top             =   840
         Width           =   4575
      End
      Begin VB.Image Image1 
         Height          =   4500
         Left            =   120
         Picture         =   "Fwizard.frx":14967
         Top             =   600
         Width           =   2220
      End
   End
   Begin VB.Label cnt 
      Caption         =   "0"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4800
      Width           =   255
   End
End
Attribute VB_Name = "Fwizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim isFirstStart As Boolean



Private Sub backB_Click()
   On Error GoTo backB_Click_Error

cnt = cnt - 1
If cnt > 0 Then
backB.Enabled = True
Else
backB.Enabled = False
End If
If cnt > -1 Then

frWz(0).Visible = False
frWz(1).Visible = False
frWz(2).Visible = False
frWz(cnt).Visible = True
End If

   On Error GoTo 0
   Exit Sub

backB_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure backB_Click of Form Fwizard" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub cancelB_Click()

   On Error GoTo cancelB_Click_Error

Unload Me

   On Error GoTo 0
   Exit Sub

cancelB_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure cancelB_Click of Form Fwizard" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub cnt_Change()
   On Error GoTo cnt_Change_Error

If cnt > 1 Then

nextB.Caption = "Finish"
Else
nextB.Caption = "Next >"
End If

   On Error GoTo 0
   Exit Sub

cnt_Change_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure cnt_Change of Form Fwizard" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub



Private Sub delList_Click()
   On Error GoTo delList_Click_Error

If LocData.ListItems.Count > 0 Then
LocData.ListItems.Remove LocData.SelectedItem.Index
LocData.Refresh
End If

   On Error GoTo 0
   Exit Sub

delList_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure delList_Click of Form Fwizard" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Public Sub fAdd_Click()
   On Error GoTo fAdd_Click_Error

If fPath = "" Then
MsgBox "Please Choose a Folder to Add.", vbInformation, "No Folder Chosen"
Else

Dim s As ListItem
Set s = LocData.ListItems.Add(, , fPath)
's.Text = fPath

'lFl.AddItem fPath


If fR.Value Then
'lFr.AddItem "Yes"

s.SubItems(1) = "Yes"
Else
'lFr.AddItem "No"
s.SubItems(1) = "No"
End If
fPath = ""


End If
LocData.Refresh

   On Error GoTo 0
   Exit Sub

fAdd_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure fAdd_Click of Form Fwizard" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub fBrowse_Click()
   On Error GoTo fBrowse_Click_Error

BrowseF.Show 1

   On Error GoTo 0
   Exit Sub

fBrowse_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure fBrowse_Click of Form Fwizard" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub



Private Sub Form_initialize()
   On Error GoTo Form_initialize_Error

isFirstStart = True
LocData.ColumnHeaders.Add , , "Location", LocData.Width - 1500
LocData.ColumnHeaders.Add , , "Recursive"

   On Error GoTo 0
   Exit Sub

Form_initialize_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_initialize of Form Fwizard" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

frWz(0).Visible = True
frWz(1).Visible = False
frWz(2).Visible = False

If Not isFirstStart Then

'Load frmMain_Pg

'frmMain_Pg.Visible = True

Unload Me
End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form Fwizard" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub






Private Sub nextB_Click()
   On Error GoTo nextB_Click_Error

If cnt = 1 Then
If Not LocData.ListItems.Count > 0 Then
MsgBox "Please add atleast one Location", vbInformation, "No Locations Added"
GoTo skip
End If
End If
cnt = cnt + 1
If cnt > -1 And cnt < 3 Then

frWz(0).Visible = False
frWz(1).Visible = False
frWz(2).Visible = False
frWz(cnt).Visible = True
End If
If cnt > 0 Then
backB.Enabled = True
Else
backB.Enabled = False
End If
If cnt > 2 Then

Dim sk As FileSystemObject
Set sk = New FileSystemObject
Dim teS As String
teS = "<Locations>" & vbNewLine
Dim i, jj
For Each i In LocData.ListItems

teS = teS & "    <Location rec=" & Chr(34) & i.SubItems(1) & Chr(34) & ">" & i.Text & "</Location>" & vbNewLine

Next
teS = teS & vbNewLine & "</Locations>"
Set jj = sk.OpenTextFile(App.path & "\Data\Loc.xml", ForWriting, True, TristateUseDefault)
jj.WriteLine teS
jj.Close

Dim flDr As New FileSystemObject
If flDr.FileExists(App.path & "\Data\FirstRun - Copy.xml") Then flDr.DeleteFile App.path & "\Data\FirstRun - Copy.xml"

'updateFiles.updateFileList
Unload Me
createDB.Show







End If


skip:

   On Error GoTo 0
   Exit Sub

nextB_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure nextB_Click of Form Fwizard" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub

