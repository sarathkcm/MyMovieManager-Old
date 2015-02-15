VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form addFolders 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Remove Scanned Folders - My Movie Manager"
   ClientHeight    =   6210
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7125
   Icon            =   "addFolders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton saveNscan 
      Caption         =   "Save and Scan Folders"
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
      Left            =   1920
      TabIndex        =   11
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   2535
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
      TabIndex        =   2
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
         TabIndex        =   3
         Top             =   360
         Width           =   6855
         Begin ComctlLib.ListView LocData 
            Height          =   4455
            Left            =   120
            TabIndex        =   10
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
            TabIndex        =   4
            Top             =   4680
            Width           =   735
         End
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
         TabIndex        =   7
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
         Left            =   5640
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox fPath 
         Enabled         =   0   'False
         Height          =   405
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   5175
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
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.CommandButton OkB 
      Caption         =   "Save"
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
      Left            =   4560
      TabIndex        =   1
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cancelB 
      Caption         =   "Cancel"
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
      Left            =   5880
      TabIndex        =   0
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
End
Attribute VB_Name = "addFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit




Private Sub cancelB_Click()

   On Error GoTo cancelB_Click_Error

Unload Me

   On Error GoTo 0
   Exit Sub

cancelB_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure cancelB_Click of Form addFolders" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

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

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure delList_Click of Form addFolders" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

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

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure fAdd_Click of Form addFolders" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub fBrowse_Click()
   On Error GoTo fBrowse_Click_Error

BrowseF.Show 1

   On Error GoTo 0
   Exit Sub

fBrowse_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure fBrowse_Click of Form addFolders" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub



Private Sub Form_init()

   On Error GoTo Form_init_Error

LocData.ColumnHeaders.Add , , "Location", LocData.Width - 1500
LocData.ColumnHeaders.Add , , "Recursive"



Dim sku As New ChilkatXml
sku.LoadXmlFile App.path & "\Data\Loc.xml"
Dim i
Dim skuC As New ChilkatXml
Dim s As ListItem
For i = 0 To sku.NumChildrenHavingTag("Location") - 1

Set skuC = sku.GetNthChildWithTag("Location", i)

'MsgBox skuC.Content
Set s = LocData.ListItems.Add(, , skuC.Content)
s.SubItems(1) = skuC.GetAttrValue("rec")

Next

   On Error GoTo 0
   Exit Sub

Form_init_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_init of Form addFolders" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub






Private Sub Form_Load()
   On Error GoTo Form_Load_Error

Form_init

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form addFolders" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub



Private Sub OkB_Click()

   On Error GoTo OkB_Click_Error

If Not LocData.ListItems.Count > 0 Then
MsgBox "Please add atleast one Location", vbInformation, "No Locations Added"
GoTo skip
End If



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









skip:

   On Error GoTo 0
   Exit Sub

OkB_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure OkB_Click of Form addFolders" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub saveNscan_Click()
   On Error GoTo saveNscan_Click_Error

If Not LocData.ListItems.Count > 0 Then
MsgBox "Please add atleast one Location", vbInformation, "No Locations Added"
GoTo skip
End If



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


mainF.scn_Click
Unload Me














skip:

   On Error GoTo 0
   Exit Sub

saveNscan_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure saveNscan_Click of Form addFolders" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub
