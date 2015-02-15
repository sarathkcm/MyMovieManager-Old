VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form mainF 
   Caption         =   "My Movie Manager"
   ClientHeight    =   9510
   ClientLeft      =   3165
   ClientTop       =   1575
   ClientWidth     =   15120
   Icon            =   "mainF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10650
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar topBar 
      Align           =   1  'Align Top
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   1402
      ButtonWidth     =   1482
      ButtonHeight    =   1244
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "buttons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   7600
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sort"
            Key             =   "sorteR"
            Object.ToolTipText     =   "Sort in Ascending or Descending order"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sort By"
            Key             =   "sortBy"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sep"
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   800
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Play"
            Key             =   "openMovie"
            Object.ToolTipText     =   "Play The Selected Movie In Default Video Player"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Folder"
            Key             =   "openFolder"
            Object.ToolTipText     =   "Open Folder Containing This Movie"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Edit"
            Key             =   "editMovie"
            Object.ToolTipText     =   "Edit the Details of the Selected Movie"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Remove"
            Key             =   "isNoMovie"
            Object.ToolTipText     =   "Remove This Movie From  Database"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Watched"
            Key             =   "movieWatched"
            Object.ToolTipText     =   "Mark This Movie as Watched"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Favourite"
            Key             =   "addFav"
            Object.ToolTipText     =   "Add/Remove This Movie From Favourites"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "IMDb"
            Key             =   "goToImdb"
            Object.ToolTipText     =   "See this Movie at www.imdb.com"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.OptionButton nWatched 
         Caption         =   "Not Watched"
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
         Left            =   6120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   490
         Width           =   1335
      End
      Begin VB.OptionButton Watched 
         Caption         =   "Watched"
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
         Left            =   6120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton All 
         Caption         =   "All Movies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   75
         ScaleHeight     =   255
         ScaleWidth      =   2700
         TabIndex        =   5
         Top             =   550
         Visible         =   0   'False
         Width           =   2700
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type and Press Enter to Search"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   165
            TabIndex        =   6
            Top             =   0
            Width           =   2310
         End
      End
      Begin VB.ComboBox Genre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "mainF.frx":C84A
         Left            =   4440
         List            =   "mainF.frx":C84C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   150
         Width           =   1575
      End
      Begin VB.ComboBox searchBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Select "
         Top             =   150
         Width           =   1455
      End
      Begin VB.TextBox searchBox 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   75
         TabIndex        =   2
         Top             =   150
         Width           =   2700
      End
   End
   Begin ComctlLib.ListView movieHolder 
      Height          =   10230
      Left            =   120
      TabIndex        =   0
      Top             =   810
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   18045
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
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
   Begin VB.Frame movieDetails 
      Caption         =   "Movie Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Left            =   9360
      TabIndex        =   7
      Top             =   810
      Width           =   10935
      Begin ComctlLib.Slider ratinG 
         Height          =   270
         Left            =   1680
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   9120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   476
         _Version        =   327682
         LargeChange     =   1
      End
      Begin RichTextLib.RichTextBox movieDispPanel 
         Height          =   7335
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   12938
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"mainF.frx":C84E
      End
      Begin VB.Label rateVal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "( 0 / 10 )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4920
         TabIndex        =   15
         Top             =   9120
         Width           =   735
      End
      Begin VB.Label rateThis 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RateThis Movie : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   9120
         Width           =   1635
      End
   End
   Begin VB.Label statusS 
      Height          =   735
      Left            =   9240
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   135
   End
   Begin ComctlLib.ImageList forSmallIcons 
      Left            =   9360
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList buttons 
      Left            =   9360
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   -2147483641
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":C8D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":CFDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":D6EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":DDFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":E508
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":EC16
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":F324
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":FA32
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":10140
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":1084E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mainF.frx":10F5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList forIcons 
      Left            =   9360
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnu 
      Caption         =   "&Menu"
      Begin VB.Menu addRem 
         Caption         =   "&Add / Remove Wached Folders"
      End
      Begin VB.Menu sp 
         Caption         =   "-"
      End
      Begin VB.Menu scn 
         Caption         =   "&Scan Watched Folders"
      End
      Begin VB.Menu upImdb 
         Caption         =   "&Update Details From IMDb"
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu ext 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu sortMenu 
      Caption         =   "&Sort"
      Begin VB.Menu srtBy 
         Caption         =   "Sort &By"
         Begin VB.Menu srtDateAdded 
            Caption         =   "Date Added"
            Checked         =   -1  'True
         End
         Begin VB.Menu srtTitle 
            Caption         =   "&Title"
            Checked         =   -1  'True
         End
         Begin VB.Menu srtYear 
            Caption         =   "&Year"
            Checked         =   -1  'True
         End
         Begin VB.Menu srtImdb 
            Caption         =   "&IMDb Rating"
            Checked         =   -1  'True
         End
         Begin VB.Menu srtMy 
            Caption         =   "Your &Rating"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu srtOrder 
         Caption         =   "Sort &Order"
         Begin VB.Menu srtAsc 
            Caption         =   "&Ascending"
            Checked         =   -1  'True
         End
         Begin VB.Menu srtDec 
            Caption         =   "&Descending"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "&Help"
      Begin VB.Menu abt 
         Caption         =   "&About"
      End
      Begin VB.Menu wb 
         Caption         =   "&Website"
      End
   End
End
Attribute VB_Name = "mainF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lastInd As Long
Dim genreIsReady As Boolean
Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const WM_PASTE = &H302



Function displayMovies()

Dim numMovie As Long
   On Error GoTo displayMovies_Error

movieHolder.ListItems.Clear
Dim lItem, lItem2 As ListItem


Dim iHaHa

For iHaHa = 0 To 6
    movieHolder.ColumnHeaders.Add , , str(iHaHa)
Next







numMovie = UBound(movieArray) - LBound(movieArray)
'For i = 1 To numMovie
'    movieArray(i).MovieSearchRelevance = -1
'Next


'Pre-Adding List
movieHolder.Icons = buttons
forSmallIcons.ListImages.Clear
forSmallIcons.ListImages.Add , , LoadPicture(App.path & "\Images\defaultCover.jpg")
movieHolder.Icons = forSmallIcons

forIcons.ListImages.Clear

Dim searchI As Long





'Adding List

For searchI = LBound(movieArray) To UBound(movieArray) - 1

    If (Not movieArray(searchI).MovieSearchRelevance = 0) And movieArray(searchI).MovieDisplayFlag = True And movieArray(searchI).isMovie = "Yes" And movieArray(searchI).wacthedCategory = True Then

        Set lItem2 = movieHolder.ListItems.Add

                    lItem2.Text = movieArray(searchI).MovieTitle
                    
                    lItem2.SubItems(1) = movieArray(searchI).MovieYear
                    
                    lItem2.SubItems(2) = movieArray(searchI).MovieIMDbRating
                    
                    lItem2.SubItems(3) = movieArray(searchI).MovieMyRating
                    
                    lItem2.SubItems(4) = movieArray(searchI).MovieSearchRelevance
                    
                    lItem2.SubItems(5) = movieArray(searchI).movieIdentifier
                    
                    lItem2.SubItems(6) = movieArray(searchI).DateAdded
                    
                    Set lItem = forIcons.ListImages.Add(, , movieArray(searchI).MovieIcon)
                    
                    movieHolder.Icons = forIcons
                    lItem2.Icon = lItem.Index
                
    End If

Next

movieHolder.Arrange = lvwAutoLeft
movieHolder.Arrange = lvwAutoTop
movieHolder.Refresh


movieHolder.Sorted = True
movieDispPanel.SelAlignment = rtfCenter
movieDispPanel.TextRTF = ""
movieDispPanel.Font = "Arial"
movieDispPanel.SelFontSize = 14
movieDispPanel.SelColor = vbRed

movieDispPanel.SelFontName = "Arial"
movieDispPanel.SelText = "No Movies To Display"
topBar.buttons(9).Value = tbrUnpressed

If movieHolder.ListItems.Count > 0 Then
movieHolder.SelectedItem = movieHolder.ListItems(1)
movieHolder.SelectedItem.EnsureVisible
movieHolder_ItemClick movieHolder.ListItems(1)
End If


movieHolder.Refresh
'movieHolder.SetFocus

'''''''''''
movieHolder.Sorted = True
movieHolder.SortKey = 6
movieHolder.SortOrder = lvwDescending
If movieHolder.SortOrder = lvwDescending Then
topBar.buttons(2).Image = 1
srtDec.Checked = True
srtAsc.Checked = False
Else
topBar.buttons(2).Image = 2
srtDec.Checked = False
srtAsc.Checked = True
End If


srtTitle.Checked = False
srtYear.Checked = False
srtImdb.Checked = False
srtMy.Checked = False
srtDateAdded.Checked = True



   On Error GoTo 0
   Exit Function

displayMovies_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure displayMovies of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function

















Private Sub abt_Click()
aboutFrm.Show 1
End Sub

Private Sub addRem_Click()
   On Error GoTo addRem_Click_Error

addFolders.Show 1

   On Error GoTo 0
   Exit Sub

addRem_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure addRem_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub All_Click()
   On Error GoTo All_Click_Error

If All.Value = True Then
Dim i

For i = LBound(movieArray) To UBound(movieArray)
movieArray(i).wacthedCategory = True
Next
End If
displayMovies

   On Error GoTo 0
   Exit Sub

All_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure All_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub ext_Click()
   On Error GoTo ext_Click_Error

Unload Me

   On Error GoTo 0
   Exit Sub

ext_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure ext_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub Form_initialize()
   On Error GoTo Form_initialize_Error

searchBox.Text = "Search Movies"
searchBox.Alignment = vbCenter

searchBox.FontItalic = True
searchBox.ForeColor = RGB(100, 100, 100)

srtTitle.Checked = True
srtYear.Checked = False
srtImdb.Checked = False
srtMy.Checked = False
srtAsc.Checked = True
srtDec.Checked = False

   On Error GoTo 0
   Exit Sub

Form_initialize_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_initialize of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

searchBy.AddItem "Title"
searchBy.AddItem "Director"
searchBy.AddItem "Cast"
searchBy.AddItem "Plot"
searchBy.AddItem "Year"

searchBy.AddItem "Language"
searchBy.AddItem "Country"
searchBy.AddItem "FileName"

searchBy.ListIndex = 0

HideCaret movieDispPanel.hWnd

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub Form_Resize()
   On Error GoTo Form_Resize_Error

On Error Resume Next
If Me.Width / Screen.TwipsPerPixelX < 1024 Then Me.Width = 1024 * Screen.TwipsPerPixelX
If Me.Height / Screen.TwipsPerPixelY < 700 Then Me.Height = 700 * Screen.TwipsPerPixelY

Dim mW, kk
mW = mainF.Width / 2 + 700
kk = mW / (Screen.TwipsPerPixelX * 148)
kk = Round(kk, 0)
mW = kk * (Screen.TwipsPerPixelX * 148) + 200
movieHolder.Move movieHolder.Left, movieHolder.Top, mW, Me.Height - 1750
movieDetails.Move movieHolder.Left + movieHolder.Width + 200, movieDetails.Top, Me.Width - (movieHolder.Left + movieHolder.Width + 600), Me.Height - 1750
movieDispPanel.Move 150, 300, movieDetails.Width - 300, movieDetails.Height - 650
rateThis.Move 150, movieDetails.Height - 100 - rateThis.Height
ratinG.Move 200 + rateThis.Width, rateThis.Top, movieDetails.Width - (200 + rateThis.Width + 1000)
rateVal.Move ratinG.Left + 100 + ratinG.Width, rateThis.Top
topBar.buttons(4).Width = movieHolder.Width - topBar.buttons(4).Left + 200

   On Error GoTo 0
   Exit Sub

Form_Resize_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Resize of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

Cancel = 1
End
Cancel = 0

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub Genre_Click()
   On Error GoTo Genre_Click_Error

On Error Resume Next
Dim igen

If genreIsReady Then
If Genre.Text = "------------" Then
    Genre.ListIndex = lastInd
Else

    lastInd = Genre.ListIndex


    Dim searchI As Long
    Dim matchFound As Boolean

    For searchI = LBound(movieArray) To UBound(movieArray) - 1


        matchFound = False
        For igen = LBound(movieArray(searchI).MovieGenre) To UBound(movieArray(searchI).MovieGenre)
            If Genre.Text = movieArray(searchI).MovieGenre(igen) Then matchFound = True
        Next


        If Genre.Text = "All" Then
            matchFound = True
        ElseIf Genre.Text = "Watched" And movieArray(searchI).MovieWatched = "Yes" Then
            matchFound = True
        ElseIf Genre.Text = "Not Watched" And movieArray(searchI).MovieWatched = "No" Then
            matchFound = True
        ElseIf Genre.Text = "No Category" And UBound(movieArray(searchI).MovieGenre) = 0 Then
            matchFound = True
        ElseIf Genre.Text = "Favourites" And movieArray(searchI).MovieIsFav = "Yes" Then
            matchFound = True
        
        End If

        movieArray(searchI).MovieDisplayFlag = matchFound
    Next
    displayMovies
End If
End If
movieHolder.SetFocus

   On Error GoTo 0
   Exit Sub

Genre_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Genre_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub


Private Sub movieDispPanel_GotFocus()
   On Error GoTo movieDispPanel_GotFocus_Error

If movieHolder.Enabled = True Then movieHolder.SetFocus
HideCaret movieDispPanel.hWnd

   On Error GoTo 0
   Exit Sub

movieDispPanel_GotFocus_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure movieDispPanel_GotFocus of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub

Private Sub movieDispPanel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   On Error GoTo movieDispPanel_MouseMove_Error

HideCaret movieDispPanel.hWnd

   On Error GoTo 0
   Exit Sub

movieDispPanel_MouseMove_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure movieDispPanel_MouseMove of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub movieHolder_DblClick()
   On Error GoTo movieHolder_DblClick_Error

On Error Resume Next
If movieHolder.ListItems.Count > 0 Then
Dim ex As New FileSystemObject
If ex.FileExists(movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieFileName) Then Shell "explorer " & Chr(34) & movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieFileName & Chr(34), vbNormalFocus
End If

   On Error GoTo 0
   Exit Sub

movieHolder_DblClick_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure movieHolder_DblClick of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub movieHolder_ItemClick(ByVal Item As ComctlLib.ListItem)

   On Error GoTo movieHolder_ItemClick_Error

displayMovieDetails Val(Item.SubItems(5))
'MsgBox movieArray(Val(Item.SubItems(5))).isMovie

   On Error GoTo 0
   Exit Sub

movieHolder_ItemClick_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure movieHolder_ItemClick of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub

Private Sub movieHolder_KeyPress(KeyAscii As Integer)
   On Error GoTo movieHolder_KeyPress_Error

On Error Resume Next
If KeyAscii = 13 Then

If movieHolder.ListItems.Count > 0 Then
Dim ex As New FileSystemObject
If ex.FileExists(movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieFileName) Then Shell "explorer " & Chr(34) & movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieFileName & Chr(34), vbNormalFocus
End If

End If

   On Error GoTo 0
   Exit Sub

movieHolder_KeyPress_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure movieHolder_KeyPress of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub



Private Sub movieHolder_KeyUp(KeyCode As Integer, Shift As Integer)
   On Error GoTo movieHolder_KeyUp_Error

If KeyCode = vbKeyDelete Then
topBar_ButtonClick topBar.buttons(8)
End If

   On Error GoTo 0
   Exit Sub

movieHolder_KeyUp_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure movieHolder_KeyUp of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub movieHolder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error GoTo movieHolder_MouseUp_Error

If Button = 2 Then
PopupMenu sortMenu, , x + 100, y + 1000
End If

   On Error GoTo 0
   Exit Sub

movieHolder_MouseUp_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure movieHolder_MouseUp of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub nWatched_Click()
   On Error GoTo nWatched_Click_Error

If nWatched.Value = True Then
Dim i

For i = LBound(movieArray) To UBound(movieArray)
If movieArray(i).MovieWatched = "No" Then
movieArray(i).wacthedCategory = True
Else
movieArray(i).wacthedCategory = False
End If
Next

End If
displayMovies

   On Error GoTo 0
   Exit Sub

nWatched_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure nWatched_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub ratinG_Change()
   On Error GoTo ratinG_Change_Error

rateVal = ""
If movieHolder.ListItems.Count > 0 Then
rateVal.Caption = "( " + str(ratinG.Value) + " / 10)"
   If Not movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieMyRating = ratinG.Value Then
            
          
          Dim mn As New ChilkatXml
          Dim mnC As New ChilkatXml
          Dim i, j
          mn.LoadXmlFile movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"
                              
          For i = 0 To mn.NumChildrenHavingTag("Movie") - 1
            Set mnC = mn.GetNthChildWithTag("Movie", i)
            If mnC.GetChildContent("UnID") = movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieUnID Then
                mnC.UpdateChildContent "MyRating", ratinG.Value
                mn.SaveXml movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"
                movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieMyRating = ratinG.Value
                movieHolder.SelectedItem.SubItems(3) = ratinG.Value
                movieHolder_ItemClick movieHolder.SelectedItem
                If movieHolder.SortKey = 3 Then movieHolder.Sorted = True
                movieHolder.SetFocus
                Exit Sub
            End If
          Next
   End If
End If

   On Error GoTo 0
   Exit Sub

ratinG_Change_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure ratinG_Change of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub ratinG_Scroll()
   On Error GoTo ratinG_Scroll_Error

rateVal.Caption = "( " + str(ratinG.Value) + " / 10)"

   On Error GoTo 0
   Exit Sub

ratinG_Scroll_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure ratinG_Scroll of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Public Sub scn_Click()
   On Error GoTo scn_Click_Error

movieDispPanel.Enabled = False
ratinG.Enabled = False

mnu.Enabled = False
sortMenu = False
hlp = False
topBar.Enabled = False
movieHolder.Enabled = False


statusS = "Scanning Folders..."
Call updateDataBase(True, True, statusS)
totalMovies = totalMovies + Val(statusS.Caption)
statusS = "Loading..."
Call loadAllMovies(statusS)
Genre_Populate
Call displayMovies

mnu.Enabled = True
sortMenu = True
hlp = True
topBar.Enabled = True
movieHolder.Enabled = True
movieDispPanel.Enabled = True
ratinG.Enabled = True

   On Error GoTo 0
   Exit Sub

scn_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure scn_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub searchBox_Change()
   On Error GoTo searchBox_Change_Error

If searchBox.Text = "" Then
Dim i
For i = LBound(movieArray) To UBound(movieArray)
        movieArray(i).MovieSearchRelevance = -1
    Next
movieHolder.SortKey = 6
displayMovies
End If

   On Error GoTo 0
   Exit Sub

searchBox_Change_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure searchBox_Change of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub searchBox_GotFocus()
   On Error GoTo searchBox_GotFocus_Error

Picture1.Visible = True
searchBox.SelStart = 0
searchBox.SelLength = Len(searchBox.Text)
searchBox.Alignment = vbLeftJustify
searchBox.FontItalic = False
searchBox.ForeColor = vbBlack

   On Error GoTo 0
   Exit Sub

searchBox_GotFocus_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure searchBox_GotFocus of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub searchBox_KeyPress(KeyAscii As Integer)
   On Error GoTo searchBox_KeyPress_Error

If KeyAscii = 13 And Not searchBox.Text = "" Then
searchMovies Trim(searchBox.Text), searchBy.Text
movieHolder.SortKey = 4
displayMovies
End If

   On Error GoTo 0
   Exit Sub

searchBox_KeyPress_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure searchBox_KeyPress of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub searchBox_LostFocus()
   On Error GoTo searchBox_LostFocus_Error

Picture1.Visible = False
If Trim(searchBox.Text) = "" Then searchBox.Text = "Search Movies"
If Trim(searchBox.Text) = "Search Movies" Then
searchBox.Alignment = vbCenter

searchBox.FontItalic = True
searchBox.ForeColor = RGB(100, 100, 100)
End If

   On Error GoTo 0
   Exit Sub

searchBox_LostFocus_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure searchBox_LostFocus of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Function Genre_Populate()
   On Error GoTo Genre_Populate_Error

genreIsReady = False
Genre.Clear
Dim searchI As Long
Dim genStr, genFlag
Dim genreNum
Dim SearchIiI
Dim genreItems() As String
ReDim genreItems(0) As String
Genre.AddItem "All", 0
Genre.ListIndex = 0

For searchI = LBound(movieArray) To UBound(movieArray) - 1
    
    For Each genStr In movieArray(searchI).MovieGenre
    
                    genFlag = 0
                    For genreNum = LBound(genreItems) To UBound(genreItems) - 1
                        If genreItems(genreNum) = genStr Then
                            genFlag = 1
                        End If
                    Next
                    
                    If genFlag = 0 And (Not genStr = "") Then
                        genreItems(UBound(genreItems)) = genStr
                        ReDim Preserve genreItems(UBound(genreItems) + 1) As String
                    End If
    
    Next


Next

StrSort genreItems, True, True
For Each SearchIiI In genreItems
If Not Trim(SearchIiI) = "" Then Genre.AddItem SearchIiI
Next
Genre.AddItem "Favourites"
Genre.AddItem "No Category"


genreIsReady = True

   On Error GoTo 0
   Exit Function

Genre_Populate_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Genre_Populate of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Function



Function init(statusObj As Control)
   On Error GoTo init_Error

loadAllMovies statusObj
Genre_Populate
displayMovies

   On Error GoTo 0
   Exit Function

init_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure init of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function

Private Sub searchBy_Click()
   On Error GoTo searchBy_Click_Error

On Error Resume Next
If Not (searchBox.Text = "" Or searchBox.Text = "Search Movies") Then
searchMovies Trim(searchBox.Text), searchBy.Text
movieHolder.SortKey = 4
displayMovies

End If
If movieHolder.SortOrder = lvwDescending Then
topBar.buttons(2).Image = 1
Else
topBar.buttons(2).Image = 2
End If

movieHolder.SetFocus

   On Error GoTo 0
   Exit Sub

searchBy_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure searchBy_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub srtAsc_Click()


   On Error GoTo srtAsc_Click_Error

srtAsc.Checked = True
srtDec.Checked = False
movieHolder.SortOrder = lvwAscending
movieHolder.Sorted = True
theButtonSort

   On Error GoTo 0
   Exit Sub

srtAsc_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure srtAsc_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub srtDateAdded_Click()


   On Error GoTo srtImdb_Click_Error

srtTitle.Checked = False
srtYear.Checked = False
srtImdb.Checked = False
srtMy.Checked = False
srtDateAdded.Checked = True
movieHolder.SortKey = 6
movieHolder.Sorted = True
movieHolder.SortOrder = lvwDescending
theButtonSort

   On Error GoTo 0
   Exit Sub

srtImdb_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure srtImdb_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub

Private Sub srtDec_Click()

   On Error GoTo srtDec_Click_Error

srtAsc.Checked = False
srtDec.Checked = True
movieHolder.SortOrder = lvwDescending
movieHolder.Sorted = True
theButtonSort

   On Error GoTo 0
   Exit Sub

srtDec_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure srtDec_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub srtImdb_Click()

   On Error GoTo srtImdb_Click_Error

srtTitle.Checked = False
srtYear.Checked = False
srtImdb.Checked = True
srtMy.Checked = False
srtDateAdded.Checked = False
movieHolder.SortKey = 2
movieHolder.Sorted = True
movieHolder.SortOrder = lvwDescending
theButtonSort

   On Error GoTo 0
   Exit Sub

srtImdb_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure srtImdb_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub srtMy_Click()

   On Error GoTo srtMy_Click_Error

srtTitle.Checked = False
srtYear.Checked = False
srtImdb.Checked = False
srtMy.Checked = True
srtDateAdded.Checked = False
movieHolder.SortKey = 3
movieHolder.Sorted = True
movieHolder.SortOrder = lvwDescending
theButtonSort

   On Error GoTo 0
   Exit Sub

srtMy_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure srtMy_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub srtTitle_Click()
   On Error GoTo srtTitle_Click_Error

srtTitle.Checked = True
srtYear.Checked = False
srtImdb.Checked = False
srtMy.Checked = False
srtDateAdded.Checked = False
movieHolder.SortKey = 0
movieHolder.Sorted = True
movieHolder.SortOrder = lvwAscending
theButtonSort

   On Error GoTo 0
   Exit Sub

srtTitle_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure srtTitle_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub srtYear_Click()
   On Error GoTo srtYear_Click_Error

srtTitle.Checked = False
srtYear.Checked = True
srtImdb.Checked = False
srtMy.Checked = False
srtDateAdded.Checked = False
movieHolder.SortKey = 1
movieHolder.Sorted = True
movieHolder.SortOrder = lvwAscending
theButtonSort

   On Error GoTo 0
   Exit Sub

srtYear_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure srtYear_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub statusS_Change()
   On Error GoTo statusS_Change_Error

movieDispPanel.TextRTF = ""
movieDispPanel.SelFontName = "Arial"
movieDispPanel.SelColor = RGB(0, 100, 0)
movieDispPanel.SelAlignment = rtfLeft
movieDispPanel.SelFontSize = 12
movieDispPanel.SelText = statusS.Caption

   On Error GoTo 0
   Exit Sub

statusS_Change_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure statusS_Change of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub


Private Sub theButtonSort()
   On Error GoTo theButtonSort_Error

   If movieHolder.ListItems.Count > 0 Then
        movieHolder.SelectedItem = movieHolder.ListItems(1)
        movieHolder.SelectedItem.EnsureVisible
        movieHolder_ItemClick movieHolder.ListItems(1)
    End If
If movieHolder.SortOrder = lvwAscending Then
        topBar.buttons(2).Image = 2
    Else
        topBar.buttons(2).Image = 1
End If


   On Error GoTo 0
   Exit Sub

theButtonSort_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure theButtonSort of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub

Private Sub topBar_ButtonClick(ByVal Button As ComctlLib.Button)

Dim i
   On Error GoTo topBar_ButtonClick_Error

Select Case Button.Key

Case "sorteR"

    If movieHolder.SortOrder = lvwAscending Then
        Button.Image = 1 'sorteR.Picture = buttons.ListImages(1).Picture
        movieHolder.SortOrder = lvwDescending
        srtAsc.Checked = False
        srtDec.Checked = True
    Else
        Button.Image = 2 'sorteR.Picture = buttons.ListImages(2).Picture
        movieHolder.SortOrder = lvwAscending
        srtAsc.Checked = True
        srtDec.Checked = False
    End If

    If movieHolder.ListItems.Count > 0 Then
        movieHolder.SelectedItem = movieHolder.ListItems(1)
        movieHolder.SelectedItem.EnsureVisible
        movieHolder_ItemClick movieHolder.ListItems(1)
    End If
    movieHolder.Sorted = True
    movieHolder.SetFocus
Case "openMovie"
If movieHolder.ListItems.Count > 0 Then
Dim ex As New FileSystemObject
If ex.FileExists(movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieFileName) Then Shell "explorer " & Chr(34) & movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieFileName & Chr(34), vbNormalFocus
End If

Case "openFolder"
If movieHolder.ListItems.Count > 0 Then


i = InStrRev(movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieFileName, "\")
Dim exF As New FileSystemObject
If exF.FolderExists(movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & Left(movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieFileName, i)) Then Shell "explorer " & Chr(34) & movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & Left(movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieFileName, i) & Chr(34), vbNormalFocus
End If

Case "isNoMovie"
    movieHolder.SetFocus
    If Not movieHolder.ListItems.Count > 0 Then Exit Sub
        Dim ms As New ChilkatXml
        
        Dim sure
        sure = MsgBox("Are you Sure To Remove the Movie " + Chr(34) + movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieTitle + Chr(34) + " ?", vbQuestion + vbYesNo)
        If sure = vbYes Then
            ms.LoadXmlFile movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"
            For i = 0 To ms.NumChildrenHavingTag("Movie") - 1
                Set ms = ms.GetNthChildWithTag("Movie", i)
                If ms.GetChildContent("UnID") = movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieUnID Then
                    ms.UpdateAttribute "isMovie", "No"
                    Set ms = ms.getParent
                    ms.SaveXml movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"
                    movieArray(Val(movieHolder.SelectedItem.SubItems(5))).isMovie = "No"
                    movieHolder.ListItems.Remove movieHolder.SelectedItem.Index
                    movieHolder.SelectedItem.EnsureVisible
                    movieHolder_ItemClick movieHolder.SelectedItem
                    
                    Exit Sub
    
                End If
                Set ms = ms.getParent
            Next
        End If
        
Case "movieWatched"
If movieHolder.ListItems.Count > 0 Then

Dim myX As New ChilkatXml
Dim myXc As New ChilkatXml
Dim strSta As String

If Button.Value = tbrPressed Then
strSta = "Yes"
Else
strSta = "No"
End If


myX.LoadXmlFile movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"

movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieWatched = strSta
For i = 0 To myX.NumChildrenHavingTag("Movie") - 1
Set myXc = myX.GetNthChildWithTag("Movie", i)
If myXc.GetChildContent("UnID") = movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieUnID Then
myXc.UpdateChildContent "Watched", strSta
myX.SaveXml movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"
Exit Sub
End If

Next

myX.SaveXml movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"
End If

Case "editMovie"
    If movieHolder.ListItems.Count > 0 Then
    editMovie.mIndex = Val(movieHolder.SelectedItem.SubItems(5))
    editMovie.Show 1
    End If
Case "goToImdb"
If movieHolder.ListItems.Count > 0 Then
If Not movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieIMDBid = "" Then
Dim retvalue
retvalue = ShellExecute(mainF.hWnd, "Open", "http://www.imdb.com/title/tt" + movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieIMDBid, 0&, 0&, 0&)
Else
MsgBox "Sorry. This movie file doesn't have an IMDb ID. ", vbInformation, "My Movie Manager"
End If
End If

Case "addFav"
If movieHolder.ListItems.Count > 0 Then
    If Button.Tag = "Yes" Then
        movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieIsFav = "No"
        Button.Tag = "No"
        Button.Image = 8
    Else
        movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieIsFav = "Yes"
        Button.Tag = "Yes"
        Button.Image = 6
    End If
    Dim msS As New ChilkatXml
    msS.LoadXmlFile movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"
            For i = 0 To msS.NumChildrenHavingTag("Movie") - 1
                Set msS = msS.GetNthChildWithTag("Movie", i)
                If msS.GetChildContent("UnID") = movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieUnID Then
                    msS.UpdateChildContent "Fav", movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieIsFav
                    Set msS = msS.getParent
                    msS.SaveXml movieArray(Val(movieHolder.SelectedItem.SubItems(5))).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"
                    If Genre.Text = "Favourites" Then Genre_Click
                    Exit Sub
    
                End If
                Set msS = msS.getParent
            Next
    

End If

Case "sortBy"
    Button.Value = tbrPressed
    PopupMenu srtBy, , Button.Left, Button.Height
    Button.Value = tbrUnpressed


End Select

   On Error GoTo 0
   Exit Sub

topBar_ButtonClick_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure topBar_ButtonClick of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

    
End Sub

Function insertRTBimg(Rtb As RichTextBox, img As StdPicture)
   On Error GoTo insertRTBimg_Error

On Error Resume Next
Dim std As New StdPicture
Dim strR As String
    strR = Clipboard.GetText
    Set std = Clipboard.GetData()
     Clipboard.Clear
'Populate the clipboard with image 1's picture
    Clipboard.SetData img
'Add it to rtb1
    SendMessage Rtb.hWnd, WM_PASTE, 0, 0
    Clipboard.Clear
    
    Clipboard.SetData std
    
    Clipboard.SetText strR


   On Error GoTo 0
   Exit Function

insertRTBimg_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure insertRTBimg of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function

Function displayMovieDetails(clickedMovie As Long)



   On Error GoTo displayMovieDetails_Error

Me.Caption = "My Movie Manager - " & movieArray(clickedMovie).MovieLocFolder & movieArray(clickedMovie).MovieFileName

movieDispPanel.TextRTF = ""
movieDispPanel.Font = "Arial"
movieDispPanel.SelAlignment = rtfCenter
movieDispPanel.SelFontSize = 20
movieDispPanel.SelBold = True
movieDispPanel.SelColor = vbBlack
movieDispPanel.SelText = movieArray(clickedMovie).MovieTitle
movieDispPanel.SelText = vbCrLf


Dim checkCover As New FileSystemObject
If checkCover.FileExists(movieArray(clickedMovie).MovieLocFolder & "\MyMovieManager_Data_XYZ\Covers\" & movieArray(clickedMovie).MovieCoverLarge) Then
insertRTBimg movieDispPanel, LoadPicture(movieArray(clickedMovie).MovieLocFolder & "\MyMovieManager_Data_XYZ\Covers\" & movieArray(clickedMovie).MovieCoverLarge)
Else
insertRTBimg movieDispPanel, LoadPicture(App.path & "\Images\defaultCover.jpg")
End If


movieDispPanel.SelFontSize = 12
If Not movieArray(clickedMovie).MovieYear = "" Then movieDispPanel.SelText = vbCrLf & movieArray(clickedMovie).MovieYear
If Not movieArray(clickedMovie).MovieIMDbRating = "" Then movieDispPanel.SelText = vbCrLf & "IMDb Rating: " & movieArray(clickedMovie).MovieIMDbRating & " / 10"
movieDispPanel.SelText = vbCrLf & "Your Rating: " & movieArray(clickedMovie).MovieMyRating & " / 10"
movieDispPanel.SelFontSize = 11
If Not movieArray(clickedMovie).MovieDirector = "" Then movieDispPanel.SelText = vbCrLf & "Directed By: " & movieArray(clickedMovie).MovieDirector
movieDispPanel.SelFontSize = 10
movieDispPanel.SelBold = False


If Not movieArray(clickedMovie).MovieDuration = "" Then movieDispPanel.SelText = vbCrLf & "Duration: " & movieArray(clickedMovie).MovieDuration

If Not movieArray(clickedMovie).MovieLanguage = "" Then movieDispPanel.SelText = vbCrLf & "Languages: " & movieArray(clickedMovie).MovieLanguage
If Not movieArray(clickedMovie).MovieCountry = "" Then movieDispPanel.SelText = vbCrLf & "Countries: " & movieArray(clickedMovie).MovieCountry

If Not UBound(movieArray(clickedMovie).MovieGenre) = 0 Then movieDispPanel.SelText = vbCrLf & "Genres: "
Dim genreTypes

For Each genreTypes In movieArray(clickedMovie).MovieGenre
If Not genreTypes = "" Then movieDispPanel.SelText = genreTypes & ", "
Next
If Not UBound(movieArray(clickedMovie).MovieGenre) = 0 Then

movieDispPanel.SelStart = movieDispPanel.SelStart - 2
End If

If Not movieArray(clickedMovie).MoviePlot = "" Then

movieDispPanel.SelText = vbCrLf
movieDispPanel.SelFontSize = 12
movieDispPanel.SelBold = True
movieDispPanel.SelText = "Description: "
movieDispPanel.SelAlignment = rtfLeft
movieDispPanel.SelFontSize = 9
movieDispPanel.SelBold = False

movieDispPanel.SelText = movieArray(clickedMovie).MoviePlot
End If

If Not movieArray(clickedMovie).MovieCast = "" Then

movieDispPanel.SelText = vbCrLf
movieDispPanel.SelFontSize = 12
movieDispPanel.SelBold = True
movieDispPanel.SelText = "Cast: "

movieDispPanel.SelFontSize = 9
movieDispPanel.SelBold = False
movieDispPanel.SelText = movieArray(clickedMovie).MovieCast
End If
If movieArray(clickedMovie).MovieWatched = "Yes" Then
topBar.buttons(9).Value = tbrPressed
Else
topBar.buttons(9).Value = tbrUnpressed
End If

If movieArray(clickedMovie).MovieCast = "" And movieArray(clickedMovie).MovieCountry = "" And movieArray(clickedMovie).MovieDirector = "" And movieArray(clickedMovie).MovieDuration = "" And UBound(movieArray(clickedMovie).MovieGenre) = 0 And movieArray(clickedMovie).MovieIMDbRating = "" And movieArray(clickedMovie).MovieLanguage = "" And movieArray(clickedMovie).MoviePlot = "" And movieArray(clickedMovie).MovieYear = "" Then
movieDispPanel.SelLength = Len(movieDispPanel.Text)
movieDispPanel.SelText = vbCrLf & vbCrLf & "No Other Details Found."
End If
movieDispPanel.SelStart = 0
movieDispPanel.SelLength = 0

If movieArray(clickedMovie).MovieIsFav = "Yes" Then
topBar.buttons(10).Tag = "Yes"
topBar.buttons(10).Image = 6
Else
topBar.buttons(10).Tag = "No"
topBar.buttons(10).Image = 8
End If
ratinG.Value = Val(movieArray(clickedMovie).MovieMyRating)

   On Error GoTo 0
   Exit Function

displayMovieDetails_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure displayMovieDetails of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function


Private Sub upImdb_Click()


   On Error GoTo upImdb_Click_Error

stopSignalNet = False
updateFromNetForm.Show 1

movieDispPanel.Enabled = False
ratinG.Enabled = False

mnu.Enabled = False
sortMenu = False
hlp = False
topBar.Enabled = False
movieHolder.Enabled = False
statusS = "Loading Movies..."
Call loadAllMovies(statusS)
Genre_Populate
displayMovies


mnu.Enabled = True
sortMenu = True
hlp = True
topBar.Enabled = True
movieHolder.Enabled = True
ratinG.Enabled = True
movieDispPanel.Enabled = True

   On Error GoTo 0
   Exit Sub

upImdb_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure upImdb_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub Watched_Click()
   On Error GoTo Watched_Click_Error

If Watched.Value = True Then
Dim i

For i = LBound(movieArray) To UBound(movieArray)
If movieArray(i).MovieWatched = "Yes" Then
movieArray(i).wacthedCategory = True
Else
movieArray(i).wacthedCategory = False
End If
Next

End If

displayMovies

   On Error GoTo 0
   Exit Sub

Watched_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Watched_Click of Form mainF" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub wb_Click()
ShellExecute Me.hWnd, "Open", "http://experimenter.x10.mx", 0&, 0&, 0&
End Sub
