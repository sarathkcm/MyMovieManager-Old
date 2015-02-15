VERSION 5.00
Begin VB.Form editMovie 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Movie Data - My Movie Manager"
   ClientHeight    =   6135
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7320
   Icon            =   "editMovie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton OKButton 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox imdbID 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton fetchDetails 
      Caption         =   "Fetch Movie Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label statusF 
      Caption         =   "Label3"
      Height          =   135
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"editMovie.frx":123F2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMDB ID :"
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
      Left            =   360
      TabIndex        =   8
      Top             =   1875
      Width           =   840
   End
   Begin VB.Image cover 
      Height          =   3015
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label nameS 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2520
      Width           =   4215
   End
   Begin VB.Label year 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label direc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   3480
      Width           =   4215
   End
   Begin VB.Label plot 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      TabIndex        =   4
      Top             =   4320
      Width           =   4215
   End
End
Attribute VB_Name = "editMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Declare Function URLDownloadToFile Lib "urlmon" Alias _
"URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
ByVal szFileName As String, ByVal dwReserved As Long, _
ByVal lpfnCB As Long) As Long

Public mIndex As Long
Dim movieData() As String
Dim statusMsg As String

Private Sub CancelButton_Click()
   On Error GoTo CancelButton_Click_Error

Unload Me

   On Error GoTo 0
   Exit Sub

CancelButton_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure CancelButton_Click of Form editMovie" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub fetchDetails_Click()
Dim loginData() As String



   On Error GoTo fetchDetails_Click_Error

loginData = xmlServerActions.loginXml()

If loginData(0) = "200 OK" Then

    movieData = xmlServerActions.getIMDBdetailsXml(loginData(3), loginData(1), imdbID.Text)

    
If movieData(0) = "200 OK" Then
    cover.Picture = xmlServerActions.LoadPicture(movieData(12))
    nameS = movieData(1)
    direc = "Directed By: " + movieData(7)
    year = movieData(2)
    plot = movieData(9)

statusF = "OK"
End If

Else
MsgBox "There was an Error in Connection. We are Sorry.", vbInformation
End If
xmlServerActions.logoutXml loginData(3), loginData(1)

   On Error GoTo 0
   Exit Sub

fetchDetails_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure fetchDetails_Click of Form editMovie" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub




Private Sub OKButton_Click()
   On Error GoTo OKButton_Click_Error

If statusF = "OK" Then

Dim ms As New ChilkatXml
Dim i


ms.LoadXmlFile movieArray(mIndex).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"
For i = 0 To ms.NumChildrenHavingTag("Movie") - 1
Set ms = ms.GetNthChildWithTag("Movie", i)
If ms.GetChildContent("UnID") = movieArray(mIndex).MovieUnID Then
    

    ms.UpdateChildContent "ImdbID", movieData(8)
    movieArray(mIndex).MovieIMDBid = movieData(8)
    
    ms.UpdateChildContent "Title", movieData(1)
    If Not movieData(1) = "" Then movieArray(mIndex).MovieTitle = movieData(1)
    
    ms.UpdateChildContent "Year", movieData(2)
    movieArray(mIndex).MovieYear = movieData(2)

    ms.UpdateChildContent "IMDBRating", movieData(3)
    movieArray(mIndex).MovieIMDbRating = movieData(3)
 
    
    ms.UpdateChildContent "Languages", movieData(4)
    movieArray(mIndex).MovieLanguage = movieData(4)
  
    
    ms.UpdateChildContent "Country", movieData(5)
    movieArray(mIndex).MovieCountry = movieData(5)
    
   
    
    ms.UpdateChildContent "Duration", movieData(6)
    movieArray(mIndex).MovieDuration = movieData(6)
    
    
    ms.UpdateChildContent "Directors", movieData(7)
    movieArray(mIndex).MovieDirector = movieData(7)
    
    
    ms.UpdateChildContent "Plot", movieData(9)
    movieArray(mIndex).MoviePlot = movieData(9)
    
    
    
    ms.removeChild ("Genres")
    Dim myGenre As New ChilkatXml
    myGenre.loadXML "<Genres>" & movieData(10) & "</Genres>"
    ms.AddChildTree myGenre
    
    Dim igen
    Dim genStr
    ReDim movieArray(mIndex).MovieGenre(0 To myGenre.NumChildrenHavingTag("value") - 1) As String
    
    For igen = 0 To myGenre.NumChildrenHavingTag("value") - 1
                     
            Set myGenre = myGenre.GetNthChildWithTag("value", igen)
            genStr = myGenre.Content
            movieArray(mIndex).MovieGenre(igen) = genStr
            Set myGenre = myGenre.getParent
    Next
                    
    
    
    ms.UpdateChildContent "Cast", movieData(13)
    movieArray(mIndex).MovieCast = movieData(13)
    
    ms.UpdateChildContent "UpdatedOnce", "Yes"
    movieArray(mIndex).MovieUpdatedOnce = "Yes"
    
    ms.UpdateChildContent "HashFailed", "Yes"
    movieArray(mIndex).MovieHashFailed = "Yes"
    
    
    If movieData(11) = "" Then
    
        ms.UpdateChildContent "CoverSmall", ""
    
        ms.UpdateChildContent "CoverLarge", ""
    
    
    Else
    
        Dim errcode As Long
        Dim urlL As String
        Dim localFileName As String
        urlL = movieData(11)
        localFileName = movieArray(mIndex).MovieLocFolder & "\MyMovieManager_Data_XYZ\Covers\" & ms.GetChildContent("UnID") & ".jpg"
        
        errcode = URLDownloadToFile(0, urlL, localFileName, 0, 0)
    
        If errcode = 0 Then
    
            ms.UpdateChildContent "CoverSmall", ms.GetChildContent("UnID") & ".jpg"
            movieArray(mIndex).MovieCoverSmall = ms.GetChildContent("UnID") & ".jpg"
            Set movieArray(mIndex).MovieIcon = LoadPicture(movieArray(mIndex).MovieLocFolder & "\MyMovieManager_Data_XYZ\Covers\" & movieArray(mIndex).MovieCoverSmall)
        Else
            statusMsg = "Error while downloading Cover Photo."
    
        End If
    
    
    
        urlL = movieData(12)
        localFileName = movieArray(mIndex).MovieLocFolder & "\MyMovieManager_Data_XYZ\Covers\" & ms.GetChildContent("UnID") & "_large.jpg"
        errcode = URLDownloadToFile(0, urlL, localFileName, 0, 0)
    
    
    
        If errcode = 0 Then
    
            ms.UpdateChildContent "CoverLarge", ms.GetChildContent("UnID") & "_large.jpg"
            movieArray(mIndex).MovieCoverLarge = ms.GetChildContent("UnID") & "_large.jpg"
        Else
            statusMsg = "Error while downloading Cover Photo."
    
        End If
    
    End If


Set ms = ms.getParent
ms.SaveXml movieArray(mIndex).MovieLocFolder & "\MyMovieManager_Data_XYZ\Data.xml"

mainF.displayMovies
Unload Me
Exit Sub

End If

Set ms = ms.getParent

Next





End If


Unload Me


   On Error GoTo 0
   Exit Sub

OKButton_Click_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure OKButton_Click of Form editMovie" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub


