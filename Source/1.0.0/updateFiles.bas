Attribute VB_Name = "updateFiles"
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Dim progressM As Long
Dim FileList() As String
Public stopSignalNet As Boolean
Public moviesNotUpdated As Long
Public totalMovies As Long
Dim imdbArray() As String


Private Type MovieInfo
        MovieFileName As String
        MovieWatched As String
        MovieHash As String
        MovieIMDBid As String
        MovieTitle As String
        MovieYear As String
        MovieIMDbRating As String
        MovieMyRating As String
        MovieLanguage As String
        MovieCountry As String
        MovieDuration As String
        MovieDirector As String
        MoviePlot As String
        MovieGenre() As String
        MovieCoverSmall As String
        MovieCoverLarge As String
        MovieCast As String
        MovieIsFav As String
        MovieHashFailed As String
        MovieUpdatedOnce As String
        MovieLocFolder As String
        MovieUnID As String
        MovieIcon As StdPicture
        MovieSearchRelevance As Integer
        movieIdentifier As Long
        MovieDisplayFlag As Boolean
        isMovie As String
        wacthedCategory As Boolean
        DateAdded As Date
        
End Type

Public movieArray() As MovieInfo


Public Function updateFileList(statusObj As Control)

Dim mObj As New ChilkatXml
Dim nObj As New ChilkatXml
Dim i, j


statusObj = "Scanning Started... Please be patient..."
   On Error GoTo updateFileList_Error

mObj.LoadXmlFile App.path & "\Data\Loc.xml"




For i = 0 To mObj.NumChildrenHavingTag("Location") - 1
    
    DoEvents
    ReDim FileList(0) As String

    Set nObj = mObj.GetNthChildWithTag("Location", i)
    statusObj = "Scanning " & nObj.Content
    
    If nObj.GetAttrValue("rec") = "Yes" Then
        recFolder nObj.Content, statusObj

    Else
        Dim noRec As New FileSystemObject
        Dim noRecF As Folder
        Set noRecF = noRec.GetFolder(nObj.Content)
        For Each j In noRecF.Files
           If LCase(Right(j, 4)) = ".avi" Or LCase(Right(j, 4)) = ".mkv" Or LCase(Right(j, 4)) = ".dat" Or LCase(Right(j, 4)) = ".vob" Or LCase(Right(j, 4)) = ".mpg" Or LCase(Right(j, 5)) = ".mpeg" Or LCase(Right(j, 4)) = ".wmv" Or LCase(Right(j, 4)) = ".mp4" Or LCase(Right(j, 4)) = ".vob" Or LCase(Right(j, 4)) = ".mov" Or LCase(Right(j, 4)) = ".flv" Then
                If j.Size > 2000000 Then
                    FileList(UBound(FileList)) = j
                    statusObj = str(UBound(FileList)) + " Files Found"
                    ReDim Preserve FileList(UBound(FileList) + 1) As String
                End If
           End If
            DoEvents
        Next
    End If
    
    createNewDatabase nObj.Content, FileList, statusObj
DoEvents
Next

   On Error GoTo 0
   Exit Function

updateFileList_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure updateFileList of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function






Public Function recFolder(path As String, statusObj As Control)

Dim fCollection As New FileSystemObject
Dim fl As Folder

Dim i, j
   On Error GoTo recFolder_Error

Set fl = fCollection.GetFolder(path)


For Each j In fl.Files
    If LCase(Right(j, 4)) = ".avi" Or LCase(Right(j, 4)) = ".mkv" Or LCase(Right(j, 4)) = ".dat" Or LCase(Right(j, 4)) = ".vob" Or LCase(Right(j, 4)) = ".mpg" Or LCase(Right(j, 5)) = ".mpeg" Or LCase(Right(j, 4)) = ".wmv" Or LCase(Right(j, 4)) = ".mp4" Or LCase(Right(j, 4)) = ".vob" Or LCase(Right(j, 4)) = ".mov" Or LCase(Right(j, 4)) = ".flv" Then
        If j.Size > 2000000 Then
            statusObj = str(UBound(FileList)) + " Files Found"
            FileList(UBound(FileList)) = j
            ReDim Preserve FileList(UBound(FileList) + 1) As String
        
            DoEvents
        End If
    End If
Next




For Each i In fl.SubFolders
    DoEvents
    recFolder i.path, statusObj
Next


   On Error GoTo 0
   Exit Function

recFolder_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure recFolder of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function


Public Function genUnID()
Dim i, k
Dim str As String
Dim id As String
   On Error GoTo genUnID_Error

str = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
id = ""
For i = 0 To 20
k = Int(Rnd * 10000) Mod 61 + 1
id = id & Mid(str, k, 1)
Next
genUnID = id

   On Error GoTo 0
   Exit Function

genUnID_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure genUnID of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Function


Public Function createNewDatabase(dir As String, vidList() As String, statusObj As Control)
'the ubound of vidlist has no value

Dim db As New ChilkatXml
Dim dbCh As ChilkatXml
Dim i, strF
Dim uID
Dim prgs As Long, maxPrgs As Long
   On Error GoTo createNewDatabase_Error

maxPrgs = UBound(vidList) - LBound(vidList)
Set db = db.NewChild("Movies", "")

For i = LBound(vidList) To UBound(vidList) - 1
    
    Set dbCh = db.NewChild("Movie", "")
    dbCh.addAttribute "isMovie", "Yes"


    strF = vidList(i)

    strF = Right(strF, Len(strF) - Len(dir))
    dbCh.NewChild "FileName", strF
    dbCh.NewChild "Watched", "No"
    dbCh.NewChild "Hash", find_Hash(vidList(i))
    dbCh.NewChild "ImdbID", ""
    
    
    Dim l
    l = InStrRev(strF, "\")
    dbCh.NewChild "Title", Right(strF, Len(strF) - l)
    
    dbCh.NewChild "Year", ""
    dbCh.NewChild "IMDBRating", ""
    dbCh.NewChild "MyRating", "0"
    dbCh.NewChild "Languages", ""
    dbCh.NewChild "Country", ""
    dbCh.NewChild "Duration", ""
    dbCh.NewChild "Directors", ""
    dbCh.NewChild "Plot", ""
    dbCh.NewChild "Genres", ""
    dbCh.NewChild "CoverSmall", ""
    dbCh.NewChild "CoverLarge", ""
    dbCh.NewChild "Cast", ""
    dbCh.NewChild "HashFailed", "No"
    dbCh.NewChild "UpdatedOnce", "No"
    dbCh.NewChild "Fav", "No"
    dbCh.NewChild "DateAdded", DateTime.Now
    Dim yesOrNo, abc, NoMovies
    Dim fDbCh As New ChilkatXml

    NoMovies = db.NumChildrenHavingTag("Movie")

Repeat:     uID = genUnID


    yesOrNo = 0
    For abc = 0 To NoMovies - 1

        Set fDbCh = db.GetNthChildWithTag("Movie", abc)

        If uID = fDbCh.GetChildContent("UnID") Then yesOrNo = 1

    Next


    If yesOrNo = 0 Then
        dbCh.NewChild "UnID", uID
    Else
        GoTo Repeat
    End If

    prgs = CInt(((i + 1) / maxPrgs) * 100)
    statusObj = "Creating Database at Location..." & str(prgs) & " %"
    DoEvents
Next

Dim sk As New FileSystemObject

If Not sk.FolderExists(dir & "\MyMovieManager_Data_XYZ") Then sk.CreateFolder (dir & "\MyMovieManager_Data_XYZ")

If sk.FolderExists(dir & "\MyMovieManager_Data_XYZ") Then
    
    Dim mmm As Folder
    Set mmm = sk.GetFolder(dir & "\MyMovieManager_Data_XYZ")
    mmm.Attributes = Hidden + Directory + System

End If

db.SaveXml dir & "\MyMovieManager_Data_XYZ\data.xml"

   On Error GoTo 0
   Exit Function

createNewDatabase_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure createNewDatabase of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function

Public Function find_Hash(path As String) As String
   On Error GoTo find_Hash_Error

On Error GoTo ErrM:
Dim fl As Long
Dim fsize As Variant
Dim byteArray(7) As Byte
Dim intArray(7) As Long
Dim Hash As String
Dim i, j
fl = FreeFile()

For j = 0 To 7
    intArray(j) = 0
Next


fsize = FileLen(path)
If fsize < 0 Then GoTo ErrM

Open path For Binary As #fl
Seek #fl, 1
For i = 0 To 8191
Get #fl, , byteArray
    For j = 0 To 7
    intArray(j) = intArray(j) + byteArray(j)
    Next
Next

If (fsize - 65536 + 1) > 0 Then
Seek #fl, fsize - 65536 + 1
Else
Seek #fl, 1
End If


For i = 0 To 8191
Get #fl, , byteArray
    For j = 0 To 7
    intArray(j) = intArray(j) + byteArray(j)
    Next
Next
Close #fl


For i = 0 To 7

intArray(i) = intArray(i) + fsize Mod 256
fsize = fsize \ 256

Next




For i = 0 To 6

intArray(i + 1) = intArray(i + 1) + (intArray(i) \ 256)
Next


For i = 0 To 7

byteArray(i) = CByte(intArray(i) Mod 256)

Next

Hash = ""
For i = 7 To 0 Step -1
If Len(Hex(byteArray(i))) < 2 Then

Hash = Hash & "0" & Hex(byteArray(i))
Else
Hash = Hash & Hex(byteArray(i))
End If
Next
Hash = LCase(Hash)
'MsgBox path + vbCrLf + Hash
find_Hash = Hash
Exit Function

ErrM:

find_Hash = 0

'8E245D9679D31E12

   On Error GoTo 0
   Exit Function

find_Hash_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure find_Hash of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Function


Public Function updateDataBase(delEntry As Boolean, addEntry As Boolean, statusObj As Control)

Dim deletedNo, addNo
Dim mObj As New ChilkatXml
Dim nObj As New ChilkatXml
Dim i, kj, j
Dim noRec As New FileSystemObject
Dim noRecF As Folder

   On Error GoTo updateDataBase_Error

If addEntry Then
    addNo = 0
    

    mObj.LoadXmlFile App.path & "\Data\Loc.xml"



    kj = mObj.NumChildrenHavingTag("Location")


    For i = 0 To kj - 1

        ReDim FileList(0) As String

        Set nObj = mObj.GetNthChildWithTag("Location", i)

        If nObj.GetAttrValue("rec") = "Yes" Then
            recFolder nObj.Content, statusObj

        Else


            Set noRecF = noRec.GetFolder(nObj.Content)
            Dim fileJ
            For Each fileJ In noRecF.Files
                
                If LCase(Right(fileJ, 4)) = ".avi" Or LCase(Right(fileJ, 4)) = ".mkv" Or LCase(Right(fileJ, 4)) = ".dat" Or LCase(Right(fileJ, 4)) = ".vob" Or LCase(Right(fileJ, 4)) = ".mpg" Or LCase(Right(fileJ, 5)) = ".mpeg" Or LCase(Right(fileJ, 4)) = ".wmv" Or LCase(Right(fileJ, 4)) = ".mp4" Or LCase(Right(fileJ, 4)) = ".vob" Or LCase(Right(fileJ, 4)) = ".mov" Or LCase(Right(fileJ, 4)) = ".flv" Then
                    If fileJ.Size > 2000000 Then
                        FileList(UBound(FileList)) = fileJ
                        ReDim Preserve FileList(UBound(FileList) + 1) As String
                    
                        DoEvents
                    End If
                End If
            Next
        End If


'''''''''''''''''''''''''''''''''''''''

        Dim db As New ChilkatXml
        Dim dbCh As New ChilkatXml
        Dim fDbCh As New ChilkatXml

        Dim uID
        Dim strF
        Dim NoMovies, abc, yesOrNo
       
        Dim sk As New FileSystemObject
        If Not sk.FolderExists(nObj.Content & "\MyMovieManager_Data_XYZ") Then sk.CreateFolder (nObj.Content & "\MyMovieManager_Data_XYZ")
        If sk.FolderExists(nObj.Content & "\MyMovieManager_Data_XYZ") Then
            Dim mmm As Folder

            Set mmm = sk.GetFolder(nObj.Content & "\MyMovieManager_Data_XYZ")
            mmm.Attributes = Hidden + Directory + System
        End If


''untested code
''''''''''''''''''''''''''''''
        If Not sk.FileExists(nObj.Content & "\MyMovieManager_Data_XYZ\data.xml") Then
            Set db = db.NewChild("Movies", "")
        Else
         db.LoadXmlFile nObj.Content & "\MyMovieManager_Data_XYZ\data.xml"
        End If
''''''''''''''''''''''''''''''


        NoMovies = db.NumChildrenHavingTag("Movie")


        For j = LBound(FileList) To UBound(FileList) - 1
           
            
            
            DoEvents

            strF = FileList(j)
            strF = Right(strF, Len(strF) - Len(nObj.Content))
            yesOrNo = 0
            
            For abc = 0 To NoMovies - 1
                DoEvents
                
                Set fDbCh = db.GetNthChildWithTag("Movie", abc)
                If strF = fDbCh.GetChildContent("FileName") Then yesOrNo = 1
                    
            Next
            
           
            
            If yesOrNo = 0 Then
                
                Dim yesOrNoHash
                
                yesOrNoHash = 0
                Dim abcY
                Dim mnA As String
                    mnA = find_Hash(FileList(j))
                For abcY = 0 To NoMovies - 1
                    DoEvents
                
                    Set fDbCh = db.GetNthChildWithTag("Movie", abcY)
                    
                    If mnA = fDbCh.GetChildContent("Hash") And (Not Trim(fDbCh.GetChildContent("Hash")) = "0") And (Not Trim(fDbCh.GetChildContent("Hash")) = "") Then
                        yesOrNoHash = 1
                        GoTo sM:
                    End If
                    
                Next
                
                
sM:
                addNo = addNo + 1
                statusObj = str(addNo) & " New Files Added..."

                Set dbCh = db.NewChild("Movie", "")
                dbCh.addAttribute "isMovie", "Yes"
                dbCh.NewChild "FileName", strF
                dbCh.NewChild "Watched", "No"
                dbCh.NewChild "Hash", find_Hash(FileList(j))
                dbCh.NewChild "ImdbID", ""
                dbCh.NewChild "Title", ""
                dbCh.NewChild "Year", ""
                dbCh.NewChild "IMDBRating", ""
                dbCh.NewChild "MyRating", "0"
                dbCh.NewChild "Languages", ""
                dbCh.NewChild "Country", ""
                dbCh.NewChild "Duration", ""
                dbCh.NewChild "Directors", ""
                dbCh.NewChild "Plot", ""
                dbCh.NewChild "Genres", ""
                dbCh.NewChild "CoverSmall", ""
                dbCh.NewChild "CoverLarge", ""
                dbCh.NewChild "Cast", ""
                dbCh.NewChild "HashFailed", "No"
                dbCh.NewChild "UpdatedOnce", "No"
                dbCh.NewChild "Fav", "No"
                dbCh.NewChild "DateAdded", DateTime.Now
                    
Repeat:         uID = genUnID

                yesOrNo = 0
                For abc = 0 To NoMovies - 1
                    DoEvents
                    Set fDbCh = db.GetNthChildWithTag("Movie", abc)

                    If uID = fDbCh.GetChildContent("UnID") Then yesOrNo = 1

                Next





                If yesOrNo = 0 Then
                    dbCh.NewChild "UnID", uID
                Else
                    GoTo Repeat
                End If
              
            If yesOrNoHash = 1 Then
                
                Set fDbCh = db.GetNthChildWithTag("Movie", abcY)
                
                
                
                
                dbCh.UpdateAttribute "isMovie", "Yes"
                dbCh.UpdateChildContent "ImdbID", fDbCh.GetChildContent("ImdbID")
                dbCh.UpdateChildContent "Title", fDbCh.GetChildContent("Title")
                dbCh.UpdateChildContent "Year", fDbCh.GetChildContent("Year")
                dbCh.UpdateChildContent "IMDBRating", fDbCh.GetChildContent("IMDBRating")
                dbCh.UpdateChildContent "MyRating", "0"
                dbCh.UpdateChildContent "Languages", fDbCh.GetChildContent("Languages")
                dbCh.UpdateChildContent "Country", fDbCh.GetChildContent("Country")
                dbCh.UpdateChildContent "Duration", fDbCh.GetChildContent("Duration")
                dbCh.UpdateChildContent "Directors", fDbCh.GetChildContent("Directors")
                dbCh.UpdateChildContent "Plot", fDbCh.GetChildContent("Plot")
                dbCh.UpdateChildContent "Cast", fDbCh.GetChildContent("Cast")
                dbCh.UpdateChildContent "HashFailed", "No"
                dbCh.UpdateChildContent "UpdatedOnce", "No"
                dbCh.UpdateChildContent "Fav", "No"
                
                Dim genCopy As New ChilkatXml
                Set genCopy = fDbCh.GetChildWithTag("Genres")
                dbCh.removeChild ("Genres")
                dbCh.AddChildTree genCopy
                
                
                Dim imgCopy As New FileSystemObject
                
                
                 
                If Not fDbCh.GetChildContent("CoverSmall") = "" Then
                
                    If imgCopy.FileExists(nObj.Content & "\MyMovieManager_Data_XYZ\Covers\" & fDbCh.GetChildContent("CoverSmall")) Then
                        imgCopy.CopyFile nObj.Content & "\MyMovieManager_Data_XYZ\Covers\" & fDbCh.GetChildContent("CoverSmall"), nObj.Content & "\MyMovieManager_Data_XYZ\Covers\" & dbCh.GetChildContent("UnID") & ".jpg", True
                        dbCh.UpdateChildContent "CoverSmall", dbCh.GetChildContent("UnID") & ".jpg"
                    End If
                    
                End If
                
                If Not fDbCh.GetChildContent("CoverLarge") = "" Then
                
                    If imgCopy.FileExists(nObj.Content & "\MyMovieManager_Data_XYZ\Covers\" & fDbCh.GetChildContent("CoverLarge")) Then
                        imgCopy.CopyFile nObj.Content & "\MyMovieManager_Data_XYZ\Covers\" & fDbCh.GetChildContent("CoverLarge"), nObj.Content & "\MyMovieManager_Data_XYZ\Covers\" & dbCh.GetChildContent("UnID") & "_large.jpg", True
                        dbCh.UpdateChildContent "CoverLarge", dbCh.GetChildContent("UnID") & "_large.jpg"
                    End If
                End If
                
            End If
            
            

            End If
            DoEvents
            db.SaveXml nObj.Content & "\MyMovieManager_Data_XYZ\data.xml"
            
            
        Next




'''''''''''''''''''''''''''''''''''''''

        DoEvents
        db.SaveXml nObj.Content & "\MyMovieManager_Data_XYZ\data.xml"
    Next

End If











If delEntry Then

    Dim minus As Long
    Dim skF As New FileSystemObject
    Dim sDB As New ChilkatXml
    Dim sDBc As New ChilkatXml
    Dim flNme
    deletedNo = 0
    minus = 0
    mObj.LoadXmlFile App.path & "\Data\Loc.xml"

    kj = mObj.NumChildrenHavingTag("Location")


    For i = 0 To kj - 1
        DoEvents
        Dim jk
        Set nObj = mObj.GetNthChildWithTag("Location", i)
        
        sDB.LoadXmlFile nObj.Content & "\MyMovieManager_Data_XYZ\data.xml"
        
        For jk = 0 To sDB.NumChildrenHavingTag("Movie") - 1
            Set sDBc = sDB.GetNthChildWithTag("Movie", jk - minus)
            flNme = nObj.Content & sDBc.GetChildContent("FileName")
            If Not skF.FileExists(flNme) Then
                sDB.RemoveChildByIndex jk - minus
                deletedNo = deletedNo + 1
                statusObj = str(deletedNo) & " Old Entries Removed..."
                minus = minus + 1
            End If
            
            sDB.SaveXml nObj.Content & "\MyMovieManager_Data_XYZ\data.xml"
        Next


        sDB.SaveXml nObj.Content & "\MyMovieManager_Data_XYZ\data.xml"

    Next

End If




   On Error GoTo 0
   Exit Function

updateDataBase_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ")  in procedure updateDataBase of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function









Public Function updateFromNet(onlyNew As Boolean, fileStatusObj As Control, statusObj As Control, progressB As ProgressBar, errorControl As Control)




   On Error GoTo updateFromNet_Error
   
   If InternetCheckConnection("http://www.google.com", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
   MsgBox "Not Connected to Internet, Please try After Connecting To Internet", vbInformation
   Exit Function
   End If

If Not moviesNotUpdated = 0 Then
progressB.Max = moviesNotUpdated
Else
progressB.Max = 1
End If
Dim mObj As New ChilkatXml
Dim nObj As New ChilkatXml
Dim i, kj




Dim neT As New ChilkatXml
Dim neTc As New ChilkatXml
Dim iNet As Long
Dim Location
mObj.LoadXmlFile App.path & "\Data\Loc.xml"

progressM = 0


kj = mObj.NumChildrenHavingTag("Location")


For i = 0 To kj - 1

    Set nObj = mObj.GetNthChildWithTag("Location", i)
    Location = nObj.Content
    Dim skC As New FileSystemObject
    If Not skC.FolderExists(Location & "\MyMovieManager_Data_XYZ") Then skC.CreateFolder (Location & "\MyMovieManager_Data_XYZ")
    If Not skC.FolderExists(Location & "\MyMovieManager_Data_XYZ\Covers") Then skC.CreateFolder (Location & "\MyMovieManager_Data_XYZ\Covers")
    If skC.FolderExists(Location & "\MyMovieManager_Data_XYZ") Then
        Dim mmmK As Folder
        Set mmmK = skC.GetFolder(Location & "\MyMovieManager_Data_XYZ")
        mmmK.Attributes = Hidden + Directory + System
    End If


    neT.LoadXmlFile Location & "\MyMovieManager_Data_XYZ\data.xml"
    'MsgBox Location & "\MyMovieManager_Data_XYZ\data.xml" 'neT.NumChildrenHavingTag("Movie") - 1

    For iNet = 0 To neT.NumChildrenHavingTag("Movie") - 1
        
        If stopSignalNet = True Then
            stopSignalNet = False
            GoTo stopUpdating
        End If
        Set neTc = neT.GetNthChildWithTag("Movie", iNet)

        If neTc.GetChildContent("UpdatedOnce") = "Yes" And onlyNew Then
            GoTo finish
        ElseIf neTc.GetChildContent("HashFailed") = "No" And neTc.GetAttrValue("isMovie") = "Yes" Then
            statusObj = "Checking " & Location & neTc.GetChildContent("FileName")
        If InternetCheckConnection("http://www.google.com", FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
            MsgBox "Not Connected to Internet, Please try After Connecting To Internet", vbInformation
            Exit Function
        Else
            Dim loginData() As String
            Dim hashResponse() As String
            Dim movieData() As String
            loginData = xmlServerActions.loginXml()
            DoEvents
            If loginData(0) = "200 OK" Then
                If Not neTc.GetChildContent("Hash") = "" Then
                    hashResponse = xmlServerActions.checkMovieHashXml(loginData(3), loginData(1), neTc.GetChildContent("Hash"))
                    DoEvents
                        If hashResponse(0) = "200 OK" And (Not hashResponse(1) = "") Then
                            movieData = xmlServerActions.getIMDBdetailsXml(loginData(3), loginData(1), hashResponse(1))
    
    
    
                                If movieData(0) = "200 OK" Then
                                    
                                    errorControl = ""
    
                                    neTc.UpdateChildContent "ImdbID", movieData(8)
                                    neTc.UpdateChildContent "Title", movieData(1)
                                    neTc.UpdateChildContent "Year", movieData(2)
                                    neTc.UpdateChildContent "IMDBRating", movieData(3)
                                    neTc.UpdateChildContent "Languages", movieData(4)
                                    neTc.UpdateChildContent "Country", movieData(5)
                                    neTc.UpdateChildContent "Duration", movieData(6)
                                    neTc.UpdateChildContent "Directors", movieData(7)
                                    neTc.UpdateChildContent "Plot", movieData(9)
                                    neTc.removeChild ("Genres")
                                    Dim myGenre As New ChilkatXml
                                    myGenre.loadXML "<Genres>" & movieData(10) & "</Genres>"
                                    neTc.AddChildTree myGenre
                                    'neTc.UpdateChildContent "Genres", movieData(10)
                                    neTc.UpdateChildContent "Cast", movieData(13)
                                    neTc.UpdateChildContent "UpdatedOnce", "Yes"
                                    'MsgBox "here"
                                    If movieData(11) = "" Then
                    
                                        neTc.UpdateChildContent "CoverSmall", ""
                                        neTc.UpdateChildContent "CoverLarge", ""
            
                                    Else
    
                                        Dim errcode As Long
                                        Dim urlL As String
                                        Dim localFileName As String
                                        urlL = movieData(11)
                                        localFileName = Location & "\MyMovieManager_Data_XYZ\Covers\" & neTc.GetChildContent("UnID") & ".jpg"
                                        'MsgBox localFileName
                                        errcode = URLDownloadToFile(0, urlL, localFileName, 0, 0)
    
                                        If errcode = 0 Then
    
                                            neTc.UpdateChildContent "CoverSmall", neTc.GetChildContent("UnID") & ".jpg"
                                        Else
                                            errorControl = "Error in Connection..."
    
                                        End If
    
                                        urlL = movieData(12)
                                        localFileName = Location & "\MyMovieManager_Data_XYZ\Covers\" & neTc.GetChildContent("UnID") & "_large.jpg"
                                        errcode = URLDownloadToFile(0, urlL, localFileName, 0, 0)
    
    
    
                                        If errcode = 0 Then
    
                                            neTc.UpdateChildContent "CoverLarge", neTc.GetChildContent("UnID") & "_large.jpg"
                                            errorControl = ""
                                        Else
                                            errorControl = "Error in Connection..."
    
                                        End If
    
                                    End If
    
                                End If
                                errorControl = ""
                            Else
                                errorControl = "Error in Connection..."
                            End If
    
                        End If
                        errorControl = ""
                    Else
                        errorControl = "Error in Connection..."
                End If



           


        End If
        If Not progressB.Max < progressM Then progressB.Value = progressM
        fileStatusObj = str(progressM) & " VideoFiles Checked" & str(moviesNotUpdated - progressM) & " More Video Files to Be Checked."
        progressM = progressM + 1
        
        ElseIf Not neTc.GetChildContent("ImdbID") = "" Then
        End If


        neT.SaveXml Location & "\MyMovieManager_Data_XYZ\data.xml"
        
        

        'frmMain_Pg.loadAllMovies
finish: DoEvents

    Next

    neT.SaveXml Location & "\MyMovieManager_Data_XYZ\data.xml"
    DoEvents

Next

stopUpdating:
neT.SaveXml Location & "\MyMovieManager_Data_XYZ\data.xml"
DoEvents
fileStatusObj = "Done Updating"
'MsgBox "Done Updating...", vbInformation
progressB.Value = progressB.Max




   On Error GoTo 0
   Exit Function

updateFromNet_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure updateFromNet of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function





Function loadAllMovies(statusObj As Control)

Dim savePref As New ChilkatXml
   On Error GoTo loadAllMovies_Error

savePref.loadXML App.path & "\Data\Preferences.xml"
totalMovies = savePref.GetChildContent("TotalMovies")

ReDim movieArray(0) As MovieInfo
moviesNotUpdated = 0

ReDim imdbArray(0) As String

Dim locationDB As New ChilkatXml
Dim locationDBCh As New ChilkatXml
Dim movieData As New ChilkatXml
Dim movieDataCh As New ChilkatXml
Dim i, j, l, noLo, noMo


Dim sk As New FileSystemObject
locationDB.LoadXmlFile App.path & "\Data\Loc.xml"
noLo = locationDB.NumChildrenHavingTag("Location")

For i = 0 To noLo - 1

DoEvents
    Set locationDBCh = locationDB.GetNthChildWithTag("Location", i)
    
    If sk.FileExists(locationDBCh.Content & "\MyMovieManager_Data_XYZ\data.xml") Then
        movieData.LoadXmlFile locationDBCh.Content & "\MyMovieManager_Data_XYZ\data.xml"
        
        
        '''''''''''''''''''''''''''''Check out follwing line''''''''''''''''''''''''''''
        
         movieData.loadXML htmlDecode(movieData.GetXml)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
        
        
        noMo = movieData.NumChildrenHavingTag("Movie")
        'MsgBox noMo
        For j = 0 To noMo - 1
            
            DoEvents
            Set movieDataCh = movieData.GetNthChildWithTag("Movie", j)
            If (movieDataCh.GetChildContent("UpdatedOnce") = "No" And movieDataCh.GetAttrValue("isMovie") = "Yes") Then moviesNotUpdated = moviesNotUpdated + 1
            
            Dim nIMDBc, imdbPresent
            imdbPresent = 0
            
            For nIMDBc = 0 To UBound(imdbArray) - 1
                If imdbArray(nIMDBc) = movieDataCh.GetChildContent("ImdbID") And (Not movieDataCh.GetChildContent("ImdbID") = "") Then
                    imdbPresent = 1
                End If
            Next
            
            If (Not movieDataCh.GetChildContent("FileName") = "") And movieDataCh.GetAttrValue("isMovie") = "Yes" And imdbPresent = 0 Then
            
            
                
                If Not movieDataCh.GetChildContent("Title") = "" Then
                    movieArray(UBound(movieArray)).MovieTitle = movieDataCh.GetChildContent("Title")
                Else
                    l = InStrRev(movieDataCh.GetChildContent("FileName"), "\")
                    movieArray(UBound(movieArray)).MovieTitle = Right(movieDataCh.GetChildContent("FileName"), Len(movieDataCh.GetChildContent("FileName")) - l)
                End If
                
                    imdbArray(UBound(imdbArray)) = movieDataCh.GetChildContent("ImdbID")
                    ReDim Preserve imdbArray(UBound(imdbArray) + 1) As String
                    
                    movieArray(UBound(movieArray)).MovieFileName = movieDataCh.GetChildContent("FileName")
                    movieArray(UBound(movieArray)).MovieWatched = movieDataCh.GetChildContent("Watched")
                    movieArray(UBound(movieArray)).MovieHash = movieDataCh.GetChildContent("Hash")
                    movieArray(UBound(movieArray)).MovieIMDBid = movieDataCh.GetChildContent("ImdbID")
                    movieArray(UBound(movieArray)).MovieYear = movieDataCh.GetChildContent("Year")
                    movieArray(UBound(movieArray)).MovieIMDbRating = movieDataCh.GetChildContent("IMDBRating")
                    movieArray(UBound(movieArray)).MovieMyRating = movieDataCh.GetChildContent("MyRating")
                    movieArray(UBound(movieArray)).MovieLanguage = movieDataCh.GetChildContent("Languages")
                    movieArray(UBound(movieArray)).MovieCountry = movieDataCh.GetChildContent("Country")
                    movieArray(UBound(movieArray)).MovieDuration = movieDataCh.GetChildContent("Duration")
                    movieArray(UBound(movieArray)).MovieDirector = movieDataCh.GetChildContent("Directors")
                    movieArray(UBound(movieArray)).MoviePlot = movieDataCh.GetChildContent("Plot")
                    movieArray(UBound(movieArray)).MovieIsFav = movieDataCh.GetChildContent("Fav")
                    movieArray(UBound(movieArray)).DateAdded = CDate(movieDataCh.GetChildContent("DateAdded"))
                    Dim skMine As New ChilkatXml
                    Set skMine = movieDataCh.GetChildWithTag("Genres")
                    Dim igen, genStr
                    
                    ReDim movieArray(UBound(movieArray)).MovieGenre(0 To skMine.NumChildrenHavingTag("value")) As String
                    
                    
                    
                    For igen = 0 To skMine.NumChildrenHavingTag("value") - 1
                     
                        Set skMine = skMine.GetNthChildWithTag("value", igen)
                        genStr = skMine.Content
                        movieArray(UBound(movieArray)).MovieGenre(igen) = genStr
                        Set skMine = skMine.getParent
                    
                       
                    Next
                    
                    
                    
                    movieArray(UBound(movieArray)).MovieCoverSmall = movieDataCh.GetChildContent("CoverSmall")
                    movieArray(UBound(movieArray)).MovieCoverLarge = movieDataCh.GetChildContent("CoverLarge")
                    movieArray(UBound(movieArray)).MovieCast = movieDataCh.GetChildContent("Cast")
                    movieArray(UBound(movieArray)).MovieHashFailed = movieDataCh.GetChildContent("HashFailed")
                    movieArray(UBound(movieArray)).MovieUpdatedOnce = movieDataCh.GetChildContent("UpdatedOnce")
                    movieArray(UBound(movieArray)).MovieLocFolder = locationDBCh.Content
                    movieArray(UBound(movieArray)).MovieUnID = movieDataCh.GetChildContent("UnID")
                    movieArray(UBound(movieArray)).isMovie = "Yes"
                    movieArray(UBound(movieArray)).wacthedCategory = True
                    
                    Dim chkImg As New FileSystemObject
                    If movieDataCh.GetChildContent("CoverSmall") = "" Or (chkImg.FileExists(locationDBCh.Content & "\MyMovieManager_Data_XYZ\Covers\" & movieDataCh.GetChildContent("CoverSmall")) = False) Then
                         Set movieArray(UBound(movieArray)).MovieIcon = LoadPicture(App.path & "\Images\defaultCoverSmall.jpg")
                    Else
                         Set movieArray(UBound(movieArray)).MovieIcon = LoadPicture(locationDBCh.Content & "\MyMovieManager_Data_XYZ\Covers\" & movieDataCh.GetChildContent("CoverSmall"))
                    End If
                    
                    
                    movieArray(UBound(movieArray)).MovieSearchRelevance = -1
                    movieArray(UBound(movieArray)).movieIdentifier = UBound(movieArray)
                    movieArray(UBound(movieArray)).MovieDisplayFlag = True
                    
'                    MsgBox movieArray(UBound(movieArray)).MovieFileName & vbCrLf & str(UBound(movieArray))
                    
                    ReDim Preserve movieArray(UBound(movieArray) + 1) As MovieInfo
                
            End If
            
            DoEvents
            
            If Not totalMovies = 0 Then statusObj = "Loading Movies... " & str(CInt(UBound(movieArray) * 100 / totalMovies)) & "%"
            
            DoEvents
        Next
    
    
    
    
    
    End If
    




Next




savePref.loadXML App.path & "\Data\Preferences.xml"
savePref.UpdateChildContent "TotalMovies", str(UBound(movieArray))
savePref.SaveXml App.path & "\Data\Preferences.xml"

statusObj = "Loading Movies... 100%"

   On Error GoTo 0
   Exit Function

loadAllMovies_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure loadAllMovies of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function

Function searchMovies(searchText As String, searchBy As String)

Dim numMovie As Long
Dim i
   On Error GoTo searchMovies_Error

numMovie = UBound(movieArray) - LBound(movieArray)


If Not searchText = "" Then

    Select Case searchBy

        Case "Title"

            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MovieTitle, searchText, vbTextCompare)
            Next

        Case "Director"
            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MovieDirector, searchText, vbTextCompare)
            Next

        Case "Cast"
            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MovieCast, searchText, vbTextCompare)
                
            Next

        Case "Plot"
            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MoviePlot, searchText, vbTextCompare)
                
            Next

        Case "Year"
            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MovieYear, searchText, vbTextCompare)
                
            Next
    
        Case "Language"
            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MovieLanguage, searchText, vbTextCompare)
                
            Next

        Case "Country"
            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MovieCountry, searchText, vbTextCompare)
                
            Next

        Case "IMDb ID"
            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MovieIMDBid, searchText, vbTextCompare)

            Next
        Case "FileName"
            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MovieFileName, searchText, vbTextCompare)

            Next
            
        Case "IMDb Rating"
            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MovieIMDbRating, searchText, vbTextCompare)

            Next
            
        Case "Your Rating"
            For i = 0 To numMovie
                movieArray(i).MovieSearchRelevance = InStr(1, movieArray(i).MovieMyRating, searchText, vbTextCompare)

            Next


    End Select


Else
    For i = LBound(movieArray) To UBound(movieArray)
        movieArray(i).MovieSearchRelevance = -1
    Next

End If


   On Error GoTo 0
   Exit Function

searchMovies_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure searchMovies of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function

Public Static Sub StrSort(words() As String, Ascending As Boolean, AllLowerCase As Boolean)
 
'Pass in string array you want to sort by reference and
'read it back

'Set Ascending to True to sort ascending, '
'false to sort descending

'If AllLowerCase is True, strings will be sorted
'without regard to case.  Otherwise, upper
'case characters take precedence over lower
'case characters

Dim i As Integer
Dim j As Integer
Dim NumInArray, LowerBound As Integer
   On Error GoTo StrSort_Error

NumInArray = UBound(words)
LowerBound = LBound(words)
For i = LowerBound To NumInArray
    j = 0
    For j = LowerBound To NumInArray
        If AllLowerCase = True Then
            If Ascending = True Then
                If StrComp(LCase(words(i)), _
                     LCase(words(j))) = -1 Then
                    Call Swap(words(i), words(j))
                End If
            Else
                If StrComp(LCase(words(i)), _
                       LCase(words(j))) = 1 Then
                    Call Swap(words(i), words(j))
                End If
            End If
        Else
            If Ascending = True Then
                If StrComp(words(i), words(j)) = -1 Then
                    Call Swap(words(i), words(j))
                End If
            Else
                If StrComp(words(i), _
                    words(j)) = 1 Then
                    Call Swap(words(i), words(j))
                End If
            End If
        End If
    Next j
Next i

   On Error GoTo 0
   Exit Sub

StrSort_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure StrSort of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub





Private Sub Swap(var1 As String, var2 As String)
    Dim x As String
   On Error GoTo Swap_Error

    x = var1
    var1 = var2
    var2 = x

   On Error GoTo 0
   Exit Sub

Swap_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Swap of Module updateFiles" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

