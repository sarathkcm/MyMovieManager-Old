Attribute VB_Name = "xmlServerActions"
Option Explicit
Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Const FLAG_ICC_FORCE_CONNECTION = &H1
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Long

    

Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
Public Function LoadPicture(ByVal strFileName As String) As Picture
Dim IID  As TGUID
   On Error GoTo LoadPicture_Error

    With IID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
On Error GoTo ERR_LINE
    OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
    Exit Function
ERR_LINE:
    Set LoadPicture = VB.LoadPicture(strFileName)

   On Error GoTo 0
   Exit Function

LoadPicture_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadPicture of Module xmlServerActions" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Function








Public Function getResponse(url As String, myXml As String) As String


On Error GoTo ErRa

'If InternetCheckConnection(url, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
'getResponse = ""
'Else

Dim m As String
m = ""

'HTTP variable
Dim myHTTP As MSXML2.XMLHTTP

'HTTP object
Set myHTTP = CreateObject("msxml2.xmlhttp")

'create dom document variable  ‘stores the xml to send
Dim myDom As MSXML2.DOMDocument

'Create the DomDocument Object
Set myDom = CreateObject("MSXML2.DOMDocument")

'Load entire Document before moving on
myDom.async = True

'xml string variable
'replace with location if sending from file or URL


'loads the xml
'change to .Load for file or url
myDom.loadXML (myXml)
 
'open the connection
myHTTP.open "post", url, True 'False true for asynchronous

myHTTP.setRequestHeader "Content-Type", "text/xml;charset=utf-8"
myHTTP.setRequestHeader "Connection", "keep-alive"


myDom.async = True
myDom.loadXML myXml
myHTTP.send myDom.xml

Do
DoEvents
If stopSignalNet = True Then GoTo ErRa
Loop While myHTTP.readyState <> 4




'Text6 = myDom.xml
'send the XML
'myHTTP.send (myDom.xml)

'Display the response
'while myhttp.status

m = myHTTP.responseText

Set myHTTP = Nothing
getResponse = m

'End If


Exit Function
ErRa:
On Error GoTo 0
Set myHTTP = Nothing
getResponse = ""

    If Not Err.Number = 0 Then
    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure getResponse of Module xmlServerActions" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source
    End If
End Function





Public Function loginXml() As String()
 

 Dim str(3) As String
 Dim ms As String, url As String, usrAgt As String, usr As String, pass As String, lang As String
 Dim prefData As New ChilkatXml
   On Error GoTo loginXml_Error

 prefData.LoadXmlFile App.path & "\Data\Preferences.xml"
 url = prefData.GetChildContent("apiURL")
 usrAgt = prefData.GetChildContent("UserAgent")
 usr = ""
 pass = ""
 lang = "en"
 'MsgBox prefData.GetXml
 
'If InternetCheckConnection(url, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
'str(0) = "Not OK"
'str(1) = ""
'str(2) = ""
'str(3) = ""
'loginXml = str
'Exit Function
'End If
 
 
 
 
 ms = "<methodCall>" & _
 "<methodName>LogIn</methodName>" & _
 "<params>" & _
  "<param>" & _
   "<value><string>" & usr & "</string></value>" & _
  "</param>" & _
  "<param>" & _
   "<value><string>" & pass & "</string></value>" & _
  "</param>" & _
  "<param>" & _
   "<value><string>" & lang & "</string></value>" & _
  "</param>" & _
  "<param>" & _
   "<value><string>" & usrAgt & "</string></value>" & _
  "</param>" & _
 "</params>" & _
"</methodCall>"


Dim chk As New ChilkatXml
chk.loadXML ms
ms = chk.GetXml

Dim xF As New ChilkatXml
Dim xFc As New ChilkatXml
Dim i

xF.loadXML getResponse(url, ms)

Set xF = xF.GetChildWithTag("params")
Set xF = xF.GetChildWithTag("param")
Set xF = xF.GetChildWithTag("value")
Set xF = xF.GetChildWithTag("struct")




For i = 0 To 2
Set xFc = xF.GetNthChildWithTag("member", i)

Select Case (xFc.GetChildContent("name"))
Case "token"
Set xFc = xFc.GetChildWithTag("value")
    str(1) = xFc.GetChildContent("string")
    
Case "status"
Set xFc = xFc.GetChildWithTag("value")
    str(0) = xFc.GetChildContent("string")
    
Case "seconds"
   Set xFc = xFc.GetChildWithTag("value")
    str(2) = xFc.GetChildContent("double")
    
Case Else
    str(0) = "0"
    
End Select

Next
str(3) = url

loginXml = str

'str(0)=status
'str(1)=token
'str(2)=seconds
'3 = url



   On Error GoTo 0
   Exit Function

loginXml_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure loginXml of Module xmlServerActions" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function


Public Function checkMovieHashXml(url As String, token As String, Hash As String) As String()
Dim str(4) As String

'If InternetCheckConnection(url, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
'str(0) = "Not OK"
'str(1) = ""
'str(2) = ""
'str(3) = ""
'str(4) = ""
'checkMovieHashXml = str
'Exit Function
'End If
Dim ms As String


   On Error GoTo checkMovieHashXml_Error

ms = "<methodCall>" & _
 "<methodName>CheckMovieHash</methodName>" & _
 "<params>" & _
  "<param>" & _
   "<value><string>" & token & "</string></value>" & _
  "</param>" & _
  "<param>" & _
   "<value>" & _
    "<array>" & _
     "<data>" & _
      "<value><string>" & Hash & "</string></value>" & _
    "</data>" & _
    "</array>" & _
   "</value>" & _
  "</param>" & _
 "</params>" & _
"</methodCall>"

Dim chk As New ChilkatXml
chk.loadXML ms
ms = chk.GetXml

'Text6 = ms


''''''''''''
'MsgBox ms
'''''''''''


Dim xF As New ChilkatXml
Dim xFc As New ChilkatXml
Dim xFcC As New ChilkatXml
Dim i


xF.loadXML getResponse(url, ms)


Set xF = xF.GetChildWithTag("params")
Set xF = xF.GetChildWithTag("param")
Set xF = xF.GetChildWithTag("value")
Set xF = xF.GetChildWithTag("struct")


Dim skcm

For i = 0 To xF.NumChildrenHavingTag("member") - 1

Set xFc = xF.GetNthChildWithTag("member", i)

Select Case (xFc.GetChildContent("name"))


Case "status"
Set xFc = xFc.GetChildWithTag("value")
    str(0) = xFc.GetChildContent("string")
    
Case "data"
   
    Set xFc = xFc.GetChildWithTag("value")
   
    Set xFc = xFc.GetChildWithTag("struct")
   
    Set xFc = xFc.GetChildWithTag("member")
   
    Set xFc = xFc.GetChildWithTag("value")
   
    Set xFc = xFc.GetChildWithTag("struct")
   
For skcm = 0 To xFc.NumChildrenHavingTag("member") - 1
            
          Set xFcC = xFc.GetNthChildWithTag("member", skcm)


  
            
            Select Case (xFcC.GetChildContent("name"))
            
            Case "MovieImdbID"
            Set xFcC = xFcC.GetChildWithTag("value")
            
            str(1) = xFcC.GetChildContent("string")
            
            
            Case "MovieName"
            Set xFcC = xFcC.GetChildWithTag("value")
            
            str(2) = xFcC.GetChildContent("string")
            
            Case "MovieHash"
            Set xFcC = xFcC.GetChildWithTag("value")
            
            str(3) = xFcC.GetChildContent("string")
            
            End Select
            
            
            
            
Next
    
    
Case "seconds"
   Set xFc = xFc.GetChildWithTag("value")
    str(4) = xFc.GetChildContent("double")


    
End Select

Next

checkMovieHashXml = str

'0->status
'1->movie IMDB ID - is "" if not identified
'2->movie name - is "" if not identified
'3->movie hash
'4->seconds

   On Error GoTo 0
   Exit Function

checkMovieHashXml_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure checkMovieHashXml of Module xmlServerActions" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function

Public Function getIMDBdetailsXml(url As String, token As String, id As String) As String()
Dim i, j
Dim moneLaddu
Dim str(13) As String

'If InternetCheckConnection(url, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then
'str(0) = "Not OK"
'For moneLaddu = 1 To 13
'str(moneLaddu) = ""
'Next
'Exit Function
'End If

Dim ms As String
   On Error GoTo getIMDBdetailsXml_Error

ms = "<methodCall>" & _
 "<methodName>GetIMDBMovieDetails</methodName>" & _
 "<params>" & _
  "<param>" & _
   "<value><string>" & token & "</string></value>" & _
  "</param>" & _
  "<param>" & _
   "<value><string>" & id & "</string></value>" & _
  "</param>" & _
"</params>" & _
"</methodCall>"
'getIMDBdetailsXml = getResponse(url, ms)
Dim chk As New ChilkatXml
chk.loadXML ms
ms = chk.GetXml


Dim xF As New ChilkatXml
Dim xFc As New ChilkatXml
Dim xFcC As New ChilkatXml
Dim nCh As New ChilkatXml
Dim v2 As New ChilkatXml



xF.loadXML htmlDecode(getResponse(url, ms))

Set xF = xF.GetChildWithTag("params")
Set xF = xF.GetChildWithTag("param")
Set xF = xF.GetChildWithTag("value")
Set xF = xF.GetChildWithTag("struct")

Dim i2, j2

For i = 0 To xF.NumChildrenHavingTag("member") - 1

Set xFc = xF.GetNthChildWithTag("member", i)

Select Case (xFc.GetChildContent("name"))


Case "status"
Set xFc = xFc.GetChildWithTag("value")
    str(0) = xFc.GetChildContent("string")
    
Case "data"

 
    Set xFc = xFc.GetChildWithTag("value")
   
    Set xFc = xFc.GetChildWithTag("struct")
    
    For j = 0 To xFc.NumChildrenHavingTag("member") - 1
    Set xFcC = xFc.GetNthChildWithTag("member", j)
    
    Select Case (xFcC.GetChildContent("name"))
    Case "title"
    Set nCh = xFcC.GetChildWithTag("value")
    
    str(1) = nCh.GetChildContent("string")
    Case "year"
    Set nCh = xFcC.GetChildWithTag("value")
    
    str(2) = nCh.GetChildContent("string")
    
    Case "rating"
    Set nCh = xFcC.GetChildWithTag("value")
    
    str(3) = nCh.GetChildContent("string")
    
    
    
    
    Case "language"
   Set nCh = xFcC.GetChildWithTag("value")
    Set nCh = nCh.GetChildWithTag("array")
    Set nCh = nCh.GetChildWithTag("data")
    str(4) = ""
    
    For i2 = 0 To nCh.NumChildrenHavingTag("value") - 1
    Set v2 = nCh.GetNthChildWithTag("value", i2)
    
    
    str(4) = str(4) & v2.GetChildContent("string") & ", "
    Next
    
    If Right(str(4), 2) = ", " Then
    str(4) = Left(str(4), Len(str(4)) - 2)
    
    
    End If
    
    
    Case "country"
    Set nCh = xFcC.GetChildWithTag("value")
    Set nCh = nCh.GetChildWithTag("array")
    Set nCh = nCh.GetChildWithTag("data")
    str(5) = ""
    
    For i2 = 0 To nCh.NumChildrenHavingTag("value") - 1
    Set v2 = nCh.GetNthChildWithTag("value", i2)
    
    
    str(5) = str(5) & v2.GetChildContent("string") & ", "
    Next
    
    If Right(str(5), 2) = ", " Then
    str(5) = Left(str(5), Len(str(5)) - 2)
    End If
    
    Case "duration"
    Set nCh = xFcC.GetChildWithTag("value")
    
    str(6) = nCh.GetChildContent("string")
    
    Case "directors"
    
    Set nCh = xFcC.GetChildWithTag("value")
    Set nCh = nCh.GetChildWithTag("struct")
    str(7) = ""
    For j2 = 0 To nCh.NumChildrenHavingTag("member") - 1
    Set v2 = nCh.GetNthChildWithTag("member", j2)
    Set v2 = v2.GetChildWithTag("value")
    str(7) = str(7) & v2.GetChildContent("string") & ", "
    Next
    
    If Right(str(7), 2) = ", " Then
    str(7) = Left(str(7), Len(str(7)) - 2)
    End If
    
    
    
    Case "id"
    Set nCh = xFcC.GetChildWithTag("value")
    
    str(8) = nCh.GetChildContent("string")
    
    Case "plot"
    Set nCh = xFcC.GetChildWithTag("value")
    
    str(9) = nCh.GetChildContent("string")
    
    
    Case "genres"
    Set nCh = xFcC.GetChildWithTag("value")
    Set nCh = nCh.GetChildWithTag("array")
    Set nCh = nCh.GetChildWithTag("data")
    str(10) = ""
    
    For i2 = 0 To nCh.NumChildrenHavingTag("value") - 1
    Set v2 = nCh.GetNthChildWithTag("value", i2)
    
    
    str(10) = str(10) & "<value>" & v2.GetChildContent("string") & "</value>" & vbNewLine
    Next
    
    Case "cover"
    Set nCh = xFcC.GetChildWithTag("value")
    
    str(11) = nCh.GetChildContent("string")
    
    
    Case "cast"
    
    Set nCh = xFcC.GetChildWithTag("value")
    Set nCh = nCh.GetChildWithTag("struct")
    str(13) = ""
    For j2 = 0 To nCh.NumChildrenHavingTag("member") - 1
    Set v2 = nCh.GetNthChildWithTag("member", j2)
    Set v2 = v2.GetChildWithTag("value")
    str(13) = str(13) & v2.GetChildContent("string") & ", "
    Next
    
    If Right(str(13), 2) = ", " Then
    str(13) = Left(str(13), Len(str(13)) - 2)
    End If
    
    
    End Select
    
    
    Next
    
    
    
    
    

End Select
Next

str(12) = ""
Dim cov As String, covLen As Long
If Not str(11) = "" Then
cov = str(11)
covLen = InStr(cov, "._")
cov = Left(cov, covLen - 1)
str(11) = cov & "._V1._SY150_CR0,0,102,150_.jpg"
str(12) = cov & "._V1._SY317_CR0,0,214,317_.jpg"
End If




getIMDBdetailsXml = str


   On Error GoTo 0
   Exit Function

getIMDBdetailsXml_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure getIMDBdetailsXml of Module xmlServerActions" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Function

Public Function logoutXml(url As String, token As String) As String
'If InternetCheckConnection(url, FLAG_ICC_FORCE_CONNECTION, 0&) = 0 Then

'Exit Function
'End If

Dim ms As String

   On Error GoTo logoutXml_Error

ms = "<methodCall>" & _
 "<methodName>LogOut</methodName>" & _
 "<params>" & _
  "<param>" & _
   "<value><string>" & token & "</string></value>" & _
  "</param>" & _
 "</params>" & _
"</methodCall>"
logoutXml = getResponse(url, ms)

   On Error GoTo 0
   Exit Function

logoutXml_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure logoutXml of Module xmlServerActions" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Function
Public Function htmlDecode(strQ As String) As String
Dim i As Long

   On Error GoTo htmlDecode_Error

For i = 1 To 255
strQ = Replace(strQ, "&#" & i & ";", ChrW(i), , , vbTextCompare)
strQ = Replace(strQ, "&amp;#" & i & ";", ChrW(i), , , vbTextCompare)
Next
strQ = Replace(strQ, "See Full Summary", "", , , vbTextCompare)
htmlDecode = strQ

   On Error GoTo 0
   Exit Function

htmlDecode_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure htmlDecode of Module xmlServerActions" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Function


