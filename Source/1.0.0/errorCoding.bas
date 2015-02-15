Attribute VB_Name = "errorCoding"
Public Function writeError(strF As String)

   On Error GoTo writeError_Error

Dim errorF As Integer
errorF = FreeFile
Open App.path & "\" & App.EXEName & ".log" For Append As errorF
Print #errorF,
Print #errorF, strF
Print #errorF, "_______________________________________________________________"
Close #errorF
   On Error GoTo 0
   Exit Function

writeError_Error:

    MsgBox "Failed to Create Log File" & vbCrLf & "Error " & Err.Number & " (" & Err.Description & ")"

End Function

Public Function ReportF(strF As String)

   On Error GoTo writeError_Error

Dim errorF As Integer
errorF = FreeFile
Open App.path & "\" & App.EXEName & ".log" For Append As errorF
Print #errorF,
Print #errorF, strF
Close #errorF
   On Error GoTo 0
   Exit Function

writeError_Error:

    MsgBox "Failed to Create Log File" & vbCrLf & "Error " & Err.Number & " (" & Err.Description & ")"

End Function

