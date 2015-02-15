Attribute VB_Name = "forThemes"
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error GoTo InitCommonControlsVB_Error

   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0

   On Error GoTo 0
   Exit Function

InitCommonControlsVB_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure InitCommonControlsVB of Module forThemes" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Function

Public Sub Main()
   On Error GoTo Main_Error

   InitCommonControlsVB
   splaSh.Show

   On Error GoTo 0
   Exit Sub

Main_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Module forThemes" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source


End Sub

