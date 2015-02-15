VERSION 5.00
Begin VB.Form splaSh 
   BorderStyle     =   0  'None
   Caption         =   "My Movie Manager Initializing"
   ClientHeight    =   5895
   ClientLeft      =   7755
   ClientTop       =   2745
   ClientWidth     =   6000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "splaSh.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "splaSh.frx":C84A
   ScaleHeight     =   5895
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   3480
   End
   Begin VB.PictureBox picBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6030
      Left            =   5880
      Picture         =   "splaSh.frx":81B8E
      ScaleHeight     =   6000
      ScaleWidth      =   6000
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Label vInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "version"
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
      Left            =   4440
      TabIndex        =   4
      Top             =   5040
      Width           =   600
   End
   Begin VB.Label ProgressL 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Movies... 35%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   120
      Picture         =   "splaSh.frx":F6ED2
      Top             =   4680
      Width           =   4500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sarath KCM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   5400
      Width           =   2895
   End
End
Attribute VB_Name = "splaSh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long


Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR = 2






Dim lngRegion As Long
Dim isFirstStart As Boolean









Private Sub Form_DblClick()
   On Error GoTo Form_DblClick_Error

Unload Me

   On Error GoTo 0
   Exit Sub

Form_DblClick_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_DblClick of Form splaSh" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    Me.Caption = "My Movie Manager v" & App.Major & "." & App.Minor & "." & App.Revision
    Dim lngRetr As Long
    lngRegion& = RegionFromBitmap(picBox)
    lngRetr& = SetWindowRgn(Me.hWnd, lngRegion&, True)
    

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form splaSh" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

    
End Sub


Private Sub Form_initialize()
Dim flDr As New FileSystemObject
   On Error GoTo Form_initialize_Error

isFirstStart = flDr.FileExists(App.path & "\Data\FirstRun - Copy.xml")
ProgressL.Caption = ""
vInfo = " version " & App.Major & "." & App.Minor & "." & App.Revision

   On Error GoTo 0
   Exit Sub

Form_initialize_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_initialize of Form splaSh" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub




Private Sub Timer1_Timer()
   On Error GoTo Timer1_Timer_Error
   ReportF "*****************************************************************************" & vbCrLf & "Started Program..."
   
Timer1.Interval = 0


If Not isFirstStart Then
'If normal start
'Timer1.Interval = 1000
mainF.init ProgressL
mainF.Show
Else
'If first start
  Fwizard.Show
End If

Unload Me

   On Error GoTo 0
   Exit Sub

Timer1_Timer_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer1_Timer of Form splaSh" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Sub


Private Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
  Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
  Dim lngRgnFinal As Long, lngRgnTmp As Long
  Dim lngStart As Long, lngRow As Long
  Dim lngCol As Long
   On Error GoTo RegionFromBitmap_Error

  If lngTransColor& < 1 Then
    lngTransColor& = GetPixel(picSource.hdc, 0, 0)
  End If
  lngHeight& = picSource.Height / Screen.TwipsPerPixelY
  lngWidth& = picSource.Width / Screen.TwipsPerPixelX
  lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
  For lngRow& = 0 To lngHeight& - 1
    lngCol& = 0
    Do While lngCol& < lngWidth&
      Do While lngCol& < lngWidth& And GetPixel(picSource.hdc, lngCol&, lngRow&) = lngTransColor&
        lngCol& = lngCol& + 1
      Loop
      If lngCol& < lngWidth& Then
        lngStart& = lngCol&
        Do While lngCol& < lngWidth& And GetPixel(picSource.hdc, lngCol&, lngRow&) <> lngTransColor&
          lngCol& = lngCol& + 1
        Loop
        If lngCol& > lngWidth& Then lngCol& = lngWidth&
        lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
        lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
        DeleteObject (lngRgnTmp&)
      End If
    Loop
  Next
  RegionFromBitmap& = lngRgnFinal&

   On Error GoTo 0
   Exit Function

RegionFromBitmap_Error:

    writeError "Error " & Err.Number & " (" & Err.Description & ") in procedure RegionFromBitmap of Form splaSh" & vbCrLf & "HelpContext = " & Err.HelpContext & "   Source = " & Err.Source

End Function

