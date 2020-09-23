Attribute VB_Name = "Module1"
Option Explicit
'
' Module Declares for the StatProgressBar example
'
' Chris Eastwood Feb. 1999
'
' http://www.codeguru.com/vb
'

'
' API Declarations
'
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long

'
' API Types
'
' RECT is used to get the size of the panel we're inserting into
'
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'
' API Messages
'
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
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
End Function

Public Sub Main()
   InitCommonControlsVB
frmBrowser.Show
End Sub


