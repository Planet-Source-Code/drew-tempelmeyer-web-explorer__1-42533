VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "Web Explorer"
   ClientHeight    =   6660
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12435
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   240
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":36AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4388
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5062
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6A16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox add 
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":76F0
   End
   Begin VB.TextBox srccode 
      Height          =   4575
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   9375
   End
   Begin InetCtlsObjects.Inet src 
      Left            =   1680
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   5100
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   8996
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   2400
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6300
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "Done"
            TextSave        =   "Done"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Text            =   "Online"
            TextSave        =   "Online"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2160
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":777B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7CD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8233
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":878F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8CEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9247
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":97A3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9CFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A25B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A7B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AD13
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B26F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B7CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BD27
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   780
      Left            =   120
      TabIndex        =   2
      Top             =   50
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   1376
      ButtonWidth     =   1376
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Description     =   "Back"
            Object.Tag             =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Description     =   "Forward"
            Object.Tag             =   "Forward"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Description     =   "Stop"
            Object.Tag             =   "Stop"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Description     =   "Refresh"
            Object.Tag             =   "Refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home"
            Description     =   "Home"
            Object.Tag             =   "Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Favorites"
            Description     =   "Favorites"
            Object.Tag             =   "Favorites"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Description     =   "Search"
            Object.Tag             =   "Search"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   1230
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   2170
      BandCount       =   2
      _CBWidth        =   12435
      _CBHeight       =   1230
      _Version        =   "6.0.8169"
      MinHeight1      =   825
      Width1          =   2880
      NewRow1         =   0   'False
      Child2          =   "address"
      MinHeight2      =   315
      Width2          =   8895
      FixedBackground2=   0   'False
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Begin VB.ComboBox address 
         Height          =   315
         Left            =   165
         TabIndex        =   3
         Top             =   885
         Width           =   12180
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu newwindow 
         Caption         =   "&New Window"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu savepage 
         Caption         =   "&Save Page As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu pagesetup 
         Caption         =   "&Page Setup..."
      End
      Begin VB.Menu print 
         Caption         =   "P&rint..."
         Shortcut        =   ^P
      End
      Begin VB.Menu printpreview 
         Caption         =   "Pr&int Preview..."
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu properties 
         Caption         =   "&Properties..."
      End
      Begin VB.Menu offline 
         Caption         =   "Work Offline"
      End
      Begin VB.Menu close 
         Caption         =   "Close Window"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "&Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu selectall 
         Caption         =   "&Select All..."
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu source 
         Caption         =   "So&urce"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuFave 
      Caption         =   "F&avorites"
      Begin VB.Menu addtofaves 
         Caption         =   "Add To Favorites"
      End
      Begin VB.Menu viewfaves 
         Caption         =   "Show Favorites"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu about 
         Caption         =   "A&bout Web Explorer"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AllowPopups As Boolean
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub address_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        address.AddItem address.Text
        wb.Navigate (address.Text)
    End If
End Sub

Private Sub addtofaves_Click()
On Error Resume Next
add.Text = wb.LocationURL
add.SaveFile App.Path & "\Favorites\" & wb.LocationName
MsgBox wb.LocationName & " was successfully added to your favorites list.", vbInformation, "Favorite Item Added"
End Sub

Private Sub chameleonButton1_Click()
address.AddItem address.Text
wb.Navigate address.Text
End Sub

Private Sub close_Click()
End
End Sub

Private Sub copy_Click()
wb.ExecWB OLECMDID_COPY, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub cut_Click()
wb.ExecWB OLECMDID_CUT, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub find_Click()
wb.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub editfile_Click()

End Sub

Private Sub Form_Initialize()
Dim X As Long
X = InitCommonControls
End Sub

Private Sub Form_Load()
On Error Resume Next
MkDir App.Path & "\Favorites"
ShowProgressInStatusBar True
AllowPopups = False
wb.GoHome
End Sub

Private Sub Form_Resize()
On Error Resume Next
wb.Height = Me.ScaleHeight - wb.Top - 375
wb.Width = Me.ScaleWidth
wb.Top = 1320
wb.Left = 0
End Sub

Private Sub fullscreen_Click()

End Sub

Private Sub newwindow_Click()
Dim nf As New frmBrowser
nf.Show
End Sub

Private Sub offline_Click()
If offline.Checked = True Then
wb.offline = False
offline.Checked = False
StatusBar1.Panels.Item(3).Text = "Online"
Else
offline.Checked = True
wb.offline = True
StatusBar1.Panels.Item(3).Text = "Offline"
End If
End Sub

Private Sub open_Click()
wb.ExecWB OLECMDID_OPEN, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub pagesetup_Click()
    On Error Resume Next
    wb.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub paste_Click()
wb.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub print_Click()
    On Error Resume Next
    wb.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub printpreview_Click()
wb.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub properties_Click()
wb.ExecWB OLECMDID_PROPERTIES, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub savepage_Click()
    On Error Resume Next
    wb.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub selectall_Click()
wb.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End Sub

Private Sub source_Click()
Dim code As String
On Error Resume Next
If source.Checked = True Then
srccode.Visible = False
source.Checked = False
Else
srccode.Visible = True
code = src.OpenURL(address.Text)
source.Checked = True
srccode.Text = code
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
With wb
Select Case Button.Index
        Case "1"
            .GoBack
        Case "2"
            .GoForward
        Case "3"
            .Stop
        Case "4"
            .Refresh
        Case "5"
            .GoHome
        Case "7"
            frmFaves.Show
        Case "8"
            .GoSearch
End Select
End With
End Sub

Private Sub viewfaves_Click()
frmFaves.Show
End Sub

Private Sub wb_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
address.Text = URL
add.Text = URL
End Sub

Private Sub wb_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
address.Text = URL
add.Text = URL
End Sub

Private Sub wb_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
    ProgressBar1.Max = ProgressMax                          ' Progress bar above the browser
    ProgressBar1.Value = Progress
    ProgressBar1.Refresh
End Sub

Private Sub wb_StatusTextChange(ByVal Text As String)
If wb.offline = True Then
offline.Checked = True
StatusBar1.Panels.Item(3).Text = "Offline"
Else
offline.Checked = False
StatusBar1.Panels.Item(3).Text = "Online"
End If
StatusBar1.Panels.Item(1).Text = Text
End Sub

Private Sub wb_TitleChange(ByVal Text As String)
Me.Caption = Text & " - Web Explorer"
End Sub

Private Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)

    Dim tRC As RECT
    
    If bShowProgressBar Then
'
' Get the size of the Panel (2) Rectangle from the status bar
' remember that Indexes in the API are always 0 based (well,
' nearly always) - therefore Panel(2) = Panel(1) to the api
'
'
        SendMessageAny StatusBar1.hWnd, SB_GETRECT, 1, tRC
'
' and convert it to twips....
'
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
'
' Now Reparent the ProgressBar to the statusbar
'
        With ProgressBar1
            SetParent .hWnd, StatusBar1.hWnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 0
        End With
        
    Else
'
' Reparent the progress bar back to the form and hide it
'
        SetParent ProgressBar1.hWnd, Me.hWnd
        ProgressBar1.Visible = False
    End If
    
End Sub
