VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFaves 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Favorites - Web Explorer"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3810
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Favorites"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin RichTextLib.RichTextBox load 
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393217
         TextRTF         =   $"frmFaves.frx":0000
      End
      Begin VB.FileListBox fave 
         Height          =   3600
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmFaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmBrowser.wb.Navigate (load.Text)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub fave_Click()
load.LoadFile fave.Path & "\" & fave.FileName
End Sub

Private Sub Form_Load()
fave.Path = App.Path & "\Favorites"
End Sub
