VERSION 5.00
Begin VB.Form FBrowserDialogBox 
   Caption         =   "Select directory:"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   Icon            =   "FBrowserDialogBox.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.DriveListBox drvDBLocation 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
   Begin VB.FileListBox flbMDBFile 
      Height          =   480
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.DirListBox dirDBLocation 
      Height          =   4140
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   4920
      Width           =   975
   End
End
Attribute VB_Name = "FBrowserDialogBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    dirDBLocation.Path = App.Path
End Sub

