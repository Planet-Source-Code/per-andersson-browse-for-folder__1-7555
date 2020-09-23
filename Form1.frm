VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Browse For Folder"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8205
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Browse For Folder"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "www.FireStormEntertainment.Cjb.Net"
      Height          =   195
      Left            =   5520
      TabIndex        =   2
      Top             =   0
      Width           =   2610
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   6480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim sPath As String

    sPath = Browse(Me.hWnd, "Browse For Folder")
    
    Label1.Caption = sPath
    
End Sub



