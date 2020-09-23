VERSION 5.00
Begin VB.Form MousePos 
   Caption         =   "Mouse Coordinate Watcher"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "UnLock"
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lock"
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   480
   End
   Begin VB.TextBox getyy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox GetXX 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "  X                     Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "MousePos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = False
End Sub

Private Sub Command2_Click()
Timer1.Enabled = True
End Sub

Public Function getx() As Long
    Dim n As POINTAPI
    GetCursorPos n
    GetXX = n.x
End Function


Public Function GetY() As Long
    Dim n As POINTAPI
    GetCursorPos n
    getyy = n.y
End Function

Private Sub Timer1_Timer()
Call getx
Call GetY

End Sub
