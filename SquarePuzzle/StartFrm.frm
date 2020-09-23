VERSION 5.00
Begin VB.Form StartFrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "StartFrm.frx":0000
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   45
      ScaleHeight     =   330
      ScaleWidth      =   6000
      TabIndex        =   2
      Top             =   3825
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Nadianne"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   225
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "By Swertvaegher Stephan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   630
      TabIndex        =   0
      Top             =   4185
      Width           =   4605
   End
End
Attribute VB_Name = "StartFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
StartBit = 1
StartFrm.Enabled = False
StartFrm.Hide
PForm2.Show
End Sub

Private Sub Form_Activate()
For xx = 1 To 19
Line ((xx * 20) - 1, 0)-((xx * 20) - 1, 249), &H888888
Line (xx * 20, 0)-(xx * 20, 249), &HCCCCCC
Next xx
For xx = 1 To 9
Line (0, (xx * 25) - 1)-(399, (xx * 25) - 1), &H888888
Line (0, (xx * 25))-(399, (xx * 25)), &HCCCCCC
Next xx
Line (0, 0)-(399, 0), &H888888
Line (0, 0)-(0, 249), &H888888
Line (399, 1)-(399, 249), &HCCCCCC
Line (1, 249)-(399, 249), &HCCCCCC
Picture1.Print "                Square Puzzle V1.0"
Picture1.SetFocus
End Sub

