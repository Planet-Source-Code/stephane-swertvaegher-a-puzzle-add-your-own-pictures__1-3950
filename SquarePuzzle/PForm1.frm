VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Square Puzzle V1.0"
   ClientHeight    =   8340
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "PForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin PicClip.PictureClip PictureClip1 
      Left            =   5895
      Top             =   2070
      _ExtentX        =   106
      _ExtentY        =   106
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7965
      Top             =   3420
   End
   Begin VB.PictureBox Pic2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Height          =   2250
      Left            =   8100
      ScaleHeight     =   146
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   244
      TabIndex        =   2
      Top             =   5580
      Visible         =   0   'False
      Width           =   3720
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3375
         TabIndex        =   4
         Top             =   45
         Width           =   240
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "       Thumbnail"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3660
      End
      Begin VB.Image Image1 
         Height          =   1140
         Left            =   0
         Stretch         =   -1  'True
         Top             =   270
         Width           =   1590
      End
   End
   Begin PicClip.PictureClip PicClip1 
      Left            =   4590
      Top             =   5805
      _ExtentX        =   2566
      _ExtentY        =   1455
      _Version        =   393216
      Rows            =   5
      Cols            =   10
   End
   Begin VB.PictureBox Pic1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   0
      Left            =   7560
      ScaleHeight     =   420
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   5130
      Width           =   435
   End
   Begin VB.Timer timMouse 
      Interval        =   1
      Left            =   8010
      Top             =   4140
   End
   Begin MSGrid.Grid Grid1 
      Height          =   5460
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   6540
      _Version        =   65536
      _ExtentX        =   11536
      _ExtentY        =   9631
      _StockProps     =   77
      ForeColor       =   12632256
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   8
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   0
      GridLines       =   0   'False
      HighLight       =   0   'False
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6975
      Top             =   3060
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":128C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":18C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":1BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":1EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":220E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":2528
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":2842
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":2B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":2E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PForm1.frx":3190
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Begin VB.Menu mnunewpuzzle 
         Caption         =   "New Puzzle"
      End
   End
   Begin VB.Menu mnuoption 
      Caption         =   "Options"
      Begin VB.Menu mnushowscr2 
         Caption         =   "Show Screen2"
      End
      Begin VB.Menu mnushowthumb 
         Caption         =   "Show Thumbnail"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MouseDown As Boolean
Dim MouseOver As Boolean
Dim Mouse As New CMouse
Enum ButtonState
    Up = 0
    Down = 1
    Flat = 2
End Enum

Private Sub Command1_Click()
Pic2.Visible = False
mnushowthumb.Enabled = True
End Sub

Private Sub Form_Activate()
If StartBit = 0 Then
Pic1(0).Visible = False
frmMain.Enabled = False
StartFrm.Show
End If
End Sub

Private Sub Form_Load()
Pieces = 2
For t = 0 To 119
PForm3.Pic1(t).Visible = False
Next t
Grid1.Visible = False
Pic2.Visible = False
mnushowthumb.Enabled = False
Timer1.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Pic2.Left = xpos - 4 - (Pic2.Width / 2)
Pic2.Top = ypos - 36 - (Pic2.Height / 2)
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Pic2.Top < 0 Then
    Pic2.Top = 0
    Exit Sub
    End If
Pic2.Left = xpos - 4 - (Label2.Width / 2)
Pic2.Top = ypos - 36 - (Label2.Height / 2)
End If
End Sub

Private Sub mnunewpuzzle_Click()
StartBit = 1: clickbit = 1
StartFrm.Enabled = False
StartFrm.Hide
PForm2.Show
PForm2.Enabled = True

End Sub

Private Sub mnushowscr2_Click()
frmMain.Enabled = False
frmMain.Hide
PForm3.Enabled = True
PForm3.Show
End Sub

Private Sub mnushowthumb_Click()
Pic2.Visible = True
Pic2.ZOrder 0
mnushowthumb.Enabled = False
Command1.Left = Pic2.Width - 25
End Sub

Private Sub pic1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Pic1(Index).Visible = False
PForm3.Pic1(Index).Visible = True
PForm3.Pic1(Index).Picture = Pic1(Index).Picture
PForm3.Pic1(Index).Left = Pic1(Index).Left
PForm3.Pic1(Index).Top = Pic1(Index).Top
End If
End Sub

Private Sub Pic1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Pic1(Index).Left = xpos - 4 - (Pic1(Index).Width / 2)
Pic1(Index).Top = ypos - 36 - (Pic1(Index).Height / 2)
Pic1(Index).ZOrder 0
End If
End Sub


Private Sub Pic1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    For xx = 0 To GrCol - 1
    For yy = 0 To GrRow - 1
        If xpos < GrCorX(xx) + 10 And xpos > GrCorX(xx) - 10 And ypos < GrCorY(yy) + 10 And ypos > GrCorY(yy) - 10 Then
        Grid1.Col = xx: Grid1.Row = yy
        If Grid1.Text = Index Then
        Idx = Index
        Grid1.Picture = Pic1(Index).Picture
        Pic1(Index).Visible = False
        ImTel = 1: Timer1.Enabled = True
        piecestel = piecestel + 1
'        Grid1.CellSelected = False
        End If
        End If
    Next yy
    Next xx

End If

End Sub

Private Sub Timer1_Timer()
Grid1.Picture = ImageList1.ListImages(ImTel).Picture
ImTel = ImTel + 1
If ImTel = 17 Then
Grid1.Picture = Pic1(Idx).Picture
DoEvents
            If piecestel = Pieces Then
            Antw = MsgBox("You solved the puzzle !" + Chr(13) + "Try another puzzle ?", vbExclamation + vbYesNo, "Congratulations")
                If Antw = vbYes Then
                StartBit = 1: clickbit = 1
                StartFrm.Enabled = False
                StartFrm.Hide
                PForm2.Show
                PForm2.Enabled = True
                'Else
                'End
                End If
            End If
Timer1.Enabled = False
End If
End Sub

Private Sub timMouse_Timer()
Dim dummy
dummy = Mouse.WindowOver(Mouse.X, Mouse.Y)
End Sub
