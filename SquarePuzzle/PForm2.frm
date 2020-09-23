VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form PForm2 
   Caption         =   "Form1"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Pieces"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   2115
      TabIndex        =   4
      Top             =   3060
      Width           =   1995
      Begin VB.OptionButton Option1 
         Caption         =   "35 (7  X 5)"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         Caption         =   "48 (8 X 6)"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   540
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         Caption         =   "80 (10 X 8)"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   810
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         Caption         =   "120 (12  X 10)"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   1080
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Left            =   2115
      TabIndex        =   3
      Top             =   0
      Width           =   4065
      Begin VB.Frame Frame2 
         Height          =   1230
         Left            =   0
         TabIndex        =   8
         Top             =   855
         Visible         =   0   'False
         Width           =   4065
         Begin ComctlLib.ProgressBar PBar1 
            Height          =   150
            Left            =   135
            TabIndex        =   10
            Top             =   900
            Width           =   3750
            _ExtentX        =   6615
            _ExtentY        =   265
            _Version        =   327682
            Appearance      =   0
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Making the puzzle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   135
            TabIndex        =   9
            Top             =   225
            Width           =   3795
         End
      End
      Begin VB.Image Image1 
         Height          =   2500
         Left            =   270
         Stretch         =   -1  'True
         Top             =   315
         Width           =   3500
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4905
      TabIndex        =   2
      Top             =   3825
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4905
      TabIndex        =   1
      Top             =   3285
      Width           =   1050
   End
   Begin VB.FileListBox File1 
      Height          =   4185
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1995
   End
End
Attribute VB_Name = "PForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
If clickbit = 0 Then Exit Sub
Frame2.Visible = True
DoEvents
'---------------------
frmMain.Pic1(0).Visible = False
If Pieces > 2 Then
PBar1.Min = 1: PBar1.Max = Pieces - 1
For xx = 1 To Pieces - 1
Unload frmMain.Pic1(xx)
PBar1.Value = xx
Next xx
End If
PBar1.Min = 0: PBar1.Max = 119
For xx = 0 To 119
PForm3.Pic1(xx).Visible = False
PBar1.Value = xx
Next xx
'----------------------
frmMain.PicClip1.Picture = LoadPicture(PicNam)
PicWidth = frmMain.PicClip1.Width
PicHeight = frmMain.PicClip1.Height
If Option1(0).Value = True Then
Pieces = 48
frmMain.PicClip1.Cols = 8: frmMain.PicClip1.Rows = 6
frmMain.Grid1.Cols = 8: frmMain.Grid1.Rows = 6
GrCol = 8: GrRow = 6
End If
If Option1(1).Value = True Then
Pieces = 80
frmMain.PicClip1.Cols = 10: frmMain.PicClip1.Rows = 8
frmMain.Grid1.Cols = 10: frmMain.Grid1.Rows = 8
GrCol = 10: GrRow = 8
End If
If Option1(2).Value = True Then
Pieces = 120
frmMain.PicClip1.Cols = 12: frmMain.PicClip1.Rows = 10
frmMain.Grid1.Cols = 12: frmMain.Grid1.Rows = 10
GrCol = 12: GrRow = 10
End If
If Option1(3).Value = True Then
Pieces = 35
frmMain.PicClip1.Cols = 7: frmMain.PicClip1.Rows = 5
frmMain.Grid1.Cols = 7: frmMain.Grid1.Rows = 5
GrCol = 7: GrRow = 5
End If
For t = 1 To Pieces - 1
PBar1.Min = 1: PBar1.Max = Pieces - 1
Load frmMain.Pic1(t)
Randomize
frmMain.Pic1(t).Left = Int((Rnd * 700) + 25)
Randomize
frmMain.Pic1(t).Top = Int((Rnd * 450) + 50)
frmMain.Pic1(t).ZOrder 0
PBar1.Value = t
Next t
Randomize
frmMain.Pic1(0).Left = Int((Rnd * 700) + 25)
Randomize
frmMain.Pic1(0).Top = Int((Rnd * 450) + 50)
frmMain.Pic1(0).ZOrder 0

PBar1.Min = 0: PBar1.Max = Pieces - 1
For t = 0 To Pieces - 1
frmMain.Pic1(t).Picture = frmMain.PicClip1.GraphicCell(t)
frmMain.Pic1(t).Visible = True
PBar1.Value = t
Next t
'------------------------------
For xx = 0 To GrCol - 1 'corners x
GrCorX(xx) = (xx * frmMain.Pic1(0).Width) + 13 + (frmMain.Pic1(0).Width / 2)
Next xx
For xx = 0 To GrRow - 1
GrCorY(xx) = (xx * frmMain.Pic1(0).Height) + 51 + (frmMain.Pic1(0).Height / 2)
Next xx
'------------------------------
frmMain.Grid1.Width = PicWidth + 2
frmMain.Grid1.Height = PicHeight + 2
frmMain.Grid1.Visible = True
frmMain.Grid1.ZOrder 1
frmMain.Grid1.Top = 10
frmMain.Grid1.Left = 10
For xx = 0 To GrCol - 1
frmMain.Grid1.ColWidth(xx) = frmMain.Pic1(0).Width * 15
Next xx
For xx = 0 To GrRow - 1
frmMain.Grid1.RowHeight(xx) = frmMain.Pic1(0).Height * 15
Next xx
'-----------------
frmMain.Grid1.ForeColor = frmMain.Grid1.BackColor
t = 0
For yy = 0 To GrRow - 1
For xx = 0 To GrCol - 1
    frmMain.Grid1.Col = xx: frmMain.Grid1.Row = yy
    frmMain.Grid1.Picture = LoadPicture("")
    frmMain.Grid1.Text = t: t = t + 1

Next xx
Next yy
'-----------------
frmMain.mnushowthumb.Enabled = True
frmMain.Image1.Picture = Image1.Picture
frmMain.Image1.Width = Image1.Width / 15
frmMain.Image1.Height = Image1.Height / 15
frmMain.Pic2.Width = (Image1.Width / 15) + 4
frmMain.Pic2.Height = (Image1.Height / 15) + 20
'------------------------------
frmMain.Pic1(0).Visible = True
Frame2.Visible = False
PForm2.Hide
PForm2.Enabled = False
frmMain.Enabled = True
frmMain.Show
DoEvents
ImTel = 1: piecestel = 0
frmMain.Pic2.Visible = False
frmMain.mnushowthumb.Enabled = True
End Sub

Private Sub Command2_Click()
PForm2.Enabled = False
frmMain.Enabled = True
frmMain.Show
End Sub

Private Sub File1_Click()
PicNam = App.Path + "\PuzzlePics\" + File1.List(File1.ListIndex)
Image1.Picture = LoadPicture(PicNam)
clickbit = 1
End Sub

Private Sub Form_Activate()
'clickbit = 0
End Sub

Private Sub Form_Load()
File1.Path = App.Path + "\PuzzlePics"
Option1(1).Value = True
File1.ListIndex = 0
End Sub
