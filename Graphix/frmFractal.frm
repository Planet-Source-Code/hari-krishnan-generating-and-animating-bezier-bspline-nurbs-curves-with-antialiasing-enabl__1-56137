VERSION 5.00
Begin VB.Form frmFractal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Fractal Curves and surfaces"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   Icon            =   "frmFractal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   105
      ScaleHeight     =   660
      ScaleWidth      =   9915
      TabIndex        =   1
      Top             =   60
      Width           =   9945
      Begin VB.CommandButton cmdGenLine 
         Caption         =   "Generate New Curve"
         Height          =   375
         Left            =   345
         TabIndex        =   10
         Top             =   135
         Width           =   1770
      End
      Begin VB.CommandButton cmdDoAgain 
         Caption         =   "Do Again"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   135
         Width           =   1770
      End
      Begin VB.CommandButton cmdClr 
         Caption         =   "Clear"
         Height          =   375
         Left            =   3975
         TabIndex        =   8
         Top             =   135
         Width           =   1095
      End
      Begin VB.TextBox txtRough 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   6585
         TabIndex        =   7
         Text            =   "1.5"
         Top             =   225
         Width           =   540
      End
      Begin VB.CommandButton cmdNegetive 
         Caption         =   "-"
         Height          =   285
         Left            =   6270
         TabIndex        =   6
         Top             =   225
         Width           =   285
      End
      Begin VB.CommandButton cmdPositive 
         Caption         =   "+"
         Height          =   285
         Left            =   7170
         TabIndex        =   5
         Top             =   225
         Width           =   285
      End
      Begin VB.CommandButton cmdP2 
         Caption         =   "+"
         Height          =   285
         Left            =   9165
         TabIndex        =   4
         Top             =   225
         Width           =   285
      End
      Begin VB.CommandButton cmdN2 
         Caption         =   "-"
         Height          =   285
         Left            =   8265
         TabIndex        =   3
         Top             =   225
         Width           =   285
      End
      Begin VB.TextBox txtDepth 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   8580
         TabIndex        =   2
         Text            =   "7"
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Roughness :"
         Height          =   195
         Left            =   5325
         TabIndex        =   12
         Top             =   270
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Depth :"
         Height          =   195
         Left            =   7695
         TabIndex        =   11
         Top             =   270
         Width           =   525
      End
   End
   Begin VB.PictureBox PP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7365
      Left            =   90
      ScaleHeight     =   487
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   663
      TabIndex        =   0
      Top             =   840
      Width           =   10005
   End
End
Attribute VB_Name = "frmFractal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x, y, xx, yy

Private Sub GenPts()
    x = 10 + (Rnd() * PP.ScaleWidth - 20)
    xx = 10 + (Rnd() * PP.ScaleWidth - 20)
    y = 10 + (Rnd() * PP.ScaleHeight - 20)
    yy = 10 + (Rnd() * PP.ScaleHeight - 20)
End Sub

Private Sub cmdClr_Click()
    PP.Cls
    cmdDoAgain.Enabled = False
End Sub

Private Sub cmdDoAgain_Click()
    Dim r, d
    r = CDbl(txtRough.Text)
    d = CInt(txtDepth.Text)
    DrawFractalLine PP.hdc, x, y, xx, yy, r, d, RGB(255, 0, 0)
    PP.Refresh
End Sub

Private Sub cmdGenLine_Click()
    Dim r, d
    PP.Cls
    cmdDoAgain.Enabled = True
    GenPts
    DrawPnts
    r = CDbl(txtRough.Text)
    d = CInt(txtDepth.Text)
    DrawFractalLine PP.hdc, x, y, xx, yy, r, d, RGB(255, 0, 0)
    PP.Refresh
End Sub

Private Sub DrawPnts()
    PP.Line (x, y)-(xx, yy), RGB(50, 50, 50)
    PP.Circle (x, y), 2, RGB(255, 255, 0)
    PP.Circle (xx, yy), 2, RGB(255, 255, 0)
    PP.CurrentX = x + 1
    PP.CurrentY = y - 1
    PP.Print "1"
    PP.CurrentX = xx + 1
    PP.CurrentY = yy - 1
    PP.Print "2"
End Sub

Private Sub cmdN2_Click()
    Dim val As Double
    val = CInt(txtDepth.Text)
    If val <= 1 Then Exit Sub
    val = val - 1
    txtDepth.Text = val
End Sub

Private Sub cmdNegetive_Click()
    Dim val As Double
    val = CDbl(txtRough.Text)
    If val <= 0.1 Then Exit Sub
    val = val - 0.1
    txtRough.Text = val
End Sub

Private Sub cmdP2_Click()
    Dim val As Double
    val = CInt(txtDepth.Text)
    If val >= 15 Then Exit Sub
    val = val + 1
    txtDepth.Text = val
End Sub

Private Sub cmdPositive_Click()
    Dim val As Double
    val = CDbl(txtRough.Text)
    If val >= 4# Then Exit Sub
    val = val + 0.1
    txtRough.Text = val
End Sub

Private Sub Form_Load()
    cmdGenLine_Click
End Sub
