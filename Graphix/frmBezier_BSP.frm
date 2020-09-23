VERSION 5.00
Begin VB.Form frmBezier_BSP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Bazier curve versus BSP curve , Comparison Window"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   Icon            =   "frmBezier_BSP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   113
      ScaleHeight     =   705
      ScaleWidth      =   9315
      TabIndex        =   4
      Top             =   45
      Width           =   9345
      Begin VB.CommandButton cmdDraw 
         Caption         =   "Draw"
         Height          =   420
         Left            =   150
         TabIndex        =   8
         Top             =   120
         Width           =   1140
      End
      Begin VB.CommandButton cmdclr 
         Caption         =   "Clear"
         Height          =   420
         Left            =   1455
         TabIndex        =   7
         Top             =   120
         Width           =   1140
      End
      Begin VB.HScrollBar scrlSmooth 
         Height          =   240
         Left            =   3075
         Max             =   100
         Min             =   2
         TabIndex        =   6
         Top             =   345
         Value           =   100
         Width           =   2580
      End
      Begin VB.HScrollBar scrlNPts 
         Height          =   240
         Left            =   6315
         Max             =   14
         Min             =   2
         TabIndex        =   5
         Top             =   345
         Value           =   10
         Width           =   2580
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Smoothing value :"
         Height          =   195
         Left            =   3210
         TabIndex        =   12
         Top             =   75
         Width           =   1275
      End
      Begin VB.Label lblSmooth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4605
         TabIndex        =   11
         Top             =   75
         Width           =   465
      End
      Begin VB.Label lblNPts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8295
         TabIndex        =   10
         Top             =   75
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Number of controll points :"
         Height          =   195
         Left            =   6360
         TabIndex        =   9
         Top             =   75
         Width           =   1845
      End
   End
   Begin VB.PictureBox PP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6660
      Left            =   90
      ScaleHeight     =   440
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   621
      TabIndex        =   0
      Top             =   1230
      Width           =   9375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BSP Curve"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   1335
      TabIndex        =   3
      Top             =   870
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Convex hull"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2385
      TabIndex        =   2
      Top             =   870
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bazier Curve"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   135
      TabIndex        =   1
      Top             =   870
      Width           =   1155
   End
End
Attribute VB_Name = "frmBezier_BSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pnt As PointArray

Private Sub cmdClr_Click()
    PP.Cls
    PP.Refresh
End Sub

Private Sub GenPts(N As Integer)
    Dim i As Integer
    For i = 0 To N - 1
        pnt.point(i).x = 10 + Rnd() * (PP.ScaleWidth - 20)
        pnt.point(i).y = 10 + Rnd() * (PP.ScaleHeight - 20)
    Next i
    pnt.count = N
    DrawPnts
End Sub

Public Sub DrawPnts()
    Dim i As Integer
    PP.Cls
    For i = 0 To pnt.count - 1
        PP.Circle (pnt.point(i).x, pnt.point(i).y), 2, RGB(255, 255, 0)
        PP.CurrentX = pnt.point(i).x
        PP.CurrentY = pnt.point(i).y
        PP.Print i + 1
        If (i > 0) Then PP.Line (pnt.point(i - 1).x, pnt.point(i - 1).y)-(pnt.point(i).x, pnt.point(i).y), RGB(80, 80, 80)
    Next i
    DrawBSP PP.hdc, pnt, RGB(255, 0, 0)
    DrawBazier PP.hdc, pnt, RGB(0, 0, 255)
End Sub

Private Sub cmdDraw_Click()
    cmdClr_Click
    GenPts (CInt(lblNPts.Caption))
End Sub

Private Sub Form_Activate()
    PP.Refresh
    scrlSmooth.Value = Segments
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.SetFocus
End Sub

Private Sub Form_Load()
    cmdDraw_Click
End Sub

Private Sub scrlNPts_Change()
    lblNPts.Caption = scrlNPts.Value
End Sub

Private Sub scrlNPts_Scroll()
    lblNPts.Caption = scrlNPts.Value
End Sub

Private Sub scrlSmooth_Change()
    lblSmooth.Caption = scrlSmooth.Value
    SetSmooth (CInt(lblSmooth.Caption))
    cmdClr_Click
    DrawPnts
End Sub

Private Sub scrlSmooth_Scroll()
    lblSmooth.Caption = scrlSmooth.Value
End Sub


