VERSION 5.00
Begin VB.Form frmBazier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Bazier Drawing Window"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   90
      ScaleHeight     =   705
      ScaleWidth      =   9360
      TabIndex        =   1
      Top             =   90
      Width           =   9390
      Begin VB.HScrollBar scrlNPts 
         Height          =   240
         Left            =   6480
         Max             =   14
         Min             =   2
         TabIndex        =   5
         Top             =   360
         Value           =   10
         Width           =   2580
      End
      Begin VB.HScrollBar scrlSmooth 
         Height          =   240
         Left            =   3240
         Max             =   100
         Min             =   2
         TabIndex        =   4
         Top             =   360
         Value           =   50
         Width           =   2580
      End
      Begin VB.CommandButton cmdclr 
         Caption         =   "Clear"
         Height          =   420
         Left            =   1620
         TabIndex        =   3
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdDraw 
         Caption         =   "Draw"
         Height          =   420
         Left            =   315
         TabIndex        =   2
         Top             =   135
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Number of controll points :"
         Height          =   195
         Left            =   6525
         TabIndex        =   9
         Top             =   90
         Width           =   1845
      End
      Begin VB.Label lblNPts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8460
         TabIndex        =   8
         Top             =   90
         Width           =   465
      End
      Begin VB.Label lblSmooth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "100"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4770
         TabIndex        =   7
         Top             =   90
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Smoothing value :"
         Height          =   195
         Left            =   3375
         TabIndex        =   6
         Top             =   90
         Width           =   1275
      End
   End
   Begin VB.PictureBox PP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6540
      Left            =   90
      ScaleHeight     =   432
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   621
      TabIndex        =   0
      Top             =   900
      Width           =   9375
   End
End
Attribute VB_Name = "frmBazier"
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
    DrawBazier PP.hdc, pnt, RGB(255, 0, 0)
End Sub

Private Sub cmdDraw_Click()
    cmdClr_Click
    GenPts (CInt(lblNPts.Caption))
End Sub

Private Sub Form_Load()
    cmdDraw_Click
End Sub

Private Sub Form_Activate()
    PP.Refresh
    scrlSmooth.Value = Segments
End Sub

Private Sub Form_Unloald(Cancel As Integer)
    frmMain.SetFocus
End Sub

Private Sub scrlNPts_Change()
    lblNPts.Caption = scrlNPts.Value
End Sub

Private Sub scrlNPts_Scroll()
    lblNPts.Caption = scrlNPts.Value
    scrlNPts_Change
End Sub

Private Sub scrlSmooth_Change()
    lblSmooth.Caption = scrlSmooth.Value
    SetSmooth (CInt(lblSmooth.Caption))
    cmdClr_Click
    DrawPnts
End Sub

Private Sub scrlSmooth_Scroll()
    lblSmooth.Caption = scrlSmooth.Value
    scrlSmooth_Change
End Sub
