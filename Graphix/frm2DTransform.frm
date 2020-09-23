VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frm2DTransform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "2-Dimentional Transformations."
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   Icon            =   "frm2DTransform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   60
      ScaleHeight     =   1020
      ScaleWidth      =   10260
      TabIndex        =   1
      Top             =   90
      Width           =   10290
      Begin VB.CommandButton cmdRotate 
         Caption         =   "START"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   4080
         TabIndex        =   14
         Top             =   180
         Width           =   960
      End
      Begin VB.CommandButton cmdGenC 
         Caption         =   "Generate Curve"
         Height          =   375
         Left            =   225
         TabIndex        =   8
         Top             =   90
         Width           =   1410
      End
      Begin VB.CommandButton cmdClr 
         Caption         =   "Clear Window"
         Height          =   375
         Left            =   225
         TabIndex        =   7
         Top             =   540
         Width           =   1410
      End
      Begin VB.TextBox txtSmooth 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3105
         TabIndex        =   6
         Text            =   "100"
         Top             =   135
         Width           =   435
      End
      Begin VB.TextBox txtNcpt 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3105
         TabIndex        =   4
         Text            =   "7"
         Top             =   585
         Width           =   435
      End
      Begin VB.HScrollBar scrlXC 
         Height          =   285
         Left            =   6660
         Max             =   100
         TabIndex        =   3
         Top             =   180
         Width           =   3345
      End
      Begin VB.HScrollBar scrlYC 
         Height          =   285
         Left            =   6660
         Max             =   100
         TabIndex        =   2
         Top             =   585
         Width           =   3345
      End
      Begin ComCtl2.UpDown udSmooth 
         Height          =   285
         Left            =   3533
         TabIndex        =   5
         Top             =   135
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Value           =   100
         BuddyControl    =   "txtSmooth"
         BuddyDispid     =   196613
         OrigLeft        =   3150
         OrigTop         =   315
         OrigRight       =   3390
         OrigBottom      =   870
         Increment       =   10
         Max             =   500
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown udNcpt 
         Height          =   285
         Left            =   3540
         TabIndex        =   9
         Top             =   585
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   327681
         Value           =   7
         BuddyControl    =   "txtNcpt"
         BuddyDispid     =   196614
         OrigLeft        =   3150
         OrigTop         =   315
         OrigRight       =   3390
         OrigBottom      =   870
         Max             =   14
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Smoothing value :"
         Height          =   195
         Left            =   1800
         TabIndex        =   13
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Control points : "
         Height          =   195
         Left            =   2025
         TabIndex        =   12
         Top             =   630
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         Height          =   960
         Left            =   4005
         Top             =   45
         Width           =   6180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Center of Rotation :"
         Height          =   195
         Left            =   5130
         TabIndex        =   11
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label lblCoordCnt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0,0 "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   5625
         TabIndex        =   10
         Top             =   585
         Width           =   435
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1125
      Top             =   1575
   End
   Begin VB.PictureBox PP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   7260
      Left            =   90
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   687
      TabIndex        =   0
      Top             =   1215
      Width           =   10365
   End
End
Attribute VB_Name = "frm2DTransform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pnt As PointArray, CurveGenerated As Boolean, ang As Double

Private Sub cmdClr_Click()
    PP.Cls
    CurveGenerated = False
End Sub

Private Sub cmdGenC_Click()
    PP.Cls
    NewTransform2d
    GenPts (CInt(txtNcpt.Text))
    DrawBSP PP.hdc, pnt, RGB(255, 255, 0)
    DrawBazier PP.hdc, pnt, RGB(255, 0, 0)
    PP.Circle ((scrlXC.Value), (scrlYC.Value)), 2, RGB(255, 0, 0)
    ang = 0
    CurveGenerated = True
    Timer1.Enabled = False
    cmdRotate.Value = 0
End Sub

Private Sub cmdRotate_Click()
    If cmdRotate.Caption = "START" Then
        cmdRotate.Caption = "STOP"
        Timer1.Enabled = True
    Else
        cmdRotate.Caption = "START"
        Timer1.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    PP.Refresh
    udSmooth.Value = Segments
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.SetFocus
End Sub

Private Sub GenPts(N As Integer)
    Dim i As Integer
    For i = 0 To N
        pnt.point(i).x = 10 + Rnd() * (PP.ScaleWidth - 20)
        pnt.point(i).y = 10 + Rnd() * (PP.ScaleHeight - 20)
    Next i
    pnt.count = N
    DrawPnts
End Sub

Public Sub DrawPnts()
    Dim i As Integer, xt As Long, yt As Long, x1t As Long, y1t As Long
    For i = 0 To pnt.count - 1
        Transform2d pnt.point(i).x, pnt.point(i).y, xt, yt
        If (i > 0) Then Transform2d pnt.point(i - 1).x, pnt.point(i - 1).y, x1t, y1t
        PP.Circle (xt, yt), 2, RGB(255, 255, 0)
        PP.CurrentX = xt + 1
        PP.CurrentY = yt - 1
        PP.Print i + 1
        If (i > 0) Then PP.Line (x1t, y1t)-(xt, yt), RGB(50, 50, 50)
    Next i
End Sub

Private Sub Form_Load()
    NewTransform2d
    Randomize
    SetSmooth (100)
    CurveGenerated = False
    scrlXC.Max = PP.ScaleWidth
    scrlYC.Max = PP.ScaleHeight
    scrlXC.Value = PP.ScaleWidth / 2
    scrlYC.Value = PP.ScaleHeight / 2
    lblCoordCnt.Caption = " " & scrlXC.Value & " , " & scrlYC.Value & " "
    cmdGenC_Click
End Sub

Private Sub scrlXC_Change()
    Dim te As Boolean
    te = False
    If Timer1.Enabled = True Then
        te = True
        Timer1.Enabled = False
    End If
    PP.Cls
    NewTransform2d
    lblCoordCnt.Caption = " " & scrlXC.Value & " , " & scrlYC.Value & " "
    If CurveGenerated = True Then
        DrawPnts
        DrawBazier PP.hdc, pnt, RGB(255, 0, 0)
        DrawBSP PP.hdc, pnt, RGB(255, 255, 0)
    End If
    PP.Circle (scrlXC.Value, scrlYC.Value), 2, RGB(255, 0, 0)
    Rotate2dXY ang, scrlXC.Value, scrlYC.Value
    PP.Refresh
    If te = True Then Timer1.Enabled = True
End Sub

Private Sub scrlXC_Scroll()
    scrlXC_Change
End Sub

Private Sub scrlYC_Change()
    Dim te As Boolean
    te = False
    If Timer1.Enabled = True Then
        te = True
        Timer1.Enabled = False
    End If
    PP.Cls
    NewTransform2d
    lblCoordCnt.Caption = " " & scrlXC.Value & " , " & scrlYC.Value & " "
    If CurveGenerated = True Then
        DrawPnts
        DrawBazier PP.hdc, pnt, RGB(255, 0, 0)
        DrawBSP PP.hdc, pnt, RGB(255, 255, 0)
    End If
    PP.Circle (scrlXC.Value, scrlYC.Value), 2, RGB(255, 0, 0)
    PP.Refresh
    If te = True Then Timer1.Enabled = True
End Sub

Private Sub scrlYC_Scroll()
    scrlYC_Change
End Sub

Private Sub Timer1_Timer()
    If CurveGenerated = False Then
        Timer1.Enabled = False
        cmdRotate.Value = 0
        Exit Sub
    End If
    PP.Cls
    Rotate2dXY 5#, (scrlXC.Value), (scrlYC.Value)
    ang = ang + 5#
    If ang >= 360 Then ang = ang - CInt(ang / 360) * 360
    DrawPnts
    
    DrawBazier PP.hdc, pnt, RGB(0, 0, 200)
    DrawBSP PP.hdc, pnt, RGB(0, 200, 0)
    
    PP.Circle ((scrlXC.Value), (scrlYC.Value)), 2, RGB(255, 0, 0)
    PP.Refresh
End Sub

Private Sub txtsmooth_Change()
    SetSmooth (udSmooth.Value)
End Sub
