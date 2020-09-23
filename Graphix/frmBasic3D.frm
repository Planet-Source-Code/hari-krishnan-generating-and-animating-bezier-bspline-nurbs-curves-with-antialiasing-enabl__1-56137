VERSION 5.00
Begin VB.Form frmBasic3D 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Basic 3-Dimentional Transformations"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   Icon            =   "frmBasic3D.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5070
      Top             =   2940
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Demo"
      Default         =   -1  'True
      Height          =   420
      Left            =   495
      TabIndex        =   1
      Top             =   90
      Width           =   1455
   End
   Begin VB.PictureBox PP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7575
      Left            =   90
      ScaleHeight     =   501
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   654
      TabIndex        =   0
      Top             =   600
      Width           =   9870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Angles:"
      Height          =   195
      Left            =   2460
      TabIndex        =   3
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   3030
      TabIndex        =   2
      Top             =   210
      Width           =   1035
   End
End
Attribute VB_Name = "frmBasic3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim angle As Double, TiltA As Double, inc As Double

Private Sub Command1_Click()
    If Timer1.Enabled = False Then
        Command1.Caption = "Stop Demo"
        Timer1.Enabled = True
    Else
        Command1.Caption = "Start Demo"
        Timer1.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Dim xt, yt, xt1, yt1
    XCentre = PP.ScaleWidth / 2
    YCentre = PP.ScaleHeight / 2
    angle = 0
    TiltA = 0
    inc = 0
    Call Timer1_Timer
'    Map3DCoordinates 0, 0, 0, xt, yt
'    Map3DCoordinates 500, 0, 0, xt1, yt1
'    PP.Line (xt, yt)-(xt1, yt1), RGB(255, 0, 0)
'    Map3DCoordinates 0, 500, 0, xt1, yt1
'    PP.Line (xt, yt)-(xt1, yt1), RGB(0, 255, 0)
'    Map3DCoordinates 0, 0, 500, xt1, yt1
'    PP.Line (xt, yt)-(xt1, yt1), RGB(0, 0, 255)
    PP.Refresh
End Sub

Private Sub Timer1_Timer()
    Dim x, y, z, x1, y1, z1, xt, yt, xt1, yt1
    angle = angle + 5#
    TiltA = TiltA + inc
    If (angle >= 360#) Then angle = 0#
    If (TiltA >= 80#) Then inc = -0.5
    If (TiltA <= 1#) Then inc = 0.5
    Label1.Caption = " " & angle & " , " & TiltA & " "
    InitialiseTransform angle, TiltA
    
    PP.Cls
    
    DrawLine3D PP.hdc, -100, -100, -100, 100, -100, -100, RGB(255, 0, 0)
    DrawLine3D PP.hdc, -100, -100, -100, -100, 100, -100, RGB(255, 0, 0)
    DrawLine3D PP.hdc, -100, 100, -100, 100, 100, -100, RGB(255, 0, 0)
    DrawLine3D PP.hdc, 100, -100, -100, 100, 100, -100, RGB(255, 0, 0)
    
    DrawLine3D PP.hdc, -100, -100, -100, -100, -100, 100, RGB(255, 0, 0)
    DrawLine3D PP.hdc, -100, -100, 100, 100, -100, 100, RGB(255, 0, 0)
    DrawLine3D PP.hdc, 100, -100, -100, 100, -100, 100, RGB(255, 0, 0)
    
    DrawLine3D PP.hdc, -100, -100, 100, -100, 100, 100, RGB(255, 0, 0)
    DrawLine3D PP.hdc, -100, 100, -100, -100, 100, 100, RGB(255, 0, 0)
    
    DrawLine3D PP.hdc, 100, 100, -100, 100, 100, 100, RGB(255, 0, 0)
    
    DrawLine3D PP.hdc, 100, -100, 100, 100, 100, 100, RGB(255, 0, 0)
    
    DrawLine3D PP.hdc, -100, 100, 100, 100, 100, 100, RGB(255, 0, 0)
    
    DrawLine3D PP.hdc, -100, -100, 100, 100, 100, 100, RGB(150, 0, 0)
    DrawLine3D PP.hdc, -100, 100, -100, 100, 100, 100, RGB(150, 0, 0)
    DrawLine3D PP.hdc, 100, -100, -100, 100, 100, 100, RGB(150, 0, 0)
    
    DrawLine3D PP.hdc, 100, -100, 100, -100, 100, 100, RGB(150, 0, 0)
    DrawLine3D PP.hdc, 100, -100, 100, 100, 100, -100, RGB(150, 0, 0)
    DrawLine3D PP.hdc, -100, 100, 100, 100, 100, -100, RGB(150, 0, 0)
    
    DrawLine3D PP.hdc, -100, -100, -100, 100, 100, -100, RGB(150, 0, 0)
    DrawLine3D PP.hdc, 100, -100, -100, -100, 100, -100, RGB(150, 0, 0)
    
    DrawLine3D PP.hdc, -100, -100, -100, -100, 100, 100, RGB(150, 0, 0)
    DrawLine3D PP.hdc, -100, 100, -100, -100, -100, 100, RGB(150, 0, 0)
    
    DrawLine3D PP.hdc, -100, -100, -100, 100, -100, 100, RGB(150, 0, 0)
    DrawLine3D PP.hdc, 100, -100, -100, -100, -100, 100, RGB(150, 0, 0)
    
    PP.Refresh
    DoEvents
End Sub
