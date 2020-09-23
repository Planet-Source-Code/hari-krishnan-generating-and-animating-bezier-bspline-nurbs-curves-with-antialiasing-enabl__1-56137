VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Main Control Form  [Graphix Library for Visual Basic 6.0]"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   2295
   Begin VB.CheckBox chkAA 
      Caption         =   "Enable Anti-alias"
      Height          =   240
      Left            =   405
      TabIndex        =   12
      Top             =   5130
      Value           =   1  'Checked
      Width           =   1500
   End
   Begin VB.CommandButton cmdBasic3D 
      Caption         =   "Basic 3D"
      Height          =   420
      Left            =   225
      TabIndex        =   6
      Top             =   4185
      Width           =   1815
   End
   Begin VB.CommandButton cmdFractal 
      Caption         =   "Fractals Demo"
      Height          =   420
      Left            =   225
      TabIndex        =   5
      Top             =   3645
      Width           =   1815
   End
   Begin VB.CommandButton cmd2DTransformation 
      Caption         =   "2D Transformations"
      Height          =   420
      Left            =   225
      TabIndex        =   4
      Top             =   3105
      Width           =   1815
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit Demo"
      Height          =   420
      Left            =   225
      TabIndex        =   3
      Top             =   6300
      Width           =   1815
   End
   Begin VB.CommandButton cmdBaz_BSP 
      Caption         =   """Bazier""  VS  ""BSP"""
      Height          =   420
      Left            =   225
      TabIndex        =   2
      Top             =   2565
      Width           =   1815
   End
   Begin VB.CommandButton cmdBSP 
      Caption         =   "BSP Curve Demo"
      Height          =   420
      Left            =   225
      TabIndex        =   1
      Top             =   2025
      Width           =   1815
   End
   Begin VB.CommandButton cmdBazier 
      Caption         =   "Bazier Demo"
      Height          =   420
      Left            =   225
      TabIndex        =   0
      Top             =   1485
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hari Krishnan G."
      Height          =   195
      Left            =   945
      TabIndex        =   11
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lib"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   885
      TabIndex        =   9
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GraphiX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   270
      TabIndex        =   7
      Top             =   375
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   5505
      Left            =   90
      Top             =   1350
      Width           =   2130
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   1140
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   90
      Width           =   2130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GraphiX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   510
      Left            =   285
      TabIndex        =   8
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lib"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   360
      Left            =   915
      TabIndex        =   10
      Top             =   150
      Width           =   465
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()


Private Sub chkAA_Click()
    If frmBazier.Visible = True Then frmBazier.DrawPnts
    If frmBsp.Visible = True Then frmBsp.DrawPnts
    If frmBezier_BSP.Visible = True Then frmBezier_BSP.DrawPnts
End Sub

Private Sub cmd2DTransformation_Click()
    frm2DTransform.Show vbModeless, Me
End Sub

Private Sub cmdBasic3D_Click()
    frmBasic3D.Show vbModeless, Me
End Sub

Private Sub cmdBaz_BSP_Click()
    frmBezier_BSP.Show vbModeless, Me
End Sub

Private Sub cmdBazier_Click()
    frmBazier.Show vbModeless, Me
End Sub

Private Sub cmdBSP_Click()
    frmBsp.Show vbModeless, Me
End Sub

Private Sub cmdFractal_Click()
    frmFractal.Show vbModeless, Me
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    InitialiseTransform 249, 18
    Randomize
    SetSmooth 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
