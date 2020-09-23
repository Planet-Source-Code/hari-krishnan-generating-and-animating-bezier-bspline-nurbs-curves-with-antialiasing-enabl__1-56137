VERSION 5.00
Begin VB.Form frmToolTip 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   58
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   252
   ShowInTaskbar   =   0   'False
   Begin VB.Label Tip 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1305
      TabIndex        =   0
      Top             =   225
      Width           =   1590
   End
   Begin VB.Shape Brdr 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   645
      Left            =   90
      Top             =   45
      Width           =   870
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Me.Hide
End Sub

Private Sub Form_Load()
    Brdr.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    Tip.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - 20
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Hide
End Sub
