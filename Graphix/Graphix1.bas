Attribute VB_Name = "Graphix"
Option Explicit

'===================================================
'*************************************************
'
'               GRAPHIX LIBRARY
'            ---------------------
'    * Some complex graphics routines
'
'     Code by,    Hari Krishnan G.
'                 hari_@ yours.com
'
'         Anybody can use this code, completely
'   or partially, anywhere they want, personally
'   or commertially. But don't forget to mail me
'   when you are using it commertially ,(just to
'   notify.)
'
'*************************************************
'===================================================
'
'  List of functions
'------------------------------------------------------------------------------
'  Function     |   Explanation      |      Parameters ...
'------------------------------------------------------------------------------
' 1) DrawBazier()  |  Draws a Bazier curve.   |  hdc - The HDC parameter of the drawing control (eg: a picturebox 'Pic' then 'Pic.hdc'), p - a "PointArray" type value containing the control points, c - colour of curve
' 2) DrawBSP()     |  Draws a BSP curve.      |  hdc - The HDC parameter of the drawing control (eg: a picturebox 'Pic' then 'Pic.hdc'), p - a "PointArray" type value containing the control points, c - colour of curve
' 3) DrawFractalLine() | Draws a FractalLine between two given points | (x,y) - First point, (xx,yy) - Second point, W - Roughness of the curve, N - Depth of recursion, c- colour

'===================================================
'  Some Global Variables,Constants, and Functions
'===================================================
Public Segments As Integer

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type PointArray
    point(100) As POINTAPI
    count As Integer
End Type

Private xdummy As Long, ydummy As Long

Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Function SetSmooth(Number_of_Segments As Integer)
    If (Number_of_Segments <= 1 Or Number_of_Segments >= 100) Then Exit Function
    Segments = Number_of_Segments
End Function

Private Function ublend(i As Integer, N As Integer, u As Double) As Double
    On Error Resume Next
    Dim c As Long, j As Long, k As Long, g As Long
    Dim f As Double
    k = 1
    g = 1
    For c = N To i + 1 Step -1
        k = k * c
    Next
    For c = N - i To 2 Step -1
        g = g * c
    Next
    f = CDbl(k) / CDbl(g)
    For j = 1 To i
        f = f * u
    Next
    For j = 1 To N - i
        f = f * (1 - u)
    Next
    ublend = f
End Function

Private Function Gauss() As Double
    Dim a As Double, i
    a = 0
    For i = 1 To 8
        a = a + Rnd() - Rnd()
    Next i
    Gauss = a / 8#
End Function

'===================================================
'  Function to draw a Bazier curve.
'===================================================

Public Function DrawBazier(hdc As Long, p As PointArray, c As Long)
    Dim i As Integer, j As Integer, px As Integer, py As Integer, oldx As Integer, oldy As Integer
    Dim b As Double, u As Double, x As Double, y As Double, xt As Long, yt As Long
    Dim dummyp As POINTAPI, curP As Long, prevP As Long
    
    If (Segments <= 1 Or Segments >= 500) Then SetSmooth (100)
    
    curP = CreatePen(0, 1, c)
    prevP = SelectObject(hdc, curP)
    
    Transform2d p.point(0).x, p.point(0).y, xt, yt
    If frmMain.chkAA.Value = 1 Then
        MoveToAA xt, yt
    Else
        MoveToEx hdc, xt, yt, dummyp
    End If
    For i = 0 To Segments
        u = CDbl(i) / CDbl(Segments)
        x = 0
        y = 0
        For j = 0 To p.count - 1
            b = ublend(j, p.count - 1, u)
            x = x + p.point(j).x * b
            y = y + p.point(j).y * b
        Next j
        Transform2d x, y, xt, yt
        If frmMain.chkAA.Value = 1 Then
            LineToAA hdc, xt, yt, c
        Else
            LineTo hdc, xt, yt
        End If
        
    Next i
    
    SelectObject hdc, prevP
    DeleteObject curP
End Function

'===================================================
'  Function to draw a BSP curve.
'===================================================

Public Function DrawBSP(hdc As Long, p As PointArray, c As Long)
    Dim i As Integer, j As Integer, x As Long, y As Long
    Dim u As Double, nc1 As Double, nc2 As Double, nc3 As Double, nc4 As Double
    Dim dummyp As POINTAPI, curP As Long, prevP As Long
    Dim xt As Long, yt As Long
    Dim dp As PointArray
    
    If (Segments <= 1 Or Segments >= 500) Then SetSmooth (100)
    
    curP = CreatePen(0, 1, c)
    prevP = SelectObject(hdc, curP)

    Transform2d p.point(0).x, p.point(0).y, xt, yt
    
    If frmMain.chkAA.Value = 1 Then
        MoveToAA xt, yt
    Else
        MoveToEx hdc, xt, yt, dummyp
    End If
    
    dp.count = p.count + 1
    
    For i = dp.count To 1 Step -1
        dp.point(i).x = p.point(i - 1).x
        dp.point(i).y = p.point(i - 1).y
    Next i

    dp.point(0).x = dp.point(1).x
    dp.point(0).y = dp.point(1).y
    For j = dp.count To dp.count + 1
        dp.point(j).x = dp.point(j - 1).x
        dp.point(j).y = dp.point(j - 1).y
    Next j
    For i = 1 To j - 3
        For u = 0# To 1# Step (1# / Segments)
            nc1 = -(u * u * u) / 6# + (u * u) / 2# - u / 2# + 1# / 6#
            nc2 = (u * u * u) / 2# - (u * u) + 2# / 3#
            nc3 = (-(u * u * u) + u * u + u) / 2# + 1# / 6#
            nc4 = (u * u * u) / 6#
            
            x = nc1 * dp.point(i - 1).x + nc2 * dp.point(i).x + nc3 * dp.point(i + 1).x + nc4 * dp.point(i + 2).x
            y = nc1 * dp.point(i - 1).y + nc2 * dp.point(i).y + nc3 * dp.point(i + 1).y + nc4 * dp.point(i + 2).y
            
            Transform2d x, y, xt, yt
            If frmMain.chkAA.Value = 1 Then
                LineToAA hdc, xt, yt, c
            Else
                LineTo hdc, xt, yt
            End If
        Next u
    Next i
    
    SelectObject hdc, prevP
    DeleteObject curP
End Function

'===================================================
'  Function to draw a Random Fractal curve.
'===================================================

Public Function DrawFractalLine(hdc As Long, x, y, xx, yy, W, N, c As Long)
    Dim l As Long, pd As POINTAPI
    l = Abs(xx - x) + Abs(yy - y)
    MoveToEx hdc, x, y, pd
    FractalLineSubdivide hdc, x, y, xx, yy, l * W, N, c
End Function
Private Function FractalLineSubdivide(hdc As Long, x1, y1, x2, y2, S, N, c)
    Dim xm, ym
    Dim curP As Long, prevP As Long
    
    curP = CreatePen(0, 1, c)
    prevP = SelectObject(hdc, curP)
    
    If N = 0 Then
        LineTo hdc, x2, y2
    Else
        xm = (x1 + x2) / 2 + S * Gauss()
        ym = (y1 + y2) / 2 + S * Gauss()
        FractalLineSubdivide hdc, x1, y1, xm, ym, S / 2, N - 1, c
        FractalLineSubdivide hdc, xm, ym, x2, y2, S / 2, N - 1, c
    End If
    SelectObject hdc, prevP
    DeleteObject curP
End Function

Public Function DrawLine3D(hdc As Long, x As Double, y As Double, z As Double, xx As Double, yy As Double, zz As Double, c As Long)
    Dim xt, yt, xxt, yyt
    Dim dummyp As POINTAPI, curP As Long, prevP As Long
    
    curP = CreatePen(0, 1, c)
    prevP = SelectObject(hdc, curP)
    Map3DCoordinates x, y, z, xt, yt
    Map3DCoordinates xx, yy, zz, xxt, yyt
    If frmMain.chkAA.Value = 1 Then
        MoveToAA xt, yt
    Else
        MoveToEx hdc, xt, yt, dummyp
    End If
    If frmMain.chkAA.Value = 1 Then
        LineToAA hdc, xxt, yyt, c
    Else
        LineTo hdc, xxt, yyt
    End If
    SelectObject hdc, prevP
    DeleteObject curP
End Function

