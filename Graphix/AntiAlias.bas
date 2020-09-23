Attribute VB_Name = "Module1"
Option Explicit

'===================================================
'*************************************************
'
'             Anti-alias GRAPHIX LIBRARY
'            ----------------------------
'    * Some Anti-aliasing routines in pure VB!
'
'     Code by,    Hari Krishnan G.
'                 harietr@yahoo.com
'
'         Anybody can use this code, completely
'   or partially, anywhere they want, personally
'   or commertially. But don't forget to mail me.
'
'*************************************************
'===================================================

Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public m_Xpos As Long, m_Ypos As Long


' Anti-aliased line
Public Function LineAA(hdc As Long, x1 As Long, y1 As Long, x2 As Long, _
y2 As Long, AColor As Long)
    Dim deltax As Integer, deltay As Integer, loopc As Integer
    Dim start As Integer, finish As Integer
    Dim dx As Single, dy As Single, dydx As Single
    Dim LR As Byte, LG As Byte, LB As Byte
    Dim pt As POINTAPI
    Dim hpen As Long
    
    deltax = Abs(x2 - x1) ' Calculate deltax and deltay for initialisation
    deltay = Abs(y2 - y1)
    If (deltax <> 0) And (deltay <> 0) Then  ' it is not a horizontal _
    or a vertical line
        LR = (AColor And &HFF&)
        LG = (AColor And &HFF00&) / &H100&
        LB = (AColor And &HFF0000) / &H10000
        If deltax > deltay Then  ' horizontal or vertical
            If y2 > y1 Then ' determine rise and run
                dydx = -(deltay / deltax)
            Else
                dydx = deltay / deltax
            End If
            If x2 < x1 Then
                start = x2 ' right to left
                finish = x1
                dy = y2
            Else
                start = x1 ' left to right
                finish = x2
                dy = y1
                dydx = -dydx ' inverse slope
            End If
            For loopc = start To finish
                AlphaBlendPixel hdc, loopc, CInt(dy - 0.5), LR, LG, LB, _
                1 - FracPart(dy)
                AlphaBlendPixel hdc, loopc, CInt(dy - 0.5) + 1, LR, LG, _
                LB, FracPart(dy)
                dy = dy + dydx ' next point
            Next loopc
        Else
            If x2 > x1 Then ' determine rise and run
                dydx = -(deltax / deltay)
            Else
                dydx = deltax / deltay
            End If
            If y2 < y1 Then
                start = y2 ' right to left
                finish = y1
                dx = x2
            Else
                start = y1 ' left to right
                finish = y2
                dx = x1
                dydx = -dydx ' inverse slope
            End If
            For loopc = start To finish
                AlphaBlendPixel hdc, CInt(dx - 0.5), loopc, LR, LG, LB, _
                1 - FracPart(dx)
                AlphaBlendPixel hdc, CInt(dx - 0.5) + 1, loopc, LR, LG, _
                LB, FracPart(dx)
                dx = dx + dydx ' next point
            Next loopc
        End If
    Else
        hpen = CreatePen(0, 1, AColor)
        SelectObject hdc, hpen
        MoveToEx hdc, x1, y1, pt
        LineTo hdc, x2, y2
        DeleteObject hpen
    End If
End Function

' blend a pixel with the current colour and a specified colour
Public Function AlphaBlendPixel(ByVal hdc As Long, ByVal x As Integer, _
ByVal y As Integer, ByVal R As Byte, ByVal g As Byte, ByVal b As Byte, _
ByVal ARatio As Double)
    Dim LMinusRatio As Double
    Dim nr As Byte, ng As Byte, nb As Byte
    Dim dstc As Long, dr As Byte, dg As Byte, db As Byte

    LMinusRatio = 1 - ARatio
    dstc = GetPixel(hdc, x, y)
    dr = (dstc And &HFF&)
    dg = (dstc And &HFF00&) / &H100&
    db = (dstc And &HFF0000) / &H10000
    
    nb = Round(b * ARatio + db * LMinusRatio)
    ng = Round(g * ARatio + dg * LMinusRatio)
    nr = Round(R * ARatio + dr * LMinusRatio)
    
        
    SetPixel hdc, x, y, RGB(nr, ng, nb)
End Function

' Returns the fractional part of a double
Public Function FracPart(ByVal a As Double) As Double
    Dim b As Double
    b = CLng(a - 0.5)
    FracPart = a - b
End Function


' Emulation
Public Function MoveToAA(ByVal x As Long, ByVal y As Long)
    m_Xpos = x
    m_Ypos = y
End Function

Public Sub LineToAA(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal clr As Long)
    LineAA hdc, x, y, m_Xpos, m_Ypos, clr
    m_Xpos = x
    m_Ypos = y
End Sub
