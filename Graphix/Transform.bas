Attribute VB_Name = "Transform"
Option Explicit

'===================================================
'*************************************************
'
'               Transformation LIBRARY
'            ----------------------------
'    * Some graphic transformation routines in pure VB!
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
'

Public Type Point3d
    x As Long
    y As Long
    z As Long
End Type

Private Const RoundOFF = 1E-20
Private Const ApproxInfinity = 1E+30
Private Const Piby180 = 0.017453293

Public T2d(3, 2) As Double, T3d(4, 3) As Double
Public XCentre As Double, YCentre As Double
Public Angl As Double, Tilt As Double, I1  As Double, J1 As Double
Private cosA As Double, cosB As Double, sinA As Double, sinB As Double
Private cosAcosB As Double, sinAsinB As Double, cosAsinB As Double, sinAcosB As Double
Public MX As Double, MY As Double, MZ As Double, DS As Double
Private cosTheta As Double, cosAlpha As Double
Public PerspectiveFlag As Boolean

Public Function InitialiseTransform(ang As Double, tlt As Double)
    NewTransform2d
    NewTransform3d
    Angl = ang
    Tilt = tlt
    cosA = Cos(Angl * Piby180)
    sinA = Sin(Angl * Piby180)
    cosB = Cos(Tilt * Piby180)
    sinB = Sin(Tilt * Piby180)
    cosAcosB = cosA * cosB
    sinAsinB = sinA * sinB
    cosAsinB = cosA * sinB
    sinAcosB = sinA * cosB
    PerspectiveFlag = True
    MX = 0
    MY = 0
    MZ = 500
    DS = 500
End Function

Public Function InitialisePerspective(Flag As Boolean, x As Double, y As Double, z As Double, m As Double)
    PerspectiveFlag = Flag
    MX = x
    MY = y
    MZ = z
    DS = m
End Function

Public Function NewTransform2d()
    Dim i, j
    For i = 1 To 3
        For j = 1 To 2
            If i = j Then
                T2d(i, j) = 1
            Else
                T2d(i, j) = 0
            End If
        Next j
    Next i
End Function

Public Function Scale2d(sx, sy)
    Dim i
    For i = 1 To 3
        T2d(i, 1) = T2d(i, 1) * sx
        T2d(i, 2) = T2d(i, 2) * sy
    Next i
End Function

Public Function Translate2d(tx, ty)
    T2d(3, 1) = T2d(3, 1) + tx
    T2d(3, 2) = T2d(3, 2) + ty
End Function

Public Function Rotate2d(a As Double)
    Dim S As Double, c As Double, i As Single, temp
    a = a * (3.141592654 / 180#)
    c = Cos(a)
    S = Sin(a)
    For i = 1 To 3
        temp = T2d(i, 1) * c - T2d(i, 2) * S
        T2d(i, 2) = T2d(i, 1) * S + T2d(i, 2) * c
        T2d(i, 1) = temp
    Next
End Function

Public Function Rotate2dXY(a As Double, x, y)
    Dim S As Double, c As Double, i As Single, temp
    Translate2d -x, -y
    a = a * (3.141592654 / 180#)
    c = Cos(a)
    S = Sin(a)
    For i = 1 To 3
        temp = T2d(i, 1) * c - T2d(i, 2) * S
        T2d(i, 2) = T2d(i, 1) * S + T2d(i, 2) * c
        T2d(i, 1) = temp
    Next
    Translate2d x, y
End Function

Public Function Transform2d(ByVal x, ByVal y, ByRef xt, ByRef yt)
    xt = x * T2d(1, 1) + y * T2d(2, 1) + T2d(3, 1)
    yt = x * T2d(1, 2) + y * T2d(2, 2) + T2d(3, 2)
End Function



Public Function NewTransform3d()
    Dim i, j
    For i = 1 To 4
        For j = 1 To 3
            T3d(i, j) = 0
        Next j
        If i <> 4 Then T3d(i, i) = 1
    Next i
End Function

Public Function Translate3d(ByVal tx As Long, ty As Long, tz As Long)
    T3d(4, 1) = T3d(4, 1) + tx
    T3d(4, 2) = T3d(4, 2) + ty
    T3d(4, 3) = T3d(4, 3) + tz
End Function

Public Function Rotate3dX(ByVal theta As Double)
    Dim S As Double, c As Double, i As Single, tmp
    theta = theta * (3.141592654 / 180#)
    S = Sin(theta)
    c = Cos(theta)
    For i = 1 To 4
        tmp = T3d(i, 2) * c - T3d(i, 3) * S
        T3d(i, 3) = T3d(i, 2) * S + T3d(i, 3) * c
        T3d(i, 2) = tmp
    Next i
End Function

Public Function Rotate3dY(ByVal theta As Double)
    Dim S As Double, c As Double, i As Single, tmp
    theta = theta * (3.141592654 / 180#)
    S = Sin(theta)
    c = Cos(theta)
    For i = 1 To 4
        tmp = T3d(i, 1) * c + T3d(i, 3) * S
        T3d(i, 3) = -T3d(i, 1) * S + T3d(i, 3) * c
        T3d(i, 1) = tmp
    Next i
End Function

Public Function Rotate3dZ(ByVal theta As Double)
    Dim S As Double, c As Double, i As Single, tmp
    theta = theta * (3.141592654 / 180#)
    S = Sin(theta)
    c = Cos(theta)
    For i = 1 To 4
        tmp = T3d(i, 1) * c - T3d(i, 2) * S
        T3d(i, 2) = T3d(i, 1) * S + T3d(i, 2) * c
        T3d(i, 1) = tmp
    Next i
End Function

Public Function Map3DCoordinates(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByRef Xp, ByRef Yp)
    Dim xt As Double, yt As Double, zt As Double
    xt = MX + x * cosA - y * sinA
    yt = MY + x * sinAsinB + y * cosAsinB + z * cosB
    If PerspectiveFlag = True Then
        zt = MZ + x * sinAcosB + y * cosAcosB - z * sinB
        Xp = XCentre + CInt((DS * xt / zt) + 0.5)
        Yp = YCentre - CInt((DS * yt / zt) + 0.5)
    Else
        Xp = XCentre + CInt(xt + 0.5)
        Yp = YCentre + CInt(yt + 0.5)
    End If
End Function

'Public Function PerspectiveTransform(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByRef xt, ByRef yt)
'        Dim d As Double
'        d = ZCentre - z
'        xt = (x * ZCentre - XCentre * z) / d
'        yt = (y * ZCentre - YCentre * z) / d
'End Function
