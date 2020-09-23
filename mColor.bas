Attribute VB_Name = "mColor"
' mColor.bas
' 2005 May 12
' redbird77@earthlink.net
' http://home.earthlink.net/~redbird77

Option Explicit

' You can remove this entire module of you do not want HLS interpolation.

Public Sub RGBtoHLS(ByVal R1 As Byte, ByVal G1 As Byte, B1 As Byte, _
                    ByRef zHLS() As Single)

Dim zMax    As Single
Dim zMin    As Single
Dim zDelta  As Single
Dim H       As Single, L As Single, S As Single
Dim R       As Single, G As Single, B As Single
    
    R = R1 / 255: G = G1 / 255: B = B1 / 255
    
    zMin = GetMinimum(R, G, B): zMax = GetMaximum(R, G, B)
    zDelta = zMax - zMin

    L = (zMax + zMin) / 2

    If zDelta = 0 Then
        S = 0: H = 0  ' really undefined
    Else
        If L <= 0.5 Then
            S = (zMax - zMin) / (zMax + zMin)
        Else
            S = (zMax - zMin) / (2 - zMax - zMin)
        End If

        If R = zMax Then
            H = (G - B) / zDelta
        ElseIf G = zMax Then
            H = 2 + (B - R) / zDelta
        ElseIf B = zMax Then
            H = 4 + (R - G) / zDelta
        End If

        H = H / 6

        If H < 0 Then H = H + 1
    End If

    zHLS(0) = H: zHLS(1) = L: zHLS(2) = S
    
End Sub

Public Function HLStoLNG(ByVal H As Single, ByVal L As Single, ByVal S As Single) As Long

Dim m1  As Single
Dim m2  As Single
Dim R   As Single, G As Single, B As Single

    H = H - Int(H)
    L = Abs((Int(L) And 1) - L + Int(L))
    S = Abs((Int(S) And 1) - S + Int(S))

    If S = 0 Then
        R = L: B = L: G = L
    Else
        If L <= 0.5 Then
            m2 = L * (1 + S)
        Else
            m2 = L + S - L * S
        End If

        m1 = 2 * L - m2

        R = V(m1, m2, H + 1 / 3)
        G = V(m1, m2, H)
        B = V(m1, m2, H - 1 / 3)
    End If
    
    HLStoLNG = RGB(R * 255, G * 255, B * 255)

End Function

Private Function V(ByVal m1 As Single, ByVal m2 As Single, ByVal H As Single) As Single
    
    If H > 1 Then H = H - 1
    If H < 0 Then H = H + 1
    
    If (6 * H < 1) Then
        V = (m1 + (m2 - m1) * H * 6)
    ElseIf (2 * H < 1) Then
        V = m2
    ElseIf (3 * H < 2) Then
        V = (m1 + (m2 - m1) * ((2 / 3) - H) * 6)
    Else
        V = m1
    End If
    
End Function

Private Function GetMinimum(ParamArray vVals() As Variant) As Single

Dim i    As Integer
    
    GetMinimum = vVals(0)
    
    For i = 1 To UBound(vVals)
        If vVals(i) < GetMinimum Then GetMinimum = vVals(i)
    Next
    
End Function

Private Function GetMaximum(ParamArray vVals() As Variant) As Single

Dim i   As Integer
    
    GetMaximum = vVals(0)
    
    For i = 1 To UBound(vVals)
        If vVals(i) > GetMaximum Then GetMaximum = vVals(i)
    Next

End Function
