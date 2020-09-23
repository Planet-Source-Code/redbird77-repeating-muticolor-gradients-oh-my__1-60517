Attribute VB_Name = "mGradients"
' mGradients.bas (plural because of mutiple methods)
' 2005 May 12
' redbird77@earthlink.net
' http://home.earthlink.net/~redbird77

Option Explicit

' I usually don't like to change the MS given names (though arbitrary and
' inconsistent) of the parameters of API functions.  However in this case
' since the functions aren't even in the API viewer shipped with VB, I
' decided to come up with a simpler naming scheme.

Private Const PI    As Single = 3.14159265358979

' MESH_RECT and MESH_TRIANGLE contain the indicies of the meshes' verticies.
Private Type MESH_RECT
    v1  As Long
    v2  As Long
End Type

Private Type MESH_TRIANGLE
    v1  As Long
    v2  As Long
    V3  As Long
End Type

' Using a type, you can avoid any tricky calls to RtlMoveMemory.
Private Type USHORT
    Lo  As Byte
    Hi  As Byte
End Type

Private Type VERTEX
    X   As Long
    Y   As Long
    R   As USHORT     ' R, G, B and A = [0x0000..0xFF00]
    G   As USHORT
    B   As USHORT
    A   As USHORT
End Type

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveTo Lib "gdi32.dll" Alias "MoveToEx" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByRef lpPoint As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Dst As Any, Src As Any, ByVal lLen As Long)
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

Private Declare Function GradientFill Lib "msimg32.dll" (ByVal hDC As Long, ByRef pVertex As VERTEX, ByVal lVertexCount As Long, ByRef pMesh As Any, ByVal lMeshCount As Long, ByVal lMode As Long) As Long
' These are simply specialized versions of GradientFill that specify either
' a rectangular or triangular gradient.
Private Declare Function GradientFillRect Lib "msimg32.dll" Alias "GradientFill" (ByVal hDC As Long, ByRef pVertex As VERTEX, ByVal lVertexCount As Long, ByRef pMesh As MESH_RECT, ByVal lMeshCount As Long, ByVal lMode As Long) As Long
Private Declare Function GradientFillTri Lib "msimg32.dll" Alias "GradientFill" (ByVal hDC As Long, ByRef pVertex As VERTEX, ByVal lVertexCount As Long, ByRef pMesh As MESH_TRIANGLE, ByVal lMeshCount As Long, ByVal lMode As Long) As Long

Public Function GradientSimple(ByVal hDC As Long, _
                               ByVal X1 As Long, ByVal Y1 As Long, _
                               ByVal X2 As Long, ByVal Y2 As Long, _
                               ByVal lColor1 As Long, ByVal lColor2 As Long, _
                               ByVal lVertical As Long) As Long
                                         
' USE  : Draws a single non-repeating vertical or horizontal gradient
'        between two colors using GradientFillRect API function.

' Minimal code size and complexity.  Use when all the fancy stuff like
' repeating and mutiple colors are not needed.  The API function takes
' care of errors, returning 0 if call fails.

Dim VX(1)   As VERTEX
Dim MR      As MESH_RECT

    VX(0).X = X1
    VX(0).Y = Y1
    VX(0).R.Hi = lColor1 And &HFF&
    VX(0).G.Hi = (lColor1 \ &H100&) And &HFF&
    VX(0).B.Hi = (lColor1 \ &H10000) And &HFF&
    
    VX(1).X = X2
    VX(1).Y = Y2
    VX(1).R.Hi = lColor2 And &HFF&
    VX(1).G.Hi = (lColor2 \ &H100&) And &HFF&
    VX(1).B.Hi = (lColor2 \ &H10000) And &HFF&

    MR.v1 = 0
    MR.v2 = 1
    
    GradientSimple = GradientFillRect(hDC, VX(0), 2, MR, 1, lVertical)
    
End Function

Public Function Gradient(ByVal hDC As Long, _
                         ByVal X1 As Long, ByVal Y1 As Long, _
                         ByVal X2 As Long, ByVal Y2 As Long, _
                         ByVal lVertical As Long, ByVal lRepeats As Long, _
                         ByVal iInterp As Integer, _
                         ParamArray lColors() As Variant) As Long
                                           
' USE  : Draws a repeating vertical or horizontal gradient between
'        any number of colors using GradientFillRect API function or
'        a user defined function (to provide cosine and HLS color interpolation).

' If you do not want to use custom interpolation you can modify the last
' lines of this function and totally remove the GradientFillCustom function as
' well as the LineTo, MoveTo, CreatePen, SelectObject, and DeleteObject API functions
' and the mColor module.

Dim VX()            As VERTEX
Dim MR()            As MESH_RECT
Dim i               As Long
Dim j               As Long
Dim k               As Long
Dim iRem            As Integer
Dim lQuo            As Long
Dim lChunks()       As Long
Dim lWidth          As Long
Dim lHeight         As Long
Dim lSubChunks()    As Long
Dim lCols()         As Long
Dim iIdx            As Integer
Dim iInc            As Integer

    lWidth = X2 - X1
    lHeight = Y2 - Y1
    
    If lRepeats < 0 Then Exit Function
    
    ' Transfer all valid colors to a new array.  This is really only necessary
    ' due to the structure of the demo.  You can take this bit if you know you
    ' are going to be passing in all valid colors.
    For i = 0 To UBound(lColors)
        If lColors(i) <> -1 Then
            ReDim Preserve lCols(j)
            lCols(j) = lColors(i)
            j = j + 1
        End If
    Next
    
    ' Exit if one or no color(s).
    If j < 2 Then Exit Function
    
    ' This bit solves the "bit-left-over-at-the-end" problem.  It basically
    ' divides the width or height into same sized chunks.  Then it evenly
    ' distributes any remainder amongst the chunks.
    
    ' EX: If lWidth = 100 & lRepeats = 2, then you need 3 chunks, with
    '     chunk(0) = 33, chunk(1) = 33, and chunk(2) = 34.

    ' Now you must also create same sized subchunks for each color within
    ' each chunk.
    
    If lVertical Then
        iRem = lHeight Mod (lRepeats + 1)
        lQuo = lHeight \ (lRepeats + 1)
    Else
        iRem = lWidth Mod (lRepeats + 1)
        lQuo = lWidth \ (lRepeats + 1)
    End If

    ' Create the chunks for the number of repeats.
    ReDim lChunks(lRepeats)
    
    For i = 0 To UBound(lChunks)
        lChunks(i) = lQuo
        If UBound(lChunks) - i < iRem Then lChunks(i) = lChunks(i) + 1
    Next

    ' Create the subchunks for each color within each chunk.  The number of subchunks
    ' need came from a lot of trial and error and drawing diagrams on paper.
    ReDim lSubChunks(UBound(lCols) * (lRepeats + 1) - 1)
    
    ' For each chunk..
    For i = 0 To UBound(lChunks)
        
        iRem = lChunks(i) Mod UBound(lCols)
        lQuo = lChunks(i) \ UBound(lCols)
        
        ' Size the subchunks and apply remainder evenly.
        For j = 0 To UBound(lCols) - 1
            k = i * UBound(lCols) + j
            lSubChunks(k) = lQuo
            If (UBound(lCols) - 1) - j < iRem Then lSubChunks(k) = lSubChunks(k) + 1
        Next
    Next
    
    ' Dimension arrays for the meshes and their verticies.
    ReDim MR(UBound(lCols) * (lRepeats + 1) - 1)
    ReDim VX(UBound(MR) * 2 + 1)
    
    iInc = 1
    
    For i = 0 To UBound(VX) Step 2
    
        ' Set the properties of the top left vertex.
        If lVertical Then
            VX(i).X = X1
            If i Then VX(i).Y = VX(i - 1).Y Else VX(i).Y = Y1
        Else
            If i Then VX(i).X = VX(i - 1).X Else VX(i).X = X1
            VX(i).Y = Y1
        End If
        
        VX(i).R.Hi = lCols(iIdx) And &HFF&
        VX(i).G.Hi = (lCols(iIdx) \ &H100&) And &HFF&
        VX(i).B.Hi = (lCols(iIdx) \ &H10000) And &HFF&
        
        ' Increment/decrement the current color index.
        iIdx = iIdx + iInc
        If iIdx = 0 Or (iIdx Mod UBound(lCols) = 0) Then iInc = -iInc
        
        ' Set the properties of the bottom right vertex.
        If lVertical Then
            VX(i + 1).X = X2
            VX(i + 1).Y = VX(i).Y + lSubChunks(i \ 2)
        Else
            VX(i + 1).X = VX(i).X + lSubChunks(i \ 2)
            VX(i + 1).Y = Y2
        End If
        
        VX(i + 1).R.Hi = lCols(iIdx) And &HFF&
        VX(i + 1).G.Hi = (lCols(iIdx) \ &H100&) And &HFF&
        VX(i + 1).B.Hi = (lCols(iIdx) \ &H10000) And &HFF&
        
        ' Assign the meshes' verticies.
        MR(i \ 2).v1 = i
        MR(i \ 2).v2 = i + 1
    Next
    
    If iInterp Then
        Gradient = GradientFillCustom(hDC, VX, MR, lVertical, iInterp)
    Else
        Gradient = GradientFillRect(hDC, VX(0), UBound(VX) + 1, MR(0), UBound(MR) + 1, lVertical)
    End If
    
End Function

Private Function GradientFillCustom(ByVal hDC As Long, _
                                    ByRef VX() As VERTEX, ByRef MR() As MESH_RECT, _
                                    ByVal lVertical As Long, ByVal iInterp As Integer) As Long

' Uses the same meshes and verticies arguments defined for the GradientFillRect API
' function.  The bonus with this function is that it provides smoother cosine
' color interpolation.  Using this function with iInterp=0 (linear color
' interpolation) produces identical results as GradientFillRect.

Dim v1          As Long
Dim v2          As Long
Dim i           As Long
Dim j           As Long
Dim lCols()     As Long
Dim p           As Single, p2   As Single
Dim R1          As Byte, G1     As Byte, B1 As Byte
Dim R2          As Byte, G2     As Byte, B2 As Byte
Dim lDelta      As Long
Dim hPN         As Long
Dim hPO         As Long
Dim lRet        As Long
Dim zHLS1(2)    As Single
Dim zHLS2(2)    As Single

 ' All temporary variables are for simplicity and readability.
 
    ' For each mesh..
    For i = 0 To UBound(MR)
    
        v1 = MR(i).v1: v2 = MR(i).v2
    
        ' Dimesion an arary to hold the colors of the mesh.
        lDelta = IIf(lVertical, VX(v2).Y - VX(v1).Y - 1, VX(v2).X - VX(v1).X - 1)
        ReDim lCols(lDelta)
        
        ' Interpolate from the top left vertex color to the bottom right vertex color.
        For j = 0 To lDelta
            
            R1 = VX(v1).R.Hi: G1 = VX(v1).G.Hi: B1 = VX(v1).B.Hi
            R2 = VX(v2).R.Hi: G2 = VX(v2).G.Hi: B2 = VX(v2).B.Hi
            
            p = j / lDelta
            If iInterp = 1 Then p = (1 - Cos(p * PI)) * 0.5
            p2 = 1 - p
            
            ' Calculate the color at the current position within the mesh.
            If iInterp = 2 Then
            
                ' HLS interpolation.
                Call RGBtoHLS(R1, G1, B1, zHLS1())
                Call RGBtoHLS(R2, G2, B2, zHLS2())
                
                lCols(j) = HLStoLNG(zHLS1(0) * p2 + zHLS2(0) * p, _
                                    zHLS1(1) * p2 + zHLS2(1) * p, _
                                    zHLS1(2) * p2 + zHLS2(2) * p)
            Else
            
                ' Linear and cosine interpolation.
                lCols(j) = RGB(R1 * p2 + R2 * p, G1 * p2 + G2 * p, B1 * p2 + B2 * p)
                
            End If
            
            ' Draw the line in that color.
            hPN = CreatePen(0, 1, lCols(j)): Debug.Assert hPN
            hPO = SelectObject(hDC, hPN): Debug.Assert hPO
                
            If lVertical Then
                MoveTo hDC, VX(v1).X, VX(v1).Y + j, ByVal 0&
                LineTo hDC, VX(v2).X, VX(v1).Y + j
            Else
                MoveTo hDC, VX(v1).X + j, VX(v1).Y, ByVal 0&
                LineTo hDC, VX(v1).X + j, VX(v2).Y
            End If
            
            lRet = SelectObject(hDC, hPO): Debug.Assert lRet
            lRet = DeleteObject(hPN): Debug.Assert lRet
        Next
    Next
    
    ' TODO: More error checking.
    GradientFillCustom = True
    
End Function
