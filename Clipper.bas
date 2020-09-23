Attribute VB_Name = "Clipper"

' MODULE NAME: Clipper.BAS
' ========================
'
' This module includ all the necessary
'  clipping operations for rendering.

Option Explicit
Function BaryInterpolateLinear(X1&, Y1&, Val1%, X2&, Y2&, Val2%, X3&, Y3&, Val3%, PX&, PY&) As Single

 ' FUNCTION : BaryInterpolateLinear
 ' ================================
 '
 ' RETURNED VALUE: Single
 '
 ' 2D BaryCentric triangle interpolation (Môbius 1827)
 '
 ' Barycentric is a best way for interpolating vertices or values.
 ' In a convex polygon, we use these interpolations to find others
 '  pre-vertex values (useful in clipping such as: Zs, colors, texture
 '   coordinates...ect).
 '
 '  In this function, we use a special case for triangles.
 '
 ' There are two ways of barycentric:
 '
 ' 1. Map a point by the weight corners:
 '
 '    By giving three points (2D or 3D or nD, and not colinear) and the
 '     weight corners (U, V & W for a triangle, such that U+V+W = 1.0,
 '                     also called Barycentric coordinates),
 '      we can then map the interpolated point, an example in 3D:
 '
 '      Interpolation.X = (U * V1.X) + (V * V2.X) + (W * V3.X)
 '      Interpolation.Y = (U * V1.Y) + (V * V2.Y) + (W * V3.Y)
 '      Interpolation.Z = (U * V1.Z) + (V * V2.Z) + (W * V3.Z)
 '
 ' 2. Map any values by three points and a four point in the triangle:
 '
 '    By giving three points and a point WITHIN this triangle
 '     (AND co-planar for a 3D triangle), we can find the weight
 '      corners (U, V & W), then, we can interpolate any value,
 '       if for example we have a 2D triangle, each point has
 '        a Z coordinate, then we can map the Z coordinate for
 '         the new point suh that:
 '
 '      Z = (U * Z1) + (V * Z2) + (W * Z3)
 '
 '    Or any other values:
 '
 '      ColR = (U * ColR1) + (V * ColR2) + (W * ColR3)
 '      ColG = (U * ColG1) + (V * ColG2) + (W * ColG3)
 '      ColB = (U * ColB1) + (V * ColB2) + (W * ColB3)
 '
 '      TexU = (U * TexU1) + (V * TexU2) + (W * TexU3)
 '      TexV = (U * TexV1) + (V * TexV2) + (W * TexV3)
 '
 ' So what is these UVW values ?!
 '
 '  If we compute: (U+V+W) we get 1.0, this is the parametricaly
 '   representation of the total area of the triangle.
 '
 '  The UVW coordinates are the RATIOS of this area.
 '   For example: U=0.5, V=0.3, W = 1-(U+V) ==> W = 1-(0.5+0.3) = 0.2
 '
 ' For be more understand, view the sheme:
 '         .
 '        /|\
 '       / | \
 '      /  |  \
 '     / U | V \
 '    /   / \   \
 '   /  /     \  \
 '  / /   W     \ \
 ' //_____________\\
 '
 ' So how this function find the UVW values ?
 '
 ' We need the area, or a simple ratio of the area of this triangle.
 '  There are many ways for computing the area, the cramers ruls,
 '   Heron's formula (based on the length of sides) for examples.
 '
 '  But we use the theory: "The magnitude of cross product is
 '                          the twice area of a 3D triangle"
 '
 ' Based on this, we:
 '
 '  - Convert the triangle to a 3D triangle (with zeros for Zs)
 '
 '  - Compute the cross product (with zeros for Zs, there is no
 '                               XY coordinates for the cross product,
 '                               only the Z coordinate, also, the
 '                               magnitude is simply the Z coordinate)
 '
 '     CrossZ = CrossProduct(Subtract(B,A), Subtract(C,A)).Z
 '
 '  - Compute the others cross products:
 '
 '     CrossZ1 = CrossProduct(Subtract(A,P), Subtract(B,P)).Z
 '     CrossZ2 = CrossProduct(Subtract(B,P), Subtract(C,P)).Z
 '     CrossZ3 = CrossProduct(Subtract(C,P), Subtract(A,P)).Z  (optional)
 '
 '  - Divide:
 '
 '     U = CrossZ1 / CrossZ
 '     V = CrossZ2 / CrossZ
 '     W = CrossZ3 / CrossZ or W = 1 - (U + V)
 '
 ' Note that this Barycentric function use a LINEARLY interpolation.
 '  So, if you use this fuction for texture mapping, you get a LINEARLY
 '   texture mapping (affine), not perspective-correction are done here
 '    by this function (to do this, we need the hyperbolic interpolation).
 '
 ' Return: The interpolated value.

 Dim D!, U!, V!, W!

 D = 1 / (((X2 - X1) * (Y3 - Y1)) - ((Y2 - Y1) * (X3 - X1)))
 U = (((X2 - PX) * (Y3 - PY)) - ((Y2 - PY) * (X3 - PX))) * D
 V = (((X3 - PX) * (Y1 - PY)) - ((Y3 - PY) * (X1 - PX))) * D
 W = 1 - (U + V)

 BaryInterpolateLinear = (U * Val1) + (V * Val2) + (W * Val3)

End Function
Function IsInside2D(PX!, PY!, RX1!, RY1!, RX2!, RY2!) As Boolean

 ' FUNCTION : IsInside2D
 ' =====================
 '
 ' RETURNED VALUE: Boolean
 '
 ' Check if a 2D point is inside a rectangle

 If (PX > RX1) And (PX < RX2) Then
  If (PY > RY1) And (PY < RY2) Then IsInside2D = True
 End If

End Function
Sub ClipTriangle2D(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!, X3!, Y3!, OutPts() As Point2D, NumPts As Byte, Stat As Byte)

 ' SUB : ClipTriangle2D
 ' ====================
 '
 ' RETURNED VALUE:
 '
 ' - Parameters : OutPts() : Point2D, NumPts; Stat : Byte
 '
 ' Triangle clipping steps :
 '
 ' 1- Elimenate three cases (Completly inside, Region is IN A BIG triangle, Completly outside)
 '
 ' 2- (Cases else): Clip line-by-line AB, BC, CA
 ' 3- Check if the region points boundaries is inside this triangle, if yes:
 '     add this point to output array (there are four, but 3 as maximum):
 '     (RX1-RY1, RX1-RY2, RX2-RY1, RX2-RY2)
 ' 4- Remove the doubly points from output array.
 ' 5- Set points to clockwise direction (not currently included).
 '
 ' Return: A polygon of NumPts points stored in OutPts() array.

 Dim I%, J%, K%
 Dim AOX1!, AOY1!, AOX2!, AOY2!
 Dim BOX1!, BOY1!, BOX2!, BOY2!
 Dim COX1!, COY1!, COX2!, COY2!
 Dim A1 As Byte, A2 As Byte, A3 As Byte 'Trivial Accept2D
 Dim R1 As Byte, R2 As Byte, R3 As Byte 'Trivial Reject2D

 'Get trivial cases
 '=================
 If Accept2D(RX1, RY1, RX2, RY2, X1, Y1, X2, Y2) = True Then A1 = 1 Else A1 = 0
 If Accept2D(RX1, RY1, RX2, RY2, X2, Y2, X3, Y3) = True Then A2 = 1 Else A2 = 0
 If Accept2D(RX1, RY1, RX2, RY2, X3, Y3, X1, Y1) = True Then A3 = 1 Else A3 = 0

 AOX1 = 0: AOY1 = 0: AOX2 = 0: AOY2 = 0
 ClipLine RX1, RY1, RX2, RY2, X1, Y1, X2, Y2, AOX1, AOY1, AOX2, AOY2
 If AOX1 = 0 And AOY1 = 0 And AOX2 = 0 And AOY2 = 0 Then R1 = 1

 BOX1 = 0: BOY1 = 0: BOX2 = 0: BOY2 = 0
 ClipLine RX1, RY1, RX2, RY2, X2, Y2, X3, Y3, BOX1, BOY1, BOX2, BOY2
 If BOX1 = 0 And BOY1 = 0 And BOX2 = 0 And BOY2 = 0 Then R2 = 1

 COX1 = 0: COY1 = 0: COX2 = 0: COY2 = 0
 ClipLine RX1, RY1, RX2, RY2, X3, Y3, X1, Y1, COX1, COY1, COX2, COY2
 If COX1 = 0 And COY1 = 0 And COX2 = 0 And COY2 = 0 Then R3 = 1
 '===========================================================================

 'Completly inside
 '================
 If A1 = 0 And A2 = 0 And A3 = 0 And _
    R1 = 0 And R2 = 0 And R3 = 0 Then
  ReDim OutPts(2)
  OutPts(0).X = X1: OutPts(0).Y = Y1
  OutPts(1).X = X2: OutPts(1).Y = Y2
  OutPts(2).X = X3: OutPts(2).Y = Y3
  NumPts = 2: Stat = 1
  Exit Sub
 End If

 'Region is in a BIG triangle
 '===========================
 If IsInsideTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, RX1, RY1) = True And _
    IsInsideTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, RX2, RY1) = True And _
    IsInsideTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, RX2, RY2) = True And _
    IsInsideTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, RX1, RY2) = True Then
  ReDim OutPts(3)
  OutPts(0).X = RX1: OutPts(0).Y = RY1
  OutPts(1).X = RX2: OutPts(1).Y = RY1
  OutPts(2).X = RX2: OutPts(2).Y = RY2
  OutPts(3).X = RX1: OutPts(3).Y = RY2
  NumPts = 3
  Stat = 2
  Exit Sub
 End If

 'Completly outside
 '=================
 If A1 = 1 And A2 = 1 And A3 = 1 And _
    R1 = 1 And R2 = 1 And R3 = 1 Then NumPts = 0: Stat = 0: Exit Sub

 '===========================================================================

 Stat = 2
 ReDim OutPts(0)

 'Edge1: AB =================================================================
 If A1 = 1 Then
  If R1 = 0 Then
   ReDim OutPts(1)
   OutPts(0).X = AOX1: OutPts(0).Y = AOY1
   OutPts(1).X = AOX2: OutPts(1).Y = AOY2
  End If
 Else
  ReDim OutPts(1)
  OutPts(0).X = X1: OutPts(0).Y = Y1
  OutPts(1).X = X2: OutPts(1).Y = Y2
 End If

 'Edge2: BC =================================================================
 If A2 = 1 Then
  If R2 = 0 Then
   If UBound(OutPts) = 0 Then
    ReDim OutPts(1)
    OutPts(0).X = BOX1: OutPts(0).Y = BOY1
    OutPts(1).X = BOX2: OutPts(1).Y = BOY2
   Else
    ReDim Preserve OutPts(UBound(OutPts) + 2)
    OutPts(UBound(OutPts) - 1).X = BOX1
    OutPts(UBound(OutPts) - 1).Y = BOY1
    OutPts(UBound(OutPts)).X = BOX2
    OutPts(UBound(OutPts)).Y = BOY2
   End If
  End If
 Else
  If UBound(OutPts) = 0 Then
   ReDim OutPts(1)
   OutPts(0).X = X2: OutPts(0).Y = Y2
   OutPts(1).X = X3: OutPts(1).Y = Y3
  Else
   ReDim Preserve OutPts(UBound(OutPts) + 2)
   OutPts(UBound(OutPts) - 1).X = X2
   OutPts(UBound(OutPts) - 1).Y = Y2
   OutPts(UBound(OutPts)).X = X3
   OutPts(UBound(OutPts)).Y = Y3
  End If
 End If

 'Edge3: CA =================================================================
 If A3 = 1 Then
  If R3 = 0 Then
   If UBound(OutPts) = 0 Then
    ReDim OutPts(1)
    OutPts(0).X = COX1: OutPts(0).Y = COY1
    OutPts(1).X = COX2: OutPts(1).Y = COY2
   Else
    ReDim Preserve OutPts(UBound(OutPts) + 2)
    OutPts(UBound(OutPts) - 1).X = COX1
    OutPts(UBound(OutPts) - 1).Y = COY1
    OutPts(UBound(OutPts)).X = COX2
    OutPts(UBound(OutPts)).Y = COY2
   End If
  End If
 Else
  If UBound(OutPts) = 0 Then
   ReDim OutPts(1)
   OutPts(0).X = X3: OutPts(0).Y = Y3
   OutPts(1).X = X1: OutPts(1).Y = Y1
  Else
   ReDim Preserve OutPts(UBound(OutPts) + 2)
   OutPts(UBound(OutPts) - 1).X = X3
   OutPts(UBound(OutPts) - 1).Y = Y3
   OutPts(UBound(OutPts)).X = X1
   OutPts(UBound(OutPts)).Y = Y1
  End If
 End If

 '===========================================================================

 If IsInsideTriangle(X1, Y1, X2, Y2, X3, Y3, RX1, RY1) = True Then
  ReDim Preserve OutPts(UBound(OutPts) + 1)
  OutPts(UBound(OutPts)).X = RX1
  OutPts(UBound(OutPts)).Y = RY1
 End If

 If IsInsideTriangle(X1, Y1, X2, Y2, X3, Y3, RX2, RY1) = True Then
  ReDim Preserve OutPts(UBound(OutPts) + 1)
  OutPts(UBound(OutPts)).X = RX2
  OutPts(UBound(OutPts)).Y = RY1
 End If

 If IsInsideTriangle(X1, Y1, X2, Y2, X3, Y3, RX2, RY2) = True Then
  ReDim Preserve OutPts(UBound(OutPts) + 1)
  OutPts(UBound(OutPts)).X = RX2
  OutPts(UBound(OutPts)).Y = RY2
 End If

 If IsInsideTriangle(X1, Y1, X2, Y2, X3, Y3, RX1, RY2) = True Then
  ReDim Preserve OutPts(UBound(OutPts) + 1)
  OutPts(UBound(OutPts)).X = RX1
  OutPts(UBound(OutPts)).Y = RY2
 End If

 'Remove doubly points
 '====================
ReCheck:
 For I = LBound(OutPts) To UBound(OutPts)
  For J = I To UBound(OutPts)
   If (OutPts(I).X = OutPts(J).X) And (OutPts(I).Y = OutPts(J).Y) And (J <> I) Then
    For K = I To UBound(OutPts) - 1
     OutPts(K).X = OutPts(K + 1).X
     OutPts(K).Y = OutPts(K + 1).Y
    Next K
    ReDim Preserve OutPts(UBound(OutPts) - 1)
    GoTo ReCheck
   End If
  Next J
 Next I

 NumPts = UBound(OutPts)

End Sub
Function IsInsideTriangle(X1!, Y1!, X2!, Y2!, X3!, Y3!, PX!, PY!) As Boolean

 ' FUNCTION : IsInsideTriangle
 ' ===========================
 '
 ' RETURNED VALUE: Boolean
 '
 ' Check if a 2D point is inside a 2D triangle.

 Dim CRZ1!, CRZ2!, CRZ3!

 CRZ1 = (((X2 - PX) * (Y3 - PY)) - ((Y2 - PY) * (X3 - PX)))
 CRZ2 = (((X1 - PX) * (Y2 - PY)) - ((Y1 - PY) * (X2 - PX)))
 CRZ3 = (((X3 - PX) * (Y1 - PY)) - ((Y3 - PY) * (X1 - PX)))

 'The point is inside the triangle
 ' if the vars (CRZ1, CRZ2 & CRZ3)
 '  has the same sign:
 If ((CRZ1 > 0) And (CRZ2 > 0) And (CRZ3 > 0)) Or _
    ((CRZ1 < 0) And (CRZ2 < 0) And (CRZ3 < 0)) Then IsInsideTriangle = True

End Function
Sub ClipLine(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!, OutX1!, OutY1!, OutX2!, OutY2!)

 ' SUB : ClipLine
 ' ==============
 '
 ' RETURNED VALUE:
 '
 ' - Parameters:
 '
 '   OutX1, OutY1, OutX2, OutY2 : The coordinates of the clipped line (Output)
 '
 ' Liang Barsky Line Clipping Algorithm (1984)
 ' (Parametric clipping but special case for 2D RECTANGULAR clipping regions)
 '
 ' Note that for fast checking the trivial cases, I use simply
 '  Cohen-Sutherland (codes), But for clipping the line, I use
 '   directly Liang-Barsky algorithm (parametric without codes)
 '
 ' Also note that routine is *Very* optimized!
 '  Realy, it is two modules that i have it
 '   transformed in only one procedure!!!!!

 Dim PX1!, PY1!, PX2!, PY2!, U1!, U2!, Dx!, Dy!, P!, Q!, R!, Temp!, CT As Byte

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 U1 = 0: U2 = 1: PX1 = X1: PY1 = Y1: PX2 = X2: PY2 = Y2
 Dx = (PX2 - PX1): Dy = (PY2 - PY1)

 P = -Dx: Q = (PX1 - RX1)
 If (P < 0) Then
  R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
 ElseIf (P > 0) Then
  R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
 ElseIf (Q < 0) Then
  CT = 1
 End If
 If CT = 0 Then
  P = Dx: Q = (RX2 - PX1)
  If (P < 0) Then
   R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
  ElseIf (P > 0) Then
   R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
  ElseIf (Q < 0) Then
   CT = 1
  End If
  If CT = 0 Then
   P = -Dy: Q = (PY1 - RY1)
   If (P < 0) Then
    R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
   ElseIf (P > 0) Then
    R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
   ElseIf (Q < 0) Then
    CT = 1
   End If
   If CT = 0 Then
    P = Dy: Q = (RY2 - PY1)
    If (P < 0) Then
     R = (Q / P): If (R > U2) Then CT = 1 Else If (R > U1) Then U1 = R
    ElseIf (P > 0) Then
     R = (Q / P): If (R < U1) Then CT = 1 Else If (R < U2) Then U2 = R
    ElseIf (Q < 0) Then
     CT = 1
    End If
    If CT = 0 Then
     If (U2 < 1) Then PX2 = (PX1 + (U2 * Dx)): PY2 = (PY1 + (U2 * Dy))
     If (U1 > 0) Then PX1 = (PX1 + (U1 * Dx)): PY1 = (PY1 + (U1 * Dy))
     OutX1 = PX1: OutY1 = PY1: OutX2 = PX2: OutY2 = PY2
    End If
   End If
  End If
 End If

End Sub
Function Accept2D(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!) As Boolean

 ' FUNCTION : Accept2D
 ' ===================
 '
 ' RETURNED VALUE: Boolean
 '
 ' Cohen-Sutherland trivial Accept2D (with codes)
 '  Use the 'Or' logical operator.

 Dim Code1(3) As Boolean
 Dim Code2(3) As Boolean
 Dim Temp As Single

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 If (X1 < RX1) Then Code1(0) = True Else Code1(0) = False
 If (X1 > RX2) Then Code1(1) = True Else Code1(1) = False
 If (Y1 < RY1) Then Code1(2) = True Else Code1(2) = False
 If (Y1 > RY2) Then Code1(3) = True Else Code1(3) = False

 If (X2 < RX1) Then Code2(0) = True Else Code2(0) = False
 If (X2 > RX2) Then Code2(1) = True Else Code2(1) = False
 If (Y2 < RY1) Then Code2(2) = True Else Code2(2) = False
 If (Y2 > RY2) Then Code2(3) = True Else Code2(3) = False

 If (Code1(0) Or Code2(0)) Then Accept2D = True
 If (Code1(1) Or Code2(1)) Then Accept2D = True
 If (Code1(2) Or Code2(2)) Then Accept2D = True
 If (Code1(3) Or Code2(3)) Then Accept2D = True

End Function
Function Reject2D(RX1!, RY1!, RX2!, RY2!, X1!, Y1!, X2!, Y2!) As Boolean

 ' FUNCTION : Reject2D
 ' ===================
 '
 ' RETURNED VALUE: Boolean
 '
 ' Cohen-Sutherland trivial Reject2D (with codes)
 '  Use the 'And' logical operator.

 Dim Code1(3) As Boolean
 Dim Code2(3) As Boolean
 Dim Temp As Single

 If (RX1 > RX2) Then Temp = RX1: RX1 = RX2: RX2 = Temp
 If (RY1 > RY2) Then Temp = RY1: RY1 = RY2: RY2 = Temp

 If (X1 < RX1) Then Code1(0) = True Else Code1(0) = False
 If (X1 > RX2) Then Code1(1) = True Else Code1(1) = False
 If (Y1 < RY1) Then Code1(2) = True Else Code1(2) = False
 If (Y1 > RY2) Then Code1(3) = True Else Code1(3) = False

 If (X2 < RX1) Then Code2(0) = True Else Code2(0) = False
 If (X2 > RX2) Then Code2(1) = True Else Code2(1) = False
 If (Y2 < RY1) Then Code2(2) = True Else Code2(2) = False
 If (Y2 > RY2) Then Code2(3) = True Else Code2(3) = False

 If (Code1(0) And Code2(0)) Then Reject2D = True
 If (Code1(1) And Code2(1)) Then Reject2D = True
 If (Code1(2) And Code2(2)) Then Reject2D = True
 If (Code1(3) And Code2(3)) Then Reject2D = True

End Function
