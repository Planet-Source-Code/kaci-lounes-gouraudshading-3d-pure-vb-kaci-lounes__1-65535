Attribute VB_Name = "Rasterizations"

' MODULE NAME: Rasterization.BAS
' ==============================
'
' Module of raterizations functions.

Option Explicit
Sub DrawGradientTriangle(DestDC As Long, A As Point2D, B As Point2D, C As Point2D, ColA As ColRGB, ColB As ColRGB, ColC As ColRGB)

 ' SUB : DrawGradientTriangle
 ' ==========================
 '
 ' RETURNED VALUE: None
 '
 ' Draw a gradient triangle on a specified device context (DC).
 ' This sub use the GradientFill API (MSIMG32.DLL) for drawing.
 ' This is a linearily interpolation of vertices colors across
 '  the triangle, we can do this by a simple scan-conversion algorithm,
 '   and the BaryInterpolateLinear function.

 Dim Vertices(2) As TRIVERTEX, Triangle As GRADIENT_TRIANGLE

 Vertices(0).X = A.X: Vertices(0).Y = A.Y
 Vertices(0).Red = "&H" & Hex(ColA.R) & "00"
 Vertices(0).Green = "&H" & Hex(ColA.G) & "00"
 Vertices(0).Blue = "&H" & Hex(ColA.B) & "00"

 Vertices(1).X = B.X: Vertices(1).Y = B.Y
 Vertices(1).Red = "&H" & Hex(ColB.R) & "00"
 Vertices(1).Green = "&H" & Hex(ColB.G) & "00"
 Vertices(1).Blue = "&H" & Hex(ColB.B) & "00"

 Vertices(2).X = C.X: Vertices(2).Y = C.Y
 Vertices(2).Red = "&H" & Hex(ColC.R) & "00"
 Vertices(2).Green = "&H" & Hex(ColC.G) & "00"
 Vertices(2).Blue = "&H" & Hex(ColC.B) & "00"

 Triangle.Vertex1 = 0: Triangle.Vertex2 = 1: Triangle.Vertex3 = 2

 GradientFill DestDC, Vertices(0), 3, Triangle, 1, GradientTriangle

End Sub
