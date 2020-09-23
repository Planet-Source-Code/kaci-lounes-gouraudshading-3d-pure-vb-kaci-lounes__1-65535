Attribute VB_Name = "Maths3D"

' MODULE NAME: Math3D.BAS
' =======================
'
' The 3D maths, colors computations, APIs declarations,
'  constants & data structures can be found in this module.

Option Explicit

' RETURNED VALUES:
'
' - All the functions that the name start by the word Color, returns ColRGB stucture.
' - All the functions that the name start by the word Matrix, returns Matrix stucture.
' - All the functions that the name start by the word Vector, returns Vector3D stucture.
'
' =====================================================================
' =============== APIS CALLS, STRUCTURES & CONSTANTS ==================
' =====================================================================

'Show/Hide cursor API:
Public Declare Function ShowCursor Lib "USER32.DLL" (ByVal bShow As Long) As Long

'Display settings APIs:
Public Declare Function GetDeviceCaps Lib "GDI32.DLL" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function EnumDisplaySettings Lib "USER32.DLL" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "USER32.DLL" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Public Type DEVMODE             '(122 Bytes)
 dmDeviceName As String * 32
 dmSpecVersion As Integer
 dmDriverVersion As Integer
 dmSize As Integer
 dmDriverExtra As Integer
 dmFields As Long
 dmOrientation As Integer
 dmPaperSize As Integer
 dmPaperLength As Integer
 dmPaperWidth As Integer
 dmScale As Integer
 dmCopies As Integer
 dmDefaultSource As Integer
 dmPrintQuality As Integer
 dmColor As Integer
 dmDuplex As Integer
 dmYResolution As Integer
 dmTTOption As Integer
 dmCollate As Integer
 dmFormName As String * 32
 dmUnusedPadding As Integer
 dmBitsPerPel As Integer
 dmPelsWidth As Long
 dmPelsHeight As Long
 dmDisplayFlags As Long
 dmDisplayFrequency As Long
End Type

'GradientFill API:
Public Declare Function GradientFill Lib "MSIMG32.DLL" (ByVal hdc As Long, PVertex As TRIVERTEX, ByVal dwNumVertex As Long, PMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Boolean

Global Const GradientTriangle As Long = 2

Public Type TRIVERTEX          '(16 Bytes)
 X As Long
 Y As Long
 Red As Integer
 Green As Integer
 Blue As Integer
 Alpha As Integer
End Type

Public Type GRADIENT_TRIANGLE  '(12 Bytes)
 Vertex1 As Long
 Vertex2 As Long
 Vertex3 As Long
End Type

' =====================================================================
' ========================== GLOBAL CONSTANTS =========================
' =====================================================================

Global Const Pi As Single = 3.141593
Global Const Deg As Single = 0.0174532
Global Const NaturalLogBase As Single = 2.718282
Global Const OneByThree As Single = 0.3333333
Global Const OneBy255 As Single = 0.0039215
Global Const ApproachVal As Single = 0.0000001

' =====================================================================
' ======================== DATAS STRUCTURES ===========================
' =====================================================================

Public Type Matrix        '(64 Bytes)
 M11 As Single: M12 As Single: M13 As Single: M14 As Single
 M21 As Single: M22 As Single: M23 As Single: M24 As Single
 M31 As Single: M32 As Single: M33 As Single: M34 As Single
 M41 As Single: M42 As Single: M43 As Single: M44 As Single
End Type

Public Type Vector3D      '(12 Bytes)
 X As Single
 Y As Single
 Z As Single
End Type

Public Type Point2D       '(8 Bytes)
 X As Long
 Y As Long
End Type

Public Type Triangle      '(For triangulation, 6 Bytes)
 A As Integer
 B As Integer
 C As Integer
End Type

Public Type ColRGB        '(6 Bytes)
 R As Integer
 G As Integer
 B As Integer
End Type

Public Type Vertex3D      '(37 Bytes)
 OrgPos As Vector3D
 TmpPos As Vector3D
 OrgCol As ColRGB
 ShdCol As ColRGB
 OldShaded As Byte
End Type

Public Type Face          '(18 Bytes)
 A As Integer
 B As Integer
 C As Integer
 Normal As Vector3D
End Type

Public Type Mesh          '(Variant memory size)
 Origin As Vector3D
 Scales As Vector3D
 Pitch As Single
 Yaw As Single
 Roll As Single
 DiffsReflct As Single
 SpecReflctK As Single
 SpecReflctN As Single
 Vertices() As Vertex3D
 Faces() As Face
 WorldMatrix As Matrix
End Type

Public Type SpotLight3D   '(66 Bytes)
 Origin As Vector3D
 Direction As Vector3D
 Diffuse As ColRGB
 Specular As ColRGB
 Ambiance As ColRGB
 Density As Single
 Falloff As Single
 Hotspot As Single
 ConstantRange As Single
 LinearRange As Single
 AttenEnable As Boolean
 Enabled As Boolean
End Type
Function ColorDiffuse(ColA As ColRGB, ColB As ColRGB) As ColRGB

 ' When use diffuse lighting, for example,
 '  if we make a red sphere, and a blue light source,
 '   and the ambiant light is disable, we get a BLACK
 '    color (no light) on the sphere.
 ' (you can test this in any 3D software)
 '
 ' The ColorDiffuse function do the same thing,
 '  this is helpful when we compute the diffuse value.

 ColorDiffuse.R = (ColA.R * (ColB.R * OneBy255))
 ColorDiffuse.G = (ColA.G * (ColB.G * OneBy255))
 ColorDiffuse.B = (ColA.B * (ColB.B * OneBy255))

End Function
Function ColorInput(Red%, Green%, Blue%) As ColRGB

 ColorInput.R = Red
 ColorInput.G = Green
 ColorInput.B = Blue

End Function
Function ColorLimit(Col As ColRGB) As ColRGB

 ColorLimit.R = Col.R: If (ColorLimit.R > 255) Then ColorLimit.R = 255 Else If (ColorLimit.R < 0) Then ColorLimit.R = 0
 ColorLimit.G = Col.G: If (ColorLimit.G > 255) Then ColorLimit.G = 255 Else If (ColorLimit.G < 0) Then ColorLimit.G = 0
 ColorLimit.B = Col.B: If (ColorLimit.B > 255) Then ColorLimit.B = 255 Else If (ColorLimit.B < 0) Then ColorLimit.B = 0

End Function
Function VectorDistance3(VecA As Vector3D, VecB As Vector3D) As Single

 VectorDistance3 = VectorLength3(VectorSubtract3(VecA, VecB))

End Function
Function ColorAdd(ColA As ColRGB, ColB As ColRGB) As ColRGB

 ColorAdd.R = (ColA.R + ColB.R)
 ColorAdd.G = (ColA.G + ColB.G)
 ColorAdd.B = (ColA.B + ColB.B)

End Function
Function ColorInterpolate(ColA As ColRGB, ColB As ColRGB, Alpha As Single) As ColRGB

 ' A linear 24-Bits color (RGB) interpolation:
 ' (Pixel by pixel, we get AlphaBlending !)
 '
 ' Set Interpolation = A + ((B-A)*t)
 '
 ' The time parameter t (named Alpha in the top)
 '  is the coordinate interval, vary between 0 to 1,
 '
 ' - At time t (or Alpha) is 0, we get ColorA
 ' - At time t (or Alpha) is 1, we get ColorB
 ' - At time t (or Alpha) is 0.##, we get Color = ColorA + ((ColorB-ColorA)*t)
 '
 ' (Can also use this linear interpolation for vectors)

 ColorInterpolate.R = ColA.R + ((ColB.R - ColA.R) * Alpha)
 ColorInterpolate.G = ColA.G + ((ColB.G - ColA.G) * Alpha)
 ColorInterpolate.B = ColA.B + ((ColB.B - ColA.B) * Alpha)

End Function
Function ColorScale(ColA As ColRGB, Alpha As Single) As ColRGB

 ColorScale.R = (ColA.R * Alpha)
 ColorScale.G = (ColA.G * Alpha)
 ColorScale.B = (ColA.B * Alpha)

End Function
Function VectorAngle(VecA As Vector3D, VecB As Vector3D) As Single

 If VectorCompare3(VecA, VectorNull3) = False And VectorCompare3(VecB, VectorNull3) = False Then
  VectorAngle = VectorDotProduct3(VectorNormalize(VecA), VectorNormalize(VecB))
 End If

End Function
Function VectorCrossProduct(VecA As Vector3D, VecB As Vector3D) As Vector3D

 VectorCrossProduct.X = (VecA.Y * VecB.Z) - (VecA.Z * VecB.Y)
 VectorCrossProduct.Y = (VecA.Z * VecB.X) - (VecA.X * VecB.Z)
 VectorCrossProduct.Z = (VecA.X * VecB.Y) - (VecA.Y * VecB.X)

End Function
Function VectorDotProduct3(VecA As Vector3D, VecB As Vector3D) As Single

 VectorDotProduct3 = (VecA.X * VecB.X) + (VecA.Y * VecB.Y) + (VecA.Z * VecB.Z)

End Function
Function VectorGetNormal(VecA As Vector3D, VecB As Vector3D, VecC As Vector3D) As Vector3D

 VectorGetNormal = VectorCrossProduct(VectorSubtract3(VecA, VecB), VectorSubtract3(VecC, VecB))

End Function
Function VectorInput3(X!, Y!, Z!) As Vector3D

 VectorInput3.X = X
 VectorInput3.Y = Y
 VectorInput3.Z = Z

End Function
Function VectorLength3(Vec As Vector3D) As Single

 VectorLength3 = ((Vec.X * Vec.X) + (Vec.Y * Vec.Y) + (Vec.Z * Vec.Z)) ^ 0.5

End Function
Function VectorNormalize(Vec As Vector3D) As Vector3D

 If VectorCompare3(Vec, VectorNull3) = False Then
  VectorNormalize = VectorScale3(Vec, (1 / VectorLength3(Vec)))
 End If

End Function
Function VectorCompare3(VecA As Vector3D, VecB As Vector3D) As Boolean

 If (VecA.X = VecB.X) And (VecA.Y = VecB.Y) And (VecA.Z = VecB.Z) Then VectorCompare3 = True

End Function
Function VectorNull3() As Vector3D

End Function
Function VectorRotate(Vec As Vector3D, Axis As Byte, Angle As Single) As Vector3D

 'Basic rotations (without matrices, two more calls to Cos&Sin functions)

 Select Case Axis
  Case 0: 'X rotation, a rotation around the YZ plane
   VectorRotate.X = Vec.X
   VectorRotate.Y = (Cos(Angle) * Vec.Y) - (Sin(Angle) * Vec.Z)
   VectorRotate.Z = (Sin(Angle) * Vec.Y) + (Cos(Angle) * Vec.Z)
  Case 1: 'Y rotation, a rotation around the XZ plane
   VectorRotate.X = (Cos(Angle) * Vec.X) + (Sin(Angle) * Vec.Z)
   VectorRotate.Y = Vec.Y
   VectorRotate.Z = -(Sin(Angle) * Vec.X) + (Cos(Angle) * Vec.Z)
  Case 2: 'Z rotation, a rotation around the XY plane
   VectorRotate.X = (Cos(Angle) * Vec.X) - (Sin(Angle) * Vec.Y)
   VectorRotate.Y = (Sin(Angle) * Vec.X) + (Cos(Angle) * Vec.Y)
   VectorRotate.Z = Vec.Z
 End Select

End Function
Function VectorAdd3(VecA As Vector3D, VecB As Vector3D) As Vector3D

 VectorAdd3.X = (VecA.X + VecB.X)
 VectorAdd3.Y = (VecA.Y + VecB.Y)
 VectorAdd3.Z = (VecA.Z + VecB.Z)

End Function
Function VectorScale3(Vec As Vector3D, Alpha As Single) As Vector3D

 VectorScale3.X = (Vec.X * Alpha)
 VectorScale3.Y = (Vec.Y * Alpha)
 VectorScale3.Z = (Vec.Z * Alpha)

End Function
Function VectorSubtract3(VecA As Vector3D, VecB As Vector3D) As Vector3D

 VectorSubtract3.X = (VecA.X - VecB.X)
 VectorSubtract3.Y = (VecA.Y - VecB.Y)
 VectorSubtract3.Z = (VecA.Z - VecB.Z)

End Function
Function MatrixIdentity() As Matrix

 ' This is the 'Default' matrix, because we get:
 '   (AMatrix * IdentityMatrix) = AMatrix

 With MatrixIdentity
  .M11 = 1: .M12 = 0: .M13 = 0: .M14 = 0
  .M21 = 0: .M22 = 1: .M23 = 0: .M24 = 0
  .M31 = 0: .M32 = 0: .M33 = 1: .M34 = 0
  .M41 = 0: .M42 = 0: .M43 = 0: .M44 = 1
 End With

End Function
Function MatrixMultiply(MatA As Matrix, MatB As Matrix) As Matrix

 ' If two matrices, A & B, gives different effects,
 '  we use MatrixMultiply to give both transformations in
 '   a single matrix, in other words, matrix multiplication
 '    'Combine' two matrices.
 ' (Note well that (MatA * MatB) <> (MatB * MatA) )

 With MatrixMultiply
  .M11 = (MatA.M11 * MatB.M11) + (MatA.M21 * MatB.M12) + (MatA.M31 * MatB.M13) + (MatA.M41 * MatB.M14)
  .M12 = (MatA.M12 * MatB.M11) + (MatA.M22 * MatB.M12) + (MatA.M32 * MatB.M13) + (MatA.M42 * MatB.M14)
  .M13 = (MatA.M13 * MatB.M11) + (MatA.M23 * MatB.M12) + (MatA.M33 * MatB.M13) + (MatA.M43 * MatB.M14)
  .M14 = (MatA.M14 * MatB.M11) + (MatA.M24 * MatB.M12) + (MatA.M34 * MatB.M13) + (MatA.M44 * MatB.M14)
  .M21 = (MatA.M11 * MatB.M21) + (MatA.M21 * MatB.M22) + (MatA.M31 * MatB.M23) + (MatA.M41 * MatB.M24)
  .M22 = (MatA.M12 * MatB.M21) + (MatA.M22 * MatB.M22) + (MatA.M32 * MatB.M23) + (MatA.M42 * MatB.M24)
  .M23 = (MatA.M13 * MatB.M21) + (MatA.M23 * MatB.M22) + (MatA.M33 * MatB.M23) + (MatA.M43 * MatB.M24)
  .M24 = (MatA.M14 * MatB.M21) + (MatA.M24 * MatB.M22) + (MatA.M34 * MatB.M23) + (MatA.M44 * MatB.M24)
  .M31 = (MatA.M11 * MatB.M31) + (MatA.M21 * MatB.M32) + (MatA.M31 * MatB.M33) + (MatA.M41 * MatB.M34)
  .M32 = (MatA.M12 * MatB.M31) + (MatA.M22 * MatB.M32) + (MatA.M32 * MatB.M33) + (MatA.M42 * MatB.M34)
  .M33 = (MatA.M13 * MatB.M31) + (MatA.M23 * MatB.M32) + (MatA.M33 * MatB.M33) + (MatA.M43 * MatB.M34)
  .M34 = (MatA.M14 * MatB.M31) + (MatA.M24 * MatB.M32) + (MatA.M34 * MatB.M33) + (MatA.M44 * MatB.M34)
  .M41 = (MatA.M11 * MatB.M41) + (MatA.M21 * MatB.M42) + (MatA.M31 * MatB.M43) + (MatA.M41 * MatB.M44)
  .M42 = (MatA.M12 * MatB.M41) + (MatA.M22 * MatB.M42) + (MatA.M32 * MatB.M43) + (MatA.M42 * MatB.M44)
  .M43 = (MatA.M13 * MatB.M41) + (MatA.M23 * MatB.M42) + (MatA.M33 * MatB.M43) + (MatA.M43 * MatB.M44)
  .M44 = (MatA.M14 * MatB.M41) + (MatA.M24 * MatB.M42) + (MatA.M34 * MatB.M43) + (MatA.M44 * MatB.M44)
 End With

End Function
Function MatrixMultiplyVector3(Vec As Vector3D, Mat As Matrix) As Vector3D

 ' By giving an input vector, and a matrix,
 '  the function can map the coded transformation
 '   to the output vector by the input matrix.

 MatrixMultiplyVector3.X = (Mat.M11 * Vec.X) + (Mat.M12 * Vec.Y) + (Mat.M13 * Vec.Z) + (Mat.M14)
 MatrixMultiplyVector3.Y = (Mat.M21 * Vec.X) + (Mat.M22 * Vec.Y) + (Mat.M23 * Vec.Z) + (Mat.M24)
 MatrixMultiplyVector3.Z = (Mat.M31 * Vec.X) + (Mat.M32 * Vec.Y) + (Mat.M33 * Vec.Z) + (Mat.M34)

End Function
Function MatrixRotate(Axis As Byte, Angle As Single) As Matrix

 ' Note well that the function use only one call of
 '  Sin/Cos functions, few calculations, few memory.

 With MatrixRotate
  Select Case Axis
   Case 0: 'Rotate around X axis
    .M11 = 1
    .M22 = Cos(Angle)
    .M23 = -Sin(Angle)
    .M32 = -.M23
    .M33 = .M22
    .M44 = 1
   Case 1: 'Rotate around Y axis
    .M11 = Cos(Angle)
    .M13 = Sin(Angle)
    .M22 = 1
    .M31 = -.M13
    .M33 = .M11
    .M44 = 1
   Case 2: 'Rotate around Z axis
    .M11 = Cos(Angle)
    .M12 = -Sin(Angle)
    .M21 = -.M12
    .M22 = .M11
    .M33 = 1
    .M44 = 1
  End Select
 End With

End Function
Function MatrixScale3(Factor As Vector3D) As Matrix

 ' Note that the '1' element is set directly,
 '  instead to call MatrixIdentity function,
 '   and re-erase the memory again.

 MatrixScale3.M11 = Factor.X
 MatrixScale3.M22 = Factor.Y
 MatrixScale3.M33 = Factor.Z
 MatrixScale3.M44 = 1

End Function
Function MatrixTranslate(Distance As Vector3D) As Matrix

 ' Note that the '1' element is set directly,
 '  instead to call MatrixIdentity function,
 '   and re-erase the memory again.

 MatrixTranslate.M11 = 1
 MatrixTranslate.M14 = Distance.X
 MatrixTranslate.M22 = 1
 MatrixTranslate.M24 = Distance.Y
 MatrixTranslate.M33 = 1
 MatrixTranslate.M34 = Distance.Z
 MatrixTranslate.M44 = 1

End Function
Function MatrixView(VecFrom As Vector3D, VecLookAt As Vector3D, RollAngle As Single) As Matrix

 ' We must use 'Virtual Camera' in 3D graphics,
 '  we can specify this in three parts:
 '
 ' - Translation, Translation = -CameraTranslation
 ' - Orientation, The function can map orientations by a 'LookAt' vector
 ' - RollAngle  , This is simply the Z rotation around the screen.

 Dim N As Vector3D, U As Vector3D, V As Vector3D

 N = VectorNormalize(VectorSubtract3(VecLookAt, VecFrom))
 U = VectorNormalize(VectorCrossProduct(MatrixMultiplyVector3(VectorInput3(0, 1, 0), MatrixRotate(2, RollAngle)), N))
 V = VectorCrossProduct(N, U)

 With MatrixView
  .M11 = U.X: .M12 = U.Y: .M13 = U.Z
  .M21 = V.X: .M22 = V.Y: .M23 = V.Z
  .M31 = N.X: .M32 = N.Y: .M33 = N.Z
  .M44 = 1
 End With

 MatrixView = MatrixMultiply(MatrixTranslate(VectorInput3(-VecFrom.X, -VecFrom.Y, -VecFrom.Z)), MatrixView)

End Function
Function MatrixWorld(VecTranslate As Vector3D, VecScale As Vector3D, XPitch!, YYaw!, ZRoll!) As Matrix

 ' The world matrix is a set of a translation, a scale, and three rotations.
 '
 ' Note: We can use the MatrixView function to orient the 3D object
 '        to another, but i prefer to use directly the orientation angles.

 Dim MatTrans As Matrix, MatRotat As Matrix, MatScale As Matrix

 MatTrans = MatrixTranslate(VecTranslate)
 MatScale = MatrixScale3(VecScale)
 MatRotat = MatrixMultiply(MatrixMultiply(MatrixRotate(0, XPitch), MatrixRotate(1, YYaw)), MatrixRotate(2, ZRoll))

 MatrixWorld = MatrixMultiply(MatrixMultiply(MatScale, MatRotat), MatTrans)

End Function
