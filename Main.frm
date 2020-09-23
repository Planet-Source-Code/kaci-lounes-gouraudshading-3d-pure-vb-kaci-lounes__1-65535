VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":0000
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'####################################################################
'##                  Author: Mr KACI Lounes                        ##
'##         A 3D GouraudShading Engine in *Pure* VB Code!          ##
'##    Compile for more speed ! Mail me at KLKEANO@HOTMAIL.COM     ##
'##               Copyright © 2005 - KACI Lounes                   ##
'####################################################################

Option Explicit                   ' Stop undeclared variables

'CONSTANTS
'=========

Const FOV% = 400                  ' The Field Of View factor, or the distance
                                  '  between the eye and the projection plane
                                  '   (simply put, the focal length or
                                  '    the focal distance)
                                  '
Const NearPlane! = 0              ' Near plane clipping Zs value (0 as default)
Const FarPlane! = 500             ' Far plane clipping Zs value (500 as default)
                                  '
Const FogStart! = NearPlane       ' Fog start z value (equal to NearPlane as default)
Const FogEnd! = 300               ' Fog end z value (equal to 300 as default)
Const FogExpDensity! = 5          ' Fog density (expentional or expentional
                                  '              squared modes only)
                                  '
Const SpecularK! = 1              ' Specular power parameter K
Const SpecularN! = 5              ' Specular power parameter N
                                  ' (These parameters depends on the nature
                                  '   of the 3D object, i just prefer to put
                                  '    this as global constants)
                                  '
Const MoveSpeed! = 5              ' Player dispalcement speed
Const TurnSpeed% = 5              ' Player turning speed

'ARRAYS
'======

Dim FilesList() As String         ' A list of 3D geometry files names
                                  '
Dim SinTable(359) As Single       ' Pre-calculated Sinus(s) of 359 degrees, used to orient the player
Dim CosTable(359) As Single       ' Pre-calculated Cosinus(s) of 359 degrees, used to orient the player
                                  '
Dim Scene() As Mesh               ' 3D Models array
Dim Spots() As SpotLight3D        ' 3D Lights array
                                  '
Dim FacesDepth() As Single        ' Sort arrays : Depths array
Dim FacesIndex() As Long          '             : Faces index array
Dim MeshsIndex() As Byte          '             : Mesh index array

'CAMERA PARAMETERS
'=================

Dim CamPos As Vector3D            ' Camera position vector (in world coordinates)
Dim CamTo As Vector3D             ' Camera LookAt vector (in world coordinates)
Dim RollAngle As Single           ' Roll angle (Z rotation)
Dim ViewMatrix As Matrix          ' The view matrix
Dim ViewMode As Boolean           ' The view mode, LookAt(True), or Pitch/Yaw(False)
Dim CamAtten As Boolean           ' Camera/Lights distance attenuation flag
Dim MoveState As Integer          ' -1=Move backward, 0=No moving, 1=Move farward
Dim XAng%, YAng%                  ' Anlges rotations (Pitch/Yaw mode, in degrees)
Dim LookAtObj%                    ' Which object to look at ? (LookAt mode)

'FOG PARAMETERS
'==============

Dim FogEnable As Boolean          ' As the name indicate
Dim FogColor As ColRGB            ' As the precedent indicate
Dim FogType As Byte               '    ,,        ,,

'RASTERIZATION VARS
'==================

Dim PointA As Point2D             '
Dim PointB As Point2D             '
Dim PointC As Point2D             '
Dim TmpCol1 As ColRGB             '
Dim TmpCol2 As ColRGB             '
Dim TmpCol3 As ColRGB             '
Dim Pts() As Point2D              '
Dim Vs() As Point2D               '
Dim Ts(50) As Triangle            '
Dim Stat As Byte                  '
Dim NumPts As Byte                '
Dim NumTris As Integer            '
Dim FInx&, MInx%, S%              '

'KEYBOARD PROCESSING BOOLEANS
'============================

Dim KeyEsc As Boolean             ' Escape Key     : Exit program
Dim KeySpc As Boolean             ' Space Key      : Change view mode
Dim KeyHom As Boolean             ' Home Key       : Change X Angle (+)
Dim KeyEnd As Boolean             ' End Key        : Change X Angle (-)
Dim KeyLft As Boolean             ' Left Key       : Change Y Angle (+)
Dim KeyRgt As Boolean             ' Right Key      : Change Y Angle (-)
Dim KeyTop As Boolean             ' Up Key         : Move front
Dim KeyBot As Boolean             ' Down Key       : Move back
Dim KeyAdd As Boolean             ' NumPad + Key   : Change to look-at
Dim KeyPad1 As Boolean            ' NumPad 1 Key   : Enable/Disable Spot1
Dim KeyPad2 As Boolean            ' NumPad 2 Key   : Enable/Disable Spot2
Dim KeyPad3 As Boolean            ' NumPad 3 Key   : Enable/Disable Spot3
Dim KeyPad4 As Boolean            ' NumPad 4 Key   : Enable/Disable Spot4
Dim KeyPad5 As Boolean            ' NumPad 5 Key   : Enable/Disable Spot5

'OTHERS VARS
'===========

Dim Ambiance As ColRGB            ' The ambiant light, this is a simple additive color
Dim TotalMatrix As Matrix         ' This is the mutiplication of World/View matrices,
                                  '  because both are specified in world coordinate system,
                                  '   we can save many calculations by doing this
Dim DisplayIsChanged As Boolean   ' others !
Dim File3DName$, Buff&            ' others !
Dim I&, II&, J&, JJ&, K&, StrMsg$ ' others !
Dim InvNF!, InvFog!               ' others !
Sub LoadModels(AFilesList() As String)

 ' LoadModels
 ' ==========
 '
 ' Load the 3D structured geometry from the files list
 '  (.KLF format, my format), and setup them.

 ReDim Scene(UBound(AFilesList))

 For J = LBound(Scene) To UBound(Scene)

  'Read file & fill structures:
  Open AFilesList(J) For Binary As 1 'Open file

   Get #1, , Buff                    'Number of vertices
   ReDim Scene(J).Vertices(Buff)
   Get #1, , Buff                    'Number of faces
   ReDim Scene(J).Faces(Buff)

   'Read vertices
   For I = LBound(Scene(J).Vertices) To UBound(Scene(J).Vertices)
    Get #1, , Scene(J).Vertices(I).OrgPos.X
    Get #1, , Scene(J).Vertices(I).OrgPos.Y
    Get #1, , Scene(J).Vertices(I).OrgPos.Z
    'Set models colors:
    Select Case J
     Case 0: Scene(0).Vertices(I).OrgCol = ColorInput(50, 50, 50)   'Grey grid
     Case 1: Scene(1).Vertices(I).OrgCol = ColorInput(150, 0, 150)  'Magenta sphere
     Case 2: Scene(2).Vertices(I).OrgCol = ColorInput(0, 150, 150)  'Cyan teapot
     Case 3: Scene(3).Vertices(I).OrgCol = ColorInput(200, 100, 0)  'Orange torus
     Case 4: Scene(4).Vertices(I).OrgCol = ColorInput(0, 0, 150)    'Blue Cyborg
    End Select
   Next I

   'Read faces
   For I = LBound(Scene(J).Faces) To UBound(Scene(J).Faces)
    Get #1, , Scene(J).Faces(I).A
    Get #1, , Scene(J).Faces(I).B
    Get #1, , Scene(J).Faces(I).C
   Next I

  Close 1                            'Close the file

  'Set models reflectivity:
  Scene(J).DiffsReflct = 1
  Scene(J).SpecReflctK = SpecularK
  Scene(J).SpecReflctN = SpecularN

 Next J

 'Setup models
 '============
 'The grid should be big that others objects (as a floor):
 Scene(0).Origin = VectorInput3(0, -1, 0)
 Scene(0).Scales = VectorInput3(-2, 2, 2)
 Scene(0).Pitch = 0: Scene(0).Yaw = 0: Scene(0).Roll = 0
 'Setup the sphere:
 Scene(1).Origin = VectorInput3(20, -10, 0)
 Scene(1).Scales = VectorInput3(0.3, 0.3, 0.3)
 Scene(1).Pitch = 0: Scene(1).Yaw = 0: Scene(1).Roll = 0
 'Setup the teapot:
 Scene(2).Origin = VectorInput3(0, -5, 30)
 Scene(2).Scales = VectorInput3(0.5, 0.5, 0.5)
 Scene(2).Pitch = (Deg * 90): Scene(2).Yaw = 0: Scene(2).Roll = 0
 'Setup the torus:
 Scene(3).Origin = VectorInput3(-20, -5, -10)
 Scene(3).Scales = VectorInput3(0.3, 0.3, 0.3)
 Scene(3).Pitch = 0: Scene(3).Yaw = 0: Scene(3).Roll = 0
 'Setup the cyborg:
 Scene(4).Origin = VectorInput3(50, -20, 50)
 Scene(4).Scales = VectorInput3(20, 20, 20)
 Scene(4).Pitch = Pi: Scene(4).Yaw = 0: Scene(4).Roll = 0

End Sub
Sub MakeLights()

 ' MakeLights
 ' ==========
 '
 ' Create 5 spotlights in our beautiful scene.

 ' You can create many lights sources !!!
 '
 ' Rems:
 '
 '  - The origin vector should not be the direction vector
 '  - The origin & direction vectors should not be nulls (0, 0, 0)
 '  - The light angle limitation is 180° as maximum
 '  - The Falloff angle is less than the Hotspot angle
 '  - The Falloff & Hotspot properties is the COSINE
 '     of the angle, not the angle it-self

 'We use 5 spotlights:
 ReDim Spots(4)

 With Spots(0) 'White
  .Origin = VectorInput3(80, -80, 80)
  .Direction = VectorInput3(1, 1, 1)
  .Falloff = CosTable(5)
  .Hotspot = CosTable(15)
  .Diffuse = ColorInput(255, 255, 255)
  .Specular = ColorInput(255, 255, 255)
  .Ambiance = ColorInput(0, 0, 0)
  .Density = 0.8
  .ConstantRange = 150
  .LinearRange = 200
  .AttenEnable = True
  .Enabled = True
 End With

 With Spots(1) 'Red
  .Origin = VectorInput3(50, -50, 50)
  .Direction = VectorInput3(50, 1, 50)
  .Falloff = CosTable(15)
  .Hotspot = CosTable(45)
  .Diffuse = ColorInput(255, 0, 0)
  .Specular = ColorInput(255, 255, 255)
  .Ambiance = ColorInput(0, 0, 0)
  .Density = 1
  .ConstantRange = 150
  .LinearRange = 200
  '.AttenEnable = True
  .Enabled = True
 End With

 With Spots(2) 'Green
  .Origin = VectorInput3(-50, -50, -50)
  .Direction = VectorInput3(-50, 1, -50)
  .Falloff = CosTable(15)
  .Hotspot = CosTable(45)
  .Diffuse = ColorInput(0, 255, 0)
  .Specular = ColorInput(255, 255, 255)
  .Ambiance = ColorInput(0, 0, 0)
  .Density = 1
  .ConstantRange = 150
  .LinearRange = 200
  '.AttenEnable = True
  .Enabled = True
 End With

 With Spots(3) 'Blue
  .Origin = VectorInput3(50, -50, -50)
  .Direction = VectorInput3(50, 1, -50)
  .Falloff = CosTable(15)
  .Hotspot = CosTable(45)
  .Diffuse = ColorInput(0, 0, 255)
  .Specular = ColorInput(255, 255, 255)
  .Ambiance = ColorInput(0, 0, 0)
  .Density = 1
  .ConstantRange = 150
  .LinearRange = 200
  '.AttenEnable = True
  .Enabled = True
 End With

 With Spots(4) 'Yellow
  .Origin = VectorInput3(-50, -50, 50)
  .Direction = VectorInput3(-50, 1, 50)
  .Falloff = CosTable(15)
  .Hotspot = CosTable(45)
  .Diffuse = ColorInput(255, 255, 0)
  .Specular = ColorInput(255, 255, 255)
  .Ambiance = ColorInput(0, 0, 0)
  .Density = 1
  .ConstantRange = 150
  .LinearRange = 200
  '.AttenEnable = True
  .Enabled = True
 End With

End Sub
Sub MakeSinCosTables()

 ' MakeSinCosTables
 ' ================
 '
 ' Fill tables with the sines and the cosines of 359 (0 to 359) degrees.

 For I = 0 To 359
  SinTable(I) = Sin(Deg * I)
  CosTable(I) = Cos(Deg * I)
 Next I

End Sub
Sub Kiss(Frm As Form)

 'yeah..!, call this as Kiss Me !, not Unload Me !!

 Unload Frm: If DisplayIsChanged = True Then RestoreDisplayMode: CursorON

 StrMsg = "A 3D GouraudShading engine, Pure VB ! & FREE, but give me a little credit !" & vbNewLine & vbNewLine & _
          "                        KACI Lounes - October 2005"

 MsgBox StrMsg, vbInformation, "By !"

 End

End Sub
Function MakeStartingSure() As Boolean

 ' MakeStartingSure
 ' ================
 '
 ' This function make sure that's all the 3D parameters are
 '  corrects, if yes, True is returned as value, False otherwise.

 If (FOV < 50) Then Exit Function
 If (NearPlane < 0) Then Exit Function
 If (FarPlane < NearPlane) Then Exit Function
 If (FogStart < NearPlane) Then Exit Function
 If (FogEnd > FarPlane) Then Exit Function
 If (FogStart > FogEnd) Then Exit Function
 If (FogExpDensity < 1) Then Exit Function
 If (MoveSpeed < 1) Then Exit Function
 If (TurnSpeed < 1) Then Exit Function
 If (TurnSpeed > 45) Then Exit Function
 If VectorCompare3(CamPos, VectorNull3) = True Then Exit Function

 'Make sure lights parameters:
 For I = LBound(Spots) To UBound(Spots)
  If VectorCompare3(Spots(I).Origin, VectorNull3) = True Then Exit Function
  If VectorCompare3(Spots(I).Origin, VectorNull3) = True Then Exit Function
  If VectorCompare3(Spots(I).Origin, Spots(I).Direction) = True Then Exit Function
  If (Spots(I).ConstantRange < 0) Then Exit Function
  If (Spots(I).LinearRange < Spots(I).ConstantRange) Then Exit Function
  If (Spots(I).Falloff < 0) Then Exit Function
  If (Spots(I).Falloff > 1) Then Exit Function
  If (Spots(I).Hotspot < 0) Then Exit Function
  If (Spots(I).Hotspot > 1) Then Exit Function
  If (Spots(I).Falloff < Spots(I).Hotspot) Then Exit Function
  If (Spots(I).Density < 0) Then Exit Function
  If (Spots(I).Density > 1) Then Exit Function
 Next I

 'Correct parameters, True is returned:
 MakeStartingSure = True

End Function
Private Function Shade(FaceNormal As Vector3D, AVertex As Vertex3D, DReflect!, SReflectN!, SReflectK!) As ColRGB

 ' Shade
 ' =====
 '
 ' Shade the input vertex with all light sources,
 '  and return the shaded color 'ColRGB'.
 '
 '  This is inculd:
 '
 '   - Shape of light (spot, a cone)
 '   - Ambiant light (additif)
 '   - Diffuse reflection
 '   - Specular reflection
 '   - Light attenuation (with constant and linear ranges)
 '   - Depth-cue (Object/Camera attenuation)
 '   - Fogging (with three modes)
 '
 ' Each thing has a big-nice description!

 Dim Epsilon!, Alpha!, Beta!, Gamma!, Sigma!, Delta!
 Dim CurLight As ColRGB, ColorSum As ColRGB, LightIndx%
 Dim TLightOrigin As Vector3D, TLightDirection As Vector3D
 '(Specular calculation vars)
 Dim VAng As Single, LAng As Single
 Dim ViewDir As Vector3D, LightDir As Vector3D
 Dim ANormal As Vector3D, Reflection As Vector3D

 '=====================================================================
 '========================== SHADING LOOP =============================
 '=====================================================================

 For LightIndx = 0 To UBound(Spots)          'For each light,
  If Spots(LightIndx).Enabled = True Then    ' and only if the light is turned on,

   ' As lights are specifieds in world coordinate system,
   '  we need to transform the current light vectors (Origin & Direction),
   '   by the view matrix (in a tamporary storage):

   TLightOrigin = MatrixMultiplyVector3(Spots(LightIndx).Origin, ViewMatrix)
   TLightDirection = MatrixMultiplyVector3(Spots(LightIndx).Direction, ViewMatrix)

   ' ======================================================
   ' ====== Epsilon value define the shape of light =======
   ' ======================================================
   '
   ' Giving the 3D algorithms of Vector/Primative intersections,
   '  We can define any shape of light, cylindre, shpere, box...ect
   '
   ' Use spot light filter (a cone of light, calculation is based on angles):

   Epsilon = VectorAngle(VectorSubtract3(TLightDirection, TLightOrigin), VectorSubtract3(AVertex.TmpPos, TLightOrigin))
   If (Epsilon < 0) Then Epsilon = 0
   'Angular attenuation:
   Epsilon = (Epsilon - Spots(LightIndx).Hotspot) / (Spots(LightIndx).Falloff - Spots(LightIndx).Hotspot)
   If (Epsilon < 0) Then Epsilon = 0 Else If (Epsilon > 1) Then Epsilon = 1

   If (Epsilon <> 0) Then 'Only if we intersect the cone of light (speed up)

    ' ======================================================
    ' ======= Alpha value define the diffuse lighting ======
    ' ======================================================
    '
    ' Diffuse light: A light that's come from a knewed source,
    '                 and will go equaly in all direction.
    '
    '                Depends on normal and the light source.

    Alpha = 1 - VectorAngle(VectorSubtract3(TLightOrigin, AVertex.TmpPos), FaceNormal)
    Alpha = (Alpha * DReflect): If (Alpha < 0) Then Alpha = 0

    ' ======================================================
    ' ====== Beta value define the specular lighting =======
    ' ======================================================
    '
    ' Specular light: A light that's come from a knewed source,
    '                  and reflected exactly in the view point.
    '                 The specular reflection power is a property
    '                  of the 3D object (depends on his nature,
    '                   for example, a mirror reflect 100% light)
    '
    '                 Depends on reflection and view vectors.

    ViewDir = VectorNormalize(VectorSubtract3(CamPos, AVertex.TmpPos))
    LightDir = VectorNormalize(VectorSubtract3(TLightOrigin, AVertex.TmpPos))
    ANormal = VectorNormalize(FaceNormal)

    LAng = VectorDotProduct3(LightDir, ANormal)
    Reflection = VectorSubtract3(VectorScale3(VectorScale3(ANormal, 2), LAng), LightDir)

    VAng = VectorDotProduct3(Reflection, ViewDir)
    If (VAng > 0) Then Beta = (SReflectK * (VAng ^ SReflectN)): If (Beta < 0) Then Beta = 0

    ' ===============================================
    ' == Gamma value define the light attenuation ===
    ' ===============================================
    '
    ' The report of Gamma value is Object/Light, greater
    '  distance, less value of light (and vice-versa).

    If Spots(LightIndx).AttenEnable = True Then
     Gamma = VectorDistance3(AVertex.TmpPos, TLightOrigin)
     If (Gamma < Spots(LightIndx).ConstantRange) Then
      Gamma = 1
     ElseIf (Gamma > Spots(LightIndx).LinearRange) Then
      Gamma = 0
     Else
      Gamma = (Spots(LightIndx).LinearRange - Gamma) / (Spots(LightIndx).LinearRange - Spots(LightIndx).ConstantRange)
     End If
    Else
     Gamma = 1
    End If

    ' ======================================================
    ' ======================================================
    '
    ' Here, we have four values: Epsilon, Alpha, Beta & Gamma
    '  The next step shows how to use these values to get
    '   the vertex color.
    '
    ' For each current light, do:

    'Diffuse lighting:
    CurLight = ColorScale(ColorDiffuse(AVertex.OrgCol, ColorScale(Spots(LightIndx).Diffuse, Spots(LightIndx).Density)), Alpha)
    'Add Specular reflection:
    CurLight = ColorAdd(CurLight, ColorScale(ColorScale(Spots(LightIndx).Specular, Spots(LightIndx).Density), Beta))
    'Add the light ambiance:
    CurLight = ColorAdd(CurLight, ColorScale(Spots(LightIndx).Ambiance, Spots(LightIndx).Density))
    'Shape the light & apply Object/Light attenuation:
    CurLight = ColorScale(CurLight, (Epsilon * Gamma))

    ' ======================================================
    ' ======================================================
    '
    ' In the next line, we need to add the current light-color
    '  to the sum for applying multiple light sources:
    ColorSum = ColorAdd(ColorSum, CurLight)

   End If
  End If
 Next LightIndx

 'Set output limitations:
 Shade = ColorLimit(ColorSum)

 ' ======================================
 ' == Sigma value define the Depth-Cue ==
 ' ======================================
 '
 ' The depth-cue is simply a scale operation
 '  that's depends on the distance between
 '   the vertex and the camera, we can note
 '    this as Object/Camera attenuation.
 ' This is use NearPlane & FarPlane values.

 If CamAtten = True Then
  Sigma = (FarPlane - VectorDistance3(AVertex.TmpPos, CamPos)) * InvNF
  If (Sigma > 1) Then Sigma = 1
  Shade = ColorScale(Shade, Sigma)
 End If

 ' ======================================================
 ' ======== Delta value define the Fog effect ===========
 ' ======================================================

 'Only if the fog is enable:
 If FogEnable = True Then

  'Compute the distance between the face and the camera:
  Delta = VectorDistance3(AVertex.TmpPos, CamPos)

  'Only if we are in collision with the fog area:
  If (Delta > FogStart) And (Delta < FogEnd) Then

   Select Case FogType
    'Linear fog formula:
    Case 1: Delta = (Delta - FogStart) * InvFog
    'Expentional fog formula:
    Case 2: Delta = 1 / (NaturalLogBase ^ (FogExpDensity * ((FogEnd - Delta) * InvFog)))
    'Expentional squared fog formula:
    Case 3: Delta = 1 / (NaturalLogBase ^ (FogExpDensity * (((FogEnd - Delta) * InvFog) ^ 2)))
   End Select

   'Interpolate colors:
   Shade = ColorInterpolate(Shade, FogColor, Delta)

  End If

 End If

 ' ======================================================
 ' ================== Ambiant lighting ==================
 ' ======================================================
 '
 ' Ambient light: A light that's come everywhere,
 '                 and will go equaly in all directions.

 'Add the ambiant light:
 Shade = ColorLimit(ColorAdd(Shade, Ambiance))

End Function
Sub Process()

 ' Process
 ' =======
 '
 ' Process all the 3D here, such as transformations, shading,
 '  hidden faces removal, sorting and other stuffs.

 'World/View transform
 '====================
 For J = LBound(Scene) To UBound(Scene)
  'Make the World matrix:
  Scene(J).WorldMatrix = MatrixWorld(Scene(J).Origin, Scene(J).Scales, Scene(J).Pitch, Scene(J).Yaw, Scene(J).Roll)
  'Make the total matrix (World * View):
  TotalMatrix = MatrixMultiply(Scene(J).WorldMatrix, ViewMatrix)
  'Transform vectors by the TotalMatrix:
  For I = LBound(Scene(J).Vertices) To UBound(Scene(J).Vertices)
   Scene(J).Vertices(I).TmpPos = MatrixMultiplyVector3(Scene(J).Vertices(I).OrgPos, TotalMatrix)
   Scene(J).Vertices(I).OldShaded = 0 'Set as default (not shaded again)
  Next I
 Next J

 'Compute the face normal & Shade
 '===============================
 For J = LBound(Scene) To UBound(Scene)
  For I = LBound(Scene(J).Faces) To UBound(Scene(J).Faces)
   'Compute face normal:
   Scene(J).Faces(I).Normal = VectorGetNormal(Scene(J).Vertices(Scene(J).Faces(I).A).TmpPos, Scene(J).Vertices(Scene(J).Faces(I).B).TmpPos, Scene(J).Vertices(Scene(J).Faces(I).C).TmpPos)
   'Shade vertex A (only one):
   If (Scene(J).Vertices(Scene(J).Faces(I).A).OldShaded = 0) Then
    Scene(J).Vertices(Scene(J).Faces(I).A).ShdCol = Shade(Scene(J).Faces(I).Normal, Scene(J).Vertices(Scene(J).Faces(I).A), Scene(J).DiffsReflct, Scene(J).SpecReflctN, Scene(J).SpecReflctK)
    Scene(J).Vertices(Scene(J).Faces(I).A).OldShaded = 1
   End If
   'Shade vertex B (only one):
   If (Scene(J).Vertices(Scene(J).Faces(I).B).OldShaded = 0) Then
    Scene(J).Vertices(Scene(J).Faces(I).B).ShdCol = Shade(Scene(J).Faces(I).Normal, Scene(J).Vertices(Scene(J).Faces(I).B), Scene(J).DiffsReflct, Scene(J).SpecReflctN, Scene(J).SpecReflctK)
    Scene(J).Vertices(Scene(J).Faces(I).B).OldShaded = 1
   End If
   'Shade vertex C (only one):
   If (Scene(J).Vertices(Scene(J).Faces(I).C).OldShaded = 0) Then
    Scene(J).Vertices(Scene(J).Faces(I).C).ShdCol = Shade(Scene(J).Faces(I).Normal, Scene(J).Vertices(Scene(J).Faces(I).C), Scene(J).DiffsReflct, Scene(J).SpecReflctN, Scene(J).SpecReflctK)
    Scene(J).Vertices(Scene(J).Faces(I).C).OldShaded = 1
   End If
  Next I
 Next J

 'Projection
 '==========
 For J = LBound(Scene) To UBound(Scene)
  For I = LBound(Scene(J).Vertices) To UBound(Scene(J).Vertices)

   ' Persective projection (can change the FOV)
   '
   ' In computer graphics, the familiar perspective projection
   '  is done by the calculation: X = (X/Z), Y = (Y/Z)
   ' But because that the distance between the eye and
   '  the projection plane is null, we need to scale the X,Y
   '  coordinates by a distance: the focal length (constant).
   '
   ' So X = ((X/Z) * FOV), Y = ((Y/Z) * FOV)
   '
   ' Note that the FOV (Field Of View) factor, is
   '  exprrimed as distance, not an angle.
   '
   ' I completly ignore the perspective transformation
   '  that is exprimed in a matrix, i supose that is
   '   mush simple and faster.
   '
   ' Ignore the division by zero, by replacing it by ApproachVal constant:
   If Scene(J).Vertices(I).TmpPos.Z = 0 Then Scene(J).Vertices(I).TmpPos.Z = ApproachVal
   '
   ' Apply persective distortion:
   Scene(J).Vertices(I).TmpPos.X = (Scene(J).Vertices(I).TmpPos.X / Scene(J).Vertices(I).TmpPos.Z) * FOV
   Scene(J).Vertices(I).TmpPos.Y = (Scene(J).Vertices(I).TmpPos.Y / Scene(J).Vertices(I).TmpPos.Z) * FOV

  Next I
 Next J

 'Hidden faces removal
 '====================
 '
 ' 1. Check the visiblity of faces by the face normal (back-face culling)
 ' 2. Check if the triangle is between FarPlane & NearPlane

 For J = LBound(Scene) To UBound(Scene)
  For I = LBound(Scene(J).Faces) To UBound(Scene(J).Faces)

   'Recalculate face normal (or multiply it by the view matrix):
   Scene(J).Faces(I).Normal = VectorGetNormal(Scene(J).Vertices(Scene(J).Faces(I).A).TmpPos, Scene(J).Vertices(Scene(J).Faces(I).B).TmpPos, Scene(J).Vertices(Scene(J).Faces(I).C).TmpPos)

   If (Scene(J).Faces(I).Normal.Z > 0) Then
    '--------------------------------------
    If (Scene(J).Vertices(Scene(J).Faces(I).A).TmpPos.Z > NearPlane) And _
       (Scene(J).Vertices(Scene(J).Faces(I).B).TmpPos.Z > NearPlane) And _
       (Scene(J).Vertices(Scene(J).Faces(I).C).TmpPos.Z > NearPlane) Then
     '-------------------------------------------------------
     If (Scene(J).Vertices(Scene(J).Faces(I).A).TmpPos.Z < FarPlane) And _
        (Scene(J).Vertices(Scene(J).Faces(I).B).TmpPos.Z < FarPlane) And _
        (Scene(J).Vertices(Scene(J).Faces(I).C).TmpPos.Z < FarPlane) Then

      'Add the averaged depth of face (Zs/3) to FacesDepths array:
      ReDim Preserve FacesDepth(UBound(FacesDepth) + 1)
      FacesDepth(UBound(FacesDepth)) = (Scene(J).Vertices(Scene(J).Faces(I).A).TmpPos.Z + _
                                        Scene(J).Vertices(Scene(J).Faces(I).B).TmpPos.Z + _
                                        Scene(J).Vertices(Scene(J).Faces(I).C).TmpPos.Z) * OneByThree 'Ignore the division by 3 (constant),
                                                                                                      ' by a multiplication by (1/3)
      'Add the face index to the FacesIndex array:
      ReDim Preserve FacesIndex(UBound(FacesIndex) + 1)
      FacesIndex(UBound(FacesIndex)) = I

      'Add the mesh index to the MeshsIndexs array:
      ReDim Preserve MeshsIndex(UBound(MeshsIndex) + 1)
      MeshsIndex(UBound(MeshsIndex)) = J

     End If
    End If
   End If

  Next I
 Next J

 'Sort back to front
 '==================
 ExtractSort3D FacesDepth(), FacesIndex(), MeshsIndex()

End Sub
Sub GetKeys()

 ' GetKeys
 ' =======
 '
 ' Process keyboard entries.

 'Escape key, exit program:
 If KeyEsc = True Then Kiss Me

 'Space key, change the view mode:
 If KeySpc = True Then ViewMode = Not ViewMode

 'Numpad Add key, change the object to look-at (LookAt mode only):
 If KeyAdd = True Then
  If LookAtObj = UBound(Scene) Then
   LookAtObj = 0
  Else
   LookAtObj = (LookAtObj + 1)
  End If
 End If

 'Use the following keys to change the camera orientation:
 '
 'Home key, +Y
 If KeyRgt = True Then
  YAng = (YAng + TurnSpeed): If (YAng > 359) Then YAng = (YAng - 359)
 End If
 'End key, -Y
 If KeyLft = True Then
  YAng = (YAng - TurnSpeed): If (YAng < 0) Then YAng = (359 + YAng)
 End If
 'Right key, +X
 If KeyHom = True Then
  XAng = (XAng + TurnSpeed): If (XAng > 359) Then XAng = (XAng - 359)
 End If
 'Left key, -X
 If KeyEnd = True Then
  XAng = (XAng - TurnSpeed): If (XAng < 0) Then XAng = (359 + XAng)
 End If

 'Use the following keys to move farward/backward:
 If KeyBot = True Then MoveState = -1
 If KeyTop = False And KeyBot = False Then MoveState = 0
 If KeyTop = True Then MoveState = 1

 'Numpad numeric keys, Enable/Disable lights:
 If KeyPad1 = True Then Spots(0).Enabled = Not Spots(0).Enabled
 If KeyPad2 = True Then Spots(1).Enabled = Not Spots(1).Enabled
 If KeyPad3 = True Then Spots(2).Enabled = Not Spots(2).Enabled
 If KeyPad4 = True Then Spots(3).Enabled = Not Spots(3).Enabled
 If KeyPad5 = True Then Spots(4).Enabled = Not Spots(4).Enabled

 'Reset defaults (disable all):
 KeySpc = False: KeyTop = False: KeyBot = False: KeyAdd = False
 KeyHom = False: KeyEnd = False: KeyRgt = False: KeyLft = False
 KeyPad1 = False: KeyPad2 = False: KeyPad3 = False: KeyPad4 = False: KeyPad5 = False

End Sub
Sub LoadScene()

 ' LoadScene
 ' =========
 '
 ' The staring function, load the 3D from files, create lights,
 '  setup the camera and the fog, build Sin/Cos tables...ect

 'Make the Sin/Cos tables
 '=======================
 MakeSinCosTables

 'Load models
 '===========
 ReDim FilesList(4)
 FilesList(0) = App.Path & "\Primatives\Grid.klf"
 FilesList(1) = App.Path & "\Primatives\Sphere.klf"
 FilesList(2) = App.Path & "\Primatives\Teapot.klf"
 FilesList(3) = App.Path & "\Primatives\Torus.klf"
 FilesList(4) = App.Path & "\Primatives\Cyborg.klf"
 LoadModels FilesList()

 'Setup the camera
 '================
 ViewMatrix = MatrixIdentity                'Set view matrix as identity
 CamPos = VectorInput3(-10, -50, -75)       'Set Camera position
 'RollAngle = (Deg * 15)                    'The camera roll angle (Z rotation)
 ViewMode = False: CamAtten = False         '
 XAng = 15: YAng = 5: LookAtObj = 4         '

 'Make spot-lights
 '================
 MakeLights

 'Set the ambiant light
 '=====================
 Ambiance = ColorInput(10, 10, 10)

 'Ask the user if he want to enable fogging mode
 '==============================================
 StrMsg = "Enable fogging ?"
 If MsgBox(StrMsg, vbYesNo + vbQuestion, "Fogging") = vbYes Then
  FogEnable = True: InvFog = (1 / (FogEnd - FogStart))

  StrMsg = "Select the Fog type:" & vbNewLine & vbNewLine & _
           "1. Linear" & vbNewLine & _
           "2. Expentional" & vbNewLine & _
           "3. Expentional squared"

  Select Case InputBox(StrMsg, "Fog type", "1")
   Case "1": FogType = 1  'Linear Fog
   Case "2": FogType = 2  'Expentional Fog
   Case "3": FogType = 3  'Expentional squared Fog
   Case Else: FogType = 0 'Default (Linear Fog)
  End Select
 End If

 'Setup the fog color
 '===================
 FogColor = ColorInput(255, 255, 255) 'White fog

 'Initialize sort arrays
 '======================
 ReDim FacesDepth(0)
 ReDim FacesIndex(0)
 ReDim MeshsIndex(0)

 'Make sure the previous parameters:
 If MakeStartingSure = False Then
  StrMsg = "Please setup all the parameters correctly and reset again !"
  MsgBox StrMsg, vbCritical, "False 3D parameters !"
  Kiss Me
 End If

 '(Used for shading calculations)
 InvNF = 1 / (FarPlane - NearPlane)
 InvFog = 1 / (FogEnd - FogStart)

 StrMsg = "Controls" & vbNewLine
 StrMsg = StrMsg & "=====" & vbNewLine & vbNewLine
 StrMsg = StrMsg & "Escape   Key : Exit program" & vbNewLine
 StrMsg = StrMsg & "Space    Key : Change view mode" & vbNewLine
 StrMsg = StrMsg & "Home     Key : Change X Angle (+)" & vbNewLine
 StrMsg = StrMsg & "End      Key : Change X Angle (-)" & vbNewLine
 StrMsg = StrMsg & "Left     Key : Change Y Angle (+)" & vbNewLine
 StrMsg = StrMsg & "Right    Key : Change Y Angle (-)" & vbNewLine
 StrMsg = StrMsg & "Up       Key : Move front" & vbNewLine
 StrMsg = StrMsg & "Down     Key : Move back" & vbNewLine
 StrMsg = StrMsg & "NumPad + Key : Change to look-at" & vbNewLine
 StrMsg = StrMsg & "NumPad 1 Key : Enable/Disable Spot1" & vbNewLine
 StrMsg = StrMsg & "NumPad 2 Key : Enable/Disable Spot2" & vbNewLine
 StrMsg = StrMsg & "NumPad 3 Key : Enable/Disable Spot3" & vbNewLine
 StrMsg = StrMsg & "NumPad 4 Key : Enable/Disable Spot4" & vbNewLine
 StrMsg = StrMsg & "NumPad 5 Key : Enable/Disable Spot5"

 'Show a message box for controls:
 MsgBox StrMsg, vbInformation, "Controls"

End Sub
Sub Render()

 ' Render
 ' ======
 '
 ' Rasterization (2D)
 '
 ' Clip the faces (only if necessary) and draw them with
 '  ShdCol (shaded color), note that before rasterization,
 '   we split the clipped polygon into smalls triangles
 '   (triangulation, and only if clipped)
 '
 ' I see that is not a best way for clipping,
 '  because that is a 2D clipping, the proccess
 '   clip the projected triangles in 2D space.
 ' I want to add the 3D clipping (in a 3D box)
 '  in futur versions !!

 For I = LBound(FacesIndex) To UBound(FacesIndex)

  MInx = MeshsIndex(I): FInx = FacesIndex(I)

  PointA.X = (340 + Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).TmpPos.X)
  PointA.Y = (260 + Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).TmpPos.Y)
  PointB.X = (340 + Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).TmpPos.X)
  PointB.Y = (260 + Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).TmpPos.Y)
  PointC.X = (340 + Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).TmpPos.X)
  PointC.Y = (260 + Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).TmpPos.Y)

  ClipTriangle2D 20, 20, 660, 500, CSng(PointA.X), CSng(PointA.Y), CSng(PointB.X), CSng(PointB.Y), CSng(PointC.X), CSng(PointC.Y), Pts(), NumPts, Stat

  Select Case Stat

   Case 1: 'Fully triangle, so draw it normaly:

    DrawGradientTriangle hdc, PointA, PointB, PointC, _
                         Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol, _
                         Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol, _
                         Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol

    ' Very impotant note:
    '
    ' If you want to run the flat-shading version,
    '  we interpolate the three vertices colors,
    '   using barycentric coordinates (U=V=W = OneByThree constant) suh that:
    '
    ' FlatColor.R = (ColA.R * OneByThree) + (ColB.R * OneByThree) + (ColC.R * OneByThree)
    ' FlatColor.G = (ColA.G * OneByThree) + (ColB.G * OneByThree) + (ColC.G * OneByThree)
    ' FlatColor.B = (ColA.B * OneByThree) + (ColB.B * OneByThree) + (ColC.B * OneByThree)
    '
    ' Draw flat triangle using FlatColor as color (similar for clipped triangles).

    '===============================================================

    ' For a simple preview for the flat-shading version:
    ' (please disable the old line in the top named
    '  'DrawGradientTriangle...' and enable the next
    '  8 lines of code)

    'TmpCol1 = Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol
    'TmpCol2 = Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol
    'TmpCol3 = Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol

    'Dim FlatColor As ColRGB
    'FlatColor.R = (TmpCol1.R * OneByThree) + (TmpCol2.R * OneByThree) + (TmpCol3.R * OneByThree)
    'FlatColor.G = (TmpCol1.G * OneByThree) + (TmpCol2.G * OneByThree) + (TmpCol3.G * OneByThree)
    'FlatColor.B = (TmpCol1.B * OneByThree) + (TmpCol2.B * OneByThree) + (TmpCol3.B * OneByThree)

    'DrawGradientTriangle hdc, PointA, PointB, PointC, FlatColor, FlatColor, FlatColor

    ' Remember, this is just a simple preview.

    '===============================================================

   Case 2: 'Polygon is clipped & triangulated:

    '----------------- Triangulation part ----------------
    ReDim Vs(1)
    For S = 0 To NumPts
     If (S <> 0) Then ReDim Preserve Vs(UBound(Vs) + 1)
     Vs(UBound(Vs)).X = Pts(S).X
     Vs(UBound(Vs)).Y = Pts(S).Y
    Next S
    ReDim Preserve Vs(UBound(Vs) + 5)

    'Triangulate the new polygon:
    NumTris = Triangulate(CInt(UBound(Vs) - 5), Vs(), Ts())
    '------------------------------------------------------

    For S = 1 To NumTris 'For each segment triangle:

     'Find (interpolate) the pre-vertex A color:
     TmpCol1.R = BaryInterpolateLinear(PointA.X, PointA.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol.R, _
                                       PointB.X, PointB.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol.R, _
                                       PointC.X, PointC.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol.R, _
                                       Vs(Ts(S).A).X, Vs(Ts(S).A).Y)
     TmpCol1.G = BaryInterpolateLinear(PointA.X, PointA.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol.G, _
                                       PointB.X, PointB.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol.G, _
                                       PointC.X, PointC.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol.G, _
                                       Vs(Ts(S).A).X, Vs(Ts(S).A).Y)
     TmpCol1.B = BaryInterpolateLinear(PointA.X, PointA.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol.B, _
                                       PointB.X, PointB.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol.B, _
                                       PointC.X, PointC.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol.B, _
                                       Vs(Ts(S).A).X, Vs(Ts(S).A).Y)

     'Find (interpolate) the pre-vertex B color:
     TmpCol2.R = BaryInterpolateLinear(PointA.X, PointA.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol.R, _
                                       PointB.X, PointB.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol.R, _
                                       PointC.X, PointC.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol.R, _
                                       Vs(Ts(S).B).X, Vs(Ts(S).B).Y)
     TmpCol2.G = BaryInterpolateLinear(PointA.X, PointA.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol.G, _
                                       PointB.X, PointB.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol.G, _
                                       PointC.X, PointC.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol.G, _
                                       Vs(Ts(S).B).X, Vs(Ts(S).B).Y)
     TmpCol2.B = BaryInterpolateLinear(PointA.X, PointA.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol.B, _
                                       PointB.X, PointB.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol.B, _
                                       PointC.X, PointC.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol.B, _
                                       Vs(Ts(S).B).X, Vs(Ts(S).B).Y)

     'Find (interpolate) the pre-vertex C color:
     TmpCol3.R = BaryInterpolateLinear(PointA.X, PointA.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol.R, _
                                       PointB.X, PointB.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol.R, _
                                       PointC.X, PointC.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol.R, _
                                       Vs(Ts(S).C).X, Vs(Ts(S).C).Y)
     TmpCol3.G = BaryInterpolateLinear(PointA.X, PointA.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol.G, _
                                       PointB.X, PointB.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol.G, _
                                       PointC.X, PointC.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol.G, _
                                       Vs(Ts(S).C).X, Vs(Ts(S).C).Y)
     TmpCol3.B = BaryInterpolateLinear(PointA.X, PointA.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).A).ShdCol.B, _
                                       PointB.X, PointB.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).B).ShdCol.B, _
                                       PointC.X, PointC.Y, Scene(MInx).Vertices(Scene(MInx).Faces(FInx).C).ShdCol.B, _
                                       Vs(Ts(S).C).X, Vs(Ts(S).C).Y)

     'Draw the segement triangle:
     DrawGradientTriangle hdc, Vs(Ts(S).A), Vs(Ts(S).B), Vs(Ts(S).C), ColorLimit(TmpCol1), ColorLimit(TmpCol2), ColorLimit(TmpCol3)

    Next S
    ReDim Vs(1)

   'Case Else: The triangle is outside the screen, skip rasterization.

  End Select

 Next I

 'After rasterization, clear sort arrays:
 ReDim FacesDepth(0): ReDim FacesIndex(0): ReDim MeshsIndex(0)

End Sub
Private Sub Form_Activate()

 '=====================================================================
 '============================ MAIN LOOP ==============================
 '=====================================================================

 Do
  Cls

  '==========================================================
  'Rotate 5° our white spot around XZ plane (Y axis):
  Spots(0).Origin = VectorRotate(Spots(0).Origin, 1, Deg * 20)
  '
  'Enable this line to view the white light position:
  'Scene(1).Origin = Spots(0).Origin
  'Rotate teapot:
  Scene(2).Yaw = Scene(2).Yaw + (Deg * 5)
  'Rotate cyborg:
  Scene(4).Origin = VectorRotate(Scene(4).Origin, 1, Deg * 15)
  Scene(4).Yaw = Scene(4).Yaw + (Deg * 35)
  '==========================================================

  GetKeys

  '==========================================================

  If ViewMode = True Then

   'LookAt mode at LookAtObj%:

   CamTo.Y = CosTable(XAng) - SinTable(XAng)
   CamTo.Z = SinTable(XAng) + CosTable(XAng)
   CamTo.X = SinTable(YAng) * CamTo.Z
   CamTo.Z = -(SinTable(YAng) * CamTo.X) + (CosTable(YAng) * CamTo.Z)

   'Scale it by the MoveSpeed factor:
   CamTo = VectorScale3(CamTo, MoveSpeed)

   Select Case MoveState
    Case -1: CamPos = VectorSubtract3(CamPos, CamTo) 'Move back
    Case 0: 'NOT MOVING !
    Case 1: CamPos = VectorAdd3(CamPos, CamTo) 'Move front
   End Select

   CamTo = Scene(LookAtObj).Origin

  Else '======================================================

   'Pitch/Yaw mode:
   'Convert the spherial angles to cartesian coordinates:
   CamTo.Y = CosTable(XAng) - SinTable(XAng)
   CamTo.Z = SinTable(XAng) + CosTable(XAng)
   CamTo.X = SinTable(YAng) * CamTo.Z
   CamTo.Z = -(SinTable(YAng) * CamTo.X) + (CosTable(YAng) * CamTo.Z)

   'Scale it by the MoveSpeed factor:
   CamTo = VectorScale3(CamTo, MoveSpeed)

   Select Case MoveState
    Case -1: CamPos = VectorSubtract3(CamPos, CamTo) 'Move back
    Case 0: 'NOT MOVING !
    Case 1: CamPos = VectorAdd3(CamPos, CamTo) 'Move front
   End Select

   CamTo = VectorAdd3(CamPos, CamTo)

  End If

  '(how the white spotlight view the scene ??)
  '
  'Answer:
  'CamPos = Spots(0).Origin: CamTo = Spots(0).Direction

  '==========================================================

  'Set transform: view matrix
  ViewMatrix = MatrixView(CamPos, CamTo, RollAngle)

  Process
  Render

  DoEvents
 Loop

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 ' Form_KeyDown
 ' ============
 '
 ' Generate the keyboard

 If KeyCode = vbKeyEscape Then KeyEsc = True
 If KeyCode = vbKeySpace Then KeySpc = True
 If KeyCode = vbKeyLeft Then KeyLft = True
 If KeyCode = vbKeyRight Then KeyRgt = True
 If KeyCode = vbKeyUp Then KeyTop = True
 If KeyCode = vbKeyDown Then KeyBot = True
 If KeyCode = vbKeyHome Then KeyHom = True
 If KeyCode = vbKeyEnd Then KeyEnd = True
 If KeyCode = vbKeyAdd Then KeyAdd = True
 If KeyCode = vbKeyNumpad1 Then KeyPad1 = True
 If KeyCode = vbKeyNumpad2 Then KeyPad2 = True
 If KeyCode = vbKeyNumpad3 Then KeyPad3 = True
 If KeyCode = vbKeyNumpad4 Then KeyPad4 = True
 If KeyCode = vbKeyNumpad5 Then KeyPad5 = True

End Sub
Private Sub Form_Load()

 ' From_Load
 ' =========

 ' Redim our window as (640x480), we use 680x520
 '  resolution for showing the clipping process
 '   as (20, 20, 660, 500).
 Move 0, 0, (680 * 15), (520 * 15)

 'Ask the user if he wants to change the screen resolution:
 If MsgBox("Change the display settings (640x480x16) ?", vbYesNo + vbQuestion, "Resolution") = vbYes Then
  DisplayIsChanged = True: CursorOFF
  ChangeDisplayMode 640, 480, 16
 End If

 'Load all the 3D:
 LoadScene

End Sub
