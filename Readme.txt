 3D GouraudShading in pure VB!, by KACI Lounes October/2005
 ==========================================================

 Files included
 ==============

 FORMS      : Main.FRM

 MODULES    : Clipper.BAS
              ExtractSort.BAS
              Maths3D.BAS
              Rasterizations.BAS
              Triangulation.BAS

 CLASSES    : No

 ACTIVEX    : No

 RES-FILE   : No

 REFERENCES : No

 SNAPSHOT   : Yes

 OTHERS     : 3D datas files (.KLF)

 ==========================================

 Description
 ===========

 But NO, you don't dream !!

 This is the V4 of my projects : a 3D GouraudShading engine fully in VB!
  no references, no DLLs, no OCXs, just pure VB code. (see screen shot)

 As the screen shot, this is a 3D scene (A grid, sphere, teapot, torus & a cyborg!)

 - The scene is lighted with spotlights, I describe some useful
    lighting members (Ambiance, Diffusion, Specular, Attenuations)

 - Also i add the fog lighting, with three modes (Linear, Exp, Exp2)

 - The program use matrices calculations for transformations (world, camera...ect).

 - The sorting methode is the simple Painter's algorithm, i prefer
    to use a simple sort technique: 'ExtractionSort'
   I find that is very simple and faster that 'QuickSort',
    to change the sorting mode (ascending or desecending), we
     simply flip the sign: < to > (see the code)

    I also find that is easy to be extanted in 3D to sort back to front.

 - The clipping is 2D, so all the triangles are planars,
    or the clipping process is done on this 'projected' triangles.
   (i add the 3D clipping version in the futur !)

 - You have two mode for viewing the scene:
    1. Pitch/Yaw mode: You can control the position and the orientation
    2. LookAt mode: You can control only the position

 Controls
 ========

 Escape   Key : Exit program
 Space    Key : Change view mode
 Home     Key : Change X Angle (+)
 End      Key : Change X Angle (-)
 Left     Key : Change Y Angle (+)
 Right    Key : Change Y Angle (-)
 Up       Key : Move front
 Down     Key : Move back
 NumPad + Key : Change to look-at
 NumPad 1 Key : Enable/Disable Spot1
 NumPad 2 Key : Enable/Disable Spot2
 NumPad 3 Key : Enable/Disable Spot3
 NumPad 4 Key : Enable/Disable Spot4
 NumPad 5 Key : Enable/Disable Spot5

 Requiements
 ===========

  This program use as a compiled EXE:

   - DISK SPACE : 148 Kb

   - RAM  SPACE : 4 800 Kb (4,8 Mb)
    *  62,3 Kb of the 3D datas
    * 900,0 Kb for the background
   so only the program use: (4800 - (62,3 + 900)) = 3 837,7 Kb

 THE CORE OF THE ENGINE
 ======================

 [LOADING PART]

  - Make Sin/Cos tables .................................. ===== 'LoadScene'  Procedure =====
  - Load models from files, .............................. ===== 'LoadModels' Procedure =====
     & setup theme (3D transforms, colors, reflectivity)
  - Setup the camera ..................................... ===== 'LoadScene'  Procedure =====
  - Make 5 spot-lights ................................... ===== 'MakeLights' Procedure =====
  - Set the ambiant light color .......................... ==================================
  - Enable/Disable fogging and the fog mode (MsgBox) ..... ===== 'LoadScene'  Procedure =====
  - Set the fog color .................................... ==================================
  - Initialize sort arrays ............................... ==================================
  - [Make sure the previous parameters] .................. == 'MakeStartingSure' Procedure ==

 [MAIN LOOP]

  Set world transform parameters for 3D objects
  Generate keyboard input ................................ ====== 'GetKeys' Procedure ======
  Set view transform: make the view matrix

  Make the world matrix for the current model ............ =================================
  Calculate the total matrix (World * View) .............. ====== 'Process' Procedure ======
  Transform the current model by the total matrix ........ =================================

  [SHADING]

   - Compute face normal ................................. ====== 'Process' Procedure ======
   - Shade VertexA, VertexB & VertexC .................... ====== 'Shade'   Procedure ======

  [PROJECTION]

   - Perspective distortion .............................. ====== 'Process' Procedure ======

  [HIDDEN FACE REMOVAL]

   - Check visibility by face normal, if yes: ............ =================================
    - Face should be between Near & Far planes, if yes: .. =================================
     - Add the averaged depth of face to FacesDepth array. ====== 'Process' Procedure ======
     - Add the face index to FacesIndex array ............ =================================
     - Add the mesh index to MeshsIndex array ............ =================================

  [SORTING]

   - Sort faces back to front ............................ ====== 'Process' Procedure ======

  [RASTERIZATION]

   Clip projected faces only if necessary, if yes: ....... =================================
    - Triangulate polygons ............................... =================================
    - Compute the pre-vertices colors .................... ====== 'Render'  Procedure ======
    - Draw faces ......................................... =================================
    - Clear sort arrays .................................. =================================

 [End Main Loop]

 ========================

 Optimizations
 =============

 Note that in this program, we use only single
  precision or reals numbers (32 Bits : the Single data type).

 Note that the types of variables are just for the use,
  for exemple i need number form 2 to 120, then, the variable is 'Byte'
   data type (0 to 255), then these small optimizations reduce the
    size of the used memory.

 Some other code optimizations:

  - A = (-1 * A)                 =======>   A = -A
  - If (A - B) = 0 Then...       =======>   If (A = B) Then...
  - A=(A/X), B=(B/X), C=(C/X)    =======>   XX=(1/X), A=(A*XX), B=(B*XX), C=(C*XX)

 These smalls optimizations save more CPU processing.....
  Also that i prefer to use arithmetic as maximum.....

 The code is written of a way that it is legible and understanding,
  and the 3D programming is clearly classified, so you can see everything nicely.

 Other previous submissions on the PSC
 =====================================

  - World 3D (point rendering)
  - SpotLight 3D
  - WireFrame 3D
  - FlatShading 3D

  All in pure VB !

 About the author
 ================

 Full name: KACI Lounes
 Year Born: 08/10/1988
 Country  : Réghaia, Alger-centre (Algeria)

 I interest to Computer Graphics/Games developement.

 If you want to find my others previous submissions,
  simply type as cretaria: KACI Lounes

 Sorry for the orthographic errors (!), I don't speak 100% english !
  (If you speak french, it's will be 100% compatible !!!)

 Contact
 =======

 If you have any questions about my projects, Additions or any
  suggestions, please contact me at:
 
 klkeano@Caramail.com

 Copyright © 11/2005 - KACI Lounes

 EOF.