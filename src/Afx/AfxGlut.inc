' ########################################################################################
' Microsoft Windows
' File: AfxGlut.inc
' Contents: OpenGL utilities.
' Adaptation to Free Basic of freeglut geometric object rendering procedures.
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "GL/windows/glu.bi"

' Freeglut 2.8.1 includes twenty two routines for generating easily-recognizable 3-d geometric objects.
' These routines are effectively the same ones that are included in the GLUT library, and
' reflect the functionality available in the aux toolkit described in the OpenGL Programmer's
' Guide. They are included to allow programmers to create with a single line of code a
' three-dimensional object which can be used to test a variety of OpenGL functionality.
' None of the routines generates a display list for the object which it draws. The functions
' generate normals appropriate for lighting but, except for the teapot functions, do not generate
' texture coordinates. Do note that depth testing (GL_LESS) should be enabled for the correct
' drawing of the nonconvex objects, i.e., the glutTorus and glutTeapot.

' ========================================================================================
' * Renders a teapot.
' ========================================================================================
SUB AfxGlutTeapot (BYVAL grid AS LONG, BYVAL dbScale AS DOUBLE, BYVAL dwType AS DWORD)

   ' // Rim, body, lid, and bottom data must be rotated along all four quadrants;
   ' // handle and spout data is flipped across the x-y plane (negate z values) only.
   STATIC patchdata(9, 15) AS LONG = { _
    { 102, 103, 104, 105,   4,   5,   6,   7,   8,   9,  10,  11,  12,  13,  14,  15 }, _ ' /* rim    */
    {  12,  13,  14,  15,  16,  17,  18,  19,  20,  21,  22,  23,  24,  25,  26,  27 }, _ ' /* body   */
    {  24,  25,  26,  27,  29,  30,  31,  32,  33,  34,  35,  36,  37,  38,  39,  40 }, _
    {  96,  96,  96,  96,  97,  98,  99, 100, 101, 101, 101, 101,   0,   1,   2,   3 }, _ ' /* lid    */
    {   0,   1,   2,   3, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117 }, _
    { 118, 118, 118, 118, 124, 122, 119, 121, 123, 126, 125, 120,  40,  39,  38,  37 }, _ ' /* bottom */
    {  41,  42,  43,  44,  45,  46,  47,  48,  49,  50,  51,  52,  53,  54,  55,  56 }, _ ' /* handle */
    {  53,  54,  55,  56,  57,  58,  59,  60,  61,  62,  63,  64,  28,  65,  66,  67 }, _
    {  68,  69,  70,  71,  72,  73,  74,  75,  76,  77,  78,  79,  80,  81,  82,  83 }, _ ' /* spout  */
    {  80,  81,  82,  83,  84,  85,  86,  87,  88,  89,  90,  91,  92,  93,  94,  95 } _
    }

   STATIC cpdata(126, 2) AS DOUBLE = { _
    {0.2, 0, 2.7},             {0.2, -0.112, 2.7},        {0.112, -0.2, 2.7}, _
    {0,-0.2, 2.7},             {1.3375, 0, 2.53125},      {1.3375, -0.749, 2.53125}, _
    {0.749, -1.3375, 2.53125}, {0, -1.3375, 2.53125},     {1.4375, 0, 2.53125}, _
    {1.4375, -0.805, 2.53125}, {0.805, -1.4375, 2.53125}, {0, -1.4375, 2.53125}, _
    {1.5, 0, 2.4},             {1.5, -0.84, 2.4},         {0.84, -1.5, 2.4}, _
    {0, -1.5, 2.4},            {1.75, 0, 1.875},          {1.75, -0.98, 1.875}, _
    {0.98, -1.75, 1.875},      {0, -1.75, 1.875},         {2, 0, 1.35}, _
    {2, -1.12, 1.35},          {1.12, -2, 1.35},          {0, -2, 1.35}, _
    {2, 0, 0.9},               {2, -1.12, 0.9},           {1.12, -2, 0.9}, _
    {0, -2, 0.9},              {-2, 0, 0.9},              {2, 0, 0.45}, _
    {2, -1.12, 0.45},          {1.12, -2, 0.45},          {0, -2, 0.45}, _
    {1.5, 0, 0.225},           {1.5, -0.84, 0.225},       {0.84, -1.5, 0.225}, _
    {0, -1.5, 0.225},          {1.5, 0, 0.15},            {1.5, -0.84, 0.15}, _
    {0.84, -1.5, 0.15},        {0, -1.5, 0.15},           {-1.6, 0, 2.025}, _
    {-1.6, -0.3, 2.025},       {-1.5, -0.3, 2.25},        {-1.5, 0, 2.25}, _
    {-2.3, 0, 2.025},          {-2.3, -0.3, 2.025},       {-2.5, -0.3, 2.25}, _
    {-2.5, 0, 2.25},           {-2.7, 0, 2.025},          {-2.7, -0.3, 2.025}, _
    {-3, -0.3, 2.25},          {-3, 0, 2.25},             {-2.7, 0, 1.8}, _
    {-2.7, -0.3, 1.8},         {-3, -0.3, 1.8},           {-3, 0, 1.8}, _
    {-2.7, 0, 1.575},          {-2.7, -0.3, 1.575},       {-3, -0.3, 1.35}, _
    {-3, 0, 1.35},             {-2.5, 0, 1.125},          {-2.5, -0.3, 1.125}, _
    {-2.65, -0.3, 0.9375},     {-2.65, 0, 0.9375},        {-2, -0.3, 0.9}, _
    {-1.9, -0.3, 0.6},         {-1.9, 0, 0.6},            {1.7, 0, 1.425}, _
    {1.7, -0.66, 1.425},       {1.7, -0.66, 0.6},         {1.7, 0, 0.6}, _
    {2.6, 0, 1.425},           {2.6, -0.66, 1.425},       {3.1, -0.66, 0.825}, _
    {3.1, 0, 0.825},           {2.3, 0, 2.1},             {2.3, -0.25, 2.1}, _
    {2.4, -0.25, 2.025},       {2.4, 0, 2.025},           {2.7, 0, 2.4}, _
    {2.7, -0.25, 2.4},         {3.3, -0.25, 2.4},         {3.3, 0, 2.4}, _
    {2.8, 0, 2.475},           {2.8, -0.25, 2.475},       {3.525, -0.25, 2.49375}, _
    {3.525, 0, 2.49375},       {2.9, 0, 2.475},           {2.9, -0.15, 2.475}, _
    {3.45, -0.15, 2.5125},     {3.45, 0, 2.5125},         {2.8, 0, 2.4}, _
    {2.8, -0.15, 2.4},         {3.2, -0.15, 2.4},         {3.2, 0, 2.4}, _
    {0, 0, 3.15},              {0.8, 0, 3.15},            {0.8, -0.45, 3.15}, _
    {0.45, -0.8, 3.15},        {0, -0.8, 3.15},           {0, 0, 2.85}, _
    {1.4, 0, 2.4},             {1.4, -0.784, 2.4},        {0.784, -1.4, 2.4}, _
    {0, -1.4, 2.4},            {0.4, 0, 2.55},            {0.4, -0.224, 2.55}, _
    {0.224, -0.4, 2.55},       {0, -0.4, 2.55},           {1.3, 0, 2.55}, _
    {1.3, -0.728, 2.55},       {0.728, -1.3, 2.55},       {0, -1.3, 2.55}, _
    {1.3, 0, 2.4},             {1.3, -0.728, 2.4},        {0.728, -1.3, 2.4}, _
    {0, -1.3, 2.4},            {0, 0, 0},                 {1.425, -0.798, 0}, _
    {1.5, 0, 0.075},           {1.425, 0, 0},             {0.798, -1.425, 0}, _
    {0, -1.5, 0.075},          {0, -1.425, 0},            {1.5, -0.84, 0.075}, _
    {0.84, -1.5, 0.075} _
    }

   STATIC tex(1, 1, 1) AS DOUBLE = { _
    { {0.0, 0.0}, {1.0, 0.0} }, _
    { {0.0, 1.0}, {1.0, 1.0} } _
    }

   DIM p(3, 3, 2) AS DOUBLE
   DIM q(3, 3, 2) AS DOUBLE
   DIM r(3, 3, 2) AS DOUBLE
   DIM s(3, 3, 2) AS DOUBLE

   glPushAttrib(GL_ENABLE_BIT OR GL_EVAL_BIT)
   glEnable(GL_AUTO_NORMAL)
   glEnable(GL_NORMALIZE)
   glEnable(GL_MAP2_VERTEX_3)
   glEnable(GL_MAP2_TEXTURE_COORD_2)

   glPushMatrix
   glRotated(270.0, 1.0, 0.0, 0.0)
   glScaled(0.5 * dbScale, 0.5 * dbScale, 0.5 * dbScale)
   glTranslated(0.0, 0.0, -1.5)

   FOR i AS LONG = 0 TO 9
      FOR j AS LONG = 0 TO 3
         FOR k AS LONG = 0 TO 3
            FOR l AS LONG = 0 TO 2
               p(j,k,l) = cpdata(patchdata(i,j * 4 + k), l)
               q(j,k,l) = cpdata(patchdata(i,j * 4 + (3 - k)),l)
               IF l = 1 THEN q(j,k,l) *= -1.0
               IF i < 6 THEN r(j,k,l) = cpdata(patchdata(i,j * 4 + (3 - k)), l)
               IF l = 0 THEN r(j,k,l) *= -1.0
               s(j,k,l) = cpdata(patchdata(i,j * 4 + k),l)
               IF l = 0 THEN s(j,k,l) *= -1.0
               IF l = 1 THEN s(j,k,l) *= -1.0
            NEXT
         NEXT
      NEXT

      glMap2d(GL_MAP2_TEXTURE_COORD_2, 0.0, 1.0, 2, 2, 0.0, 1.0, 4, 2, @tex(0,0,0))
      glMap2d(GL_MAP2_VERTEX_3, 0.0, 1.0, 3, 4, 0.0, 1.0, 12, 4, @p(0,0,0))
      glMapGrid2d(grid, 0.0, 1.0, grid, 0.0, 1.0)
      glEvalMesh2(dwType, 0, grid, 0, grid)
      glMap2d(GL_MAP2_VERTEX_3, 0.0, 1.0, 3, 4, 0.0, 1.0, 12, 4, @q(0,0,0))
      glEvalMesh2(dwType, 0, grid, 0, grid)
      IF i < 6 THEN
         glMap2d(GL_MAP2_VERTEX_3, 0.0, 1.0, 3, 4, 0.0, 1.0, 12, 4, @r(0,0,0))
         glEvalMesh2(dwType, 0, grid, 0, grid)
         glMap2d(GL_MAP2_VERTEX_3, 0.0, 1.0, 3, 4, 0.0, 1.0, 12, 4, @s(0,0,0))
         glEvalMesh2(dwType, 0, grid, 0, grid)
      END IF
   NEXT

   glPopMatrix
   glPopAttrib

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a wireframe teapot.
' The AfxGlutWireTeapot procedure renders a teapot of the desired size, centered at the origin.
' This is the famous teapot created by Martin Newell. The other functions render the teacup
' and teaspoon he used in the table scene figure in his PhD thesis. Vertex data retrieved
' from: ftp://ftp.funet.fi/pub/sci/graphics/packages/objects/teasetorig.gz.
' Parameter:
' * dbSize: The desired size of the teapot, teacup and teaspoon - relative to a "standard" size.
' Bugs:
' OpenGL's default glFrontFace state assumes that front facing polygons (for the purpose of
' face culling) have vertices that wind counter clockwise when projected into window space.
' This teapot, teacup and teaspoon are rendered with their front facing polygon vertices
' winding clockwise. For OpenGL's default back face culling to work, you should use:
'   glFrontFace(GL_CW);
'   glutSolidTeapot(size);
'   glFrontFace(GL_CCW);
' This bug reflect issues in the original teaset's vertex data (and is thus present in GLUT too).
' ========================================================================================
SUB AfxGlutWireTeapot (BYVAL dbSize AS DOUBLE)
   ' /* We will use the general teapot rendering code */
   AfxGlutTeapot(10, dbSize, GL_LINE)
END SUB
' ========================================================================================

' ========================================================================================
' * Draws a solid teapot.
' The AfxGlutSolidTeapot procedure renders a teapot of the desired size, centered at the origin.
' This is the famous teapot created by Martin Newell. The other functions render the teacup
' and teaspoon he used in the table scene figure in his PhD thesis. Vertex data retrieved
' from: ftp://ftp.funet.fi/pub/sci/graphics/packages/objects/teasetorig.gz.
' Parameter:
' * dbSize: The desired size of the teapot, teacup and teaspoon - relative to a "standard" size.
' Bugs:
' OpenGL's default glFrontFace state assumes that front facing polygons (for the purpose of
' face culling) have vertices that wind counter clockwise when projected into window space.
' This teapot, teacup and teaspoon are rendered with their front facing polygon vertices
' winding clockwise. For OpenGL's default back face culling to work, you should use:
'   glFrontFace(GL_CW);
'   glutSolidTeapot(size);
'   glFrontFace(GL_CCW);
' This bug reflect issues in the original teaset's vertex data (and is thus present in GLUT too).
' ========================================================================================
SUB AfxGlutSolidTeapot (BYVAL dbSize AS DOUBLE)
   ' /* We will use the general teapot rendering code */
   AfxGlutTeapot(7, dbSize, GL_FILL)
END SUB
' ========================================================================================

' ========================================================================================
' * Draws a wire framed cube.
' The AfxGlutWireCube procedure renders a cube of the desired size, centered at the origin.
' Its faces are normal to the coordinate directions.
' Parameter:
' * dSize: The desired length of an edge of the cube
' ========================================================================================
SUB AfxGlutWireCube (BYVAL dSize AS DOUBLE)

   DIM dblSize AS DOUBLE
   dblSize = dSize * 0.5

   glBegin(GL_LINE_LOOP)
      glNormal3d(1.0, 0.0, 0.0)
      glVertex3d(dblSize, -dblSize, dblSize)
      glVertex3d(dblSize, -dblSize, -dblSize)
      glVertex3d(dblSize, dblSize, -dblSize)
      glVertex3d(dblSize, dblSize, dblSize)
   glEnd
   glBegin(GL_LINE_LOOP)
      glNormal3d(0.0, 1.0, 0.0)
      glVertex3d(dblSize, dblSize, dblSize)
      glVertex3d(dblSize, dblSize, -dblSize)
      glVertex3d(-dblSize, dblSize, -dblSize)
      glVertex3d(-dblSize, dblSize, dblSize)
   glEnd
   glBegin(GL_LINE_LOOP)
      glNormal3d(0.0, 0.0, 1.0)
      glVertex3d(dblSize, dblSize, dblSize)
      glVertex3d(-dblSize, dblSize, dblSize)
      glVertex3d(-dblSize, -dblSize, dblSize)
      glVertex3d(dblSize, -dblSize, dblSize)
   glEnd
   glBegin(GL_LINE_LOOP)
      glNormal3d(-1.0, 0.0, 0.0)
      glVertex3d(-dblSize, -dblSize, dblSize)
      glVertex3d(-dblSize, dblSize, dblSize)
      glVertex3d(-dblSize, dblSize, -dblSize)
      glVertex3d(-dblSize, -dblSize, -dblSize)
   glEnd
   glBegin(GL_LINE_LOOP)
      glNormal3d(0.0,-1.0, 0.0)
      glVertex3d(-dblSize, -dblSize, dblSize)
      glVertex3d(-dblSize, -dblSize, -dblSize)
      glVertex3d(dblSize, -dblSize, -dblSize)
      glVertex3d(dblSize, -dblSize, dblSize)
   glEnd
   glBegin(GL_LINE_LOOP)
      glNormal3d(0.0, 0.0,-1.0)
      glVertex3d(-dblSize, -dblSize, -dblSize)
      glVertex3d(-dblSize, dblSize, -dblSize)
      glVertex3d(dblSize, dblSize, -dblSize)
      glVertex3d(dblSize, -dblSize, -dblSize)
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a solid cube.
' The AfxGlutSolidCube procedure renders a cube of the desired size, centered at the origin.
' Its faces are normal to the coordinate directions.
' Parameter:
' * dSize: The desired length of an edge of the cube
' ========================================================================================
SUB AfxGlutSolidCube (BYVAL dSize AS DOUBLE)

   DIM dblSize AS DOUBLE
   dblSize = dSize * 0.5

   glBegin(GL_QUADS)
      glNormal3d(1.0, 0.0, 0.0)
      glVertex3d(dblSize, -dblSize, dblSize)
      glVertex3d(dblSize, -dblSize, -dblSize)
      glVertex3d(dblSize, dblSize, -dblSize)
      glVertex3d(dblSize, dblSize, dblSize)
      glNormal3d(0.0, 1.0, 0.0)
      glVertex3d(dblSize, dblSize, dblSize)
      glVertex3d(dblSize, dblSize, -dblSize)
      glVertex3d(-dblSize, dblSize, -dblSize)
      glVertex3d(-dblSize, dblSize, dblSize)
      glNormal3d(0.0, 0.0, 1.0)
      glVertex3d(dblSize, dblSize, dblSize)
      glVertex3d(-dblSize, dblSize, dblSize)
      glVertex3d(-dblSize, -dblSize, dblSize)
      glVertex3d(dblSize, -dblSize, dblSize)
      glNormal3d(-1.0, 0.0, 0.0)
      glVertex3d(-dblSize, -dblSize, dblSize)
      glVertex3d(-dblSize, dblSize, dblSize)
      glVertex3d(-dblSize, dblSize, -dblSize)
      glVertex3d(-dblSize, -dblSize, -dblSize)
      glNormal3d(0.0,-1.0, 0.0)
      glVertex3d(-dblSize, -dblSize, dblSize)
      glVertex3d(-dblSize, -dblSize, -dblSize)
      glVertex3d(dblSize, -dblSize, -dblSize)
      glVertex3d(dblSize, -dblSize, dblSize)
      glNormal3d(0.0, 0.0,-1.0)
      glVertex3d(-dblSize, -dblSize, -dblSize)
      glVertex3d(-dblSize, dblSize, -dblSize)
      glVertex3d(dblSize, dblSize, -dblSize)
      glVertex3d(dblSize, -dblSize, -dblSize)
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a wireframe octahedron (eight-sided Platonic solid).
' The AfxGlutWireOctahedron procedure renders an octahedron whose corners are each a
' distance of one from the origin. The length of each side is sqrt(2). The corners are
' on the positive and negative coordinate axes.
' ========================================================================================
SUB AfxGlutWireOctahedron

   DIM RADIUS AS DOUBLE
   RADIUS = 1.0

   glBegin(GL_LINE_LOOP)
      glNormal3d( 0.577350269189, 0.577350269189, 0.577350269189) : glVertex3d( RADIUS, 0.0, 0.0) : glVertex3d( 0.0, RADIUS, 0.0) : glVertex3d(0.0, 0.0, RADIUS)
      glNormal3d( 0.577350269189, 0.577350269189,-0.577350269189) : glVertex3d( RADIUS, 0.0, 0.0) : glVertex3d( 0.0, 0.0,-RADIUS) : glVertex3d(0.0, RADIUS, 0.0)
      glNormal3d( 0.577350269189,-0.577350269189, 0.577350269189) : glVertex3d( RADIUS, 0.0, 0.0) : glVertex3d( 0.0, 0.0, RADIUS) : glVertex3d(0.0,-RADIUS, 0.0)
      glNormal3d( 0.577350269189,-0.577350269189,-0.577350269189) : glVertex3d( RADIUS, 0.0, 0.0) : glVertex3d( 0.0,-RADIUS, 0.0) : glVertex3d(0.0, 0.0,-RADIUS)
      glNormal3d(-0.577350269189, 0.577350269189, 0.577350269189) : glVertex3d(-RADIUS, 0.0, 0.0) : glVertex3d( 0.0, 0.0, RADIUS) : glVertex3d(0.0, RADIUS, 0.0)
      glNormal3d(-0.577350269189, 0.577350269189,-0.577350269189) : glVertex3d(-RADIUS, 0.0, 0.0) : glVertex3d( 0.0, RADIUS, 0.0) : glVertex3d(0.0, 0.0,-RADIUS)
      glNormal3d(-0.577350269189,-0.577350269189, 0.577350269189) : glVertex3d(-RADIUS, 0.0, 0.0) : glVertex3d( 0.0,-RADIUS, 0.0) : glVertex3d(0.0, 0.0, RADIUS)
      glNormal3d(-0.577350269189,-0.577350269189,-0.577350269189) : glVertex3d(-RADIUS, 0.0, 0.0) : glVertex3d( 0.0, 0.0,-RADIUS) : glVertex3d(0.0,-RADIUS, 0.0)
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a solid octahedron (eight-sided Platonic solid).
' The AfxGlutSolidOctahedron procedure renders an octahedron whose corners are each a
' distance of one from the origin. The length of each side is sqrt(2). The corners are
' on the positive and negative coordinate axes.
' ========================================================================================
SUB AfxGlutSolidOctahedron

   DIM RADIUS AS DOUBLE
   RADIUS = 1.0

   glBegin(GL_TRIANGLES)
      glNormal3d( 0.577350269189, 0.577350269189, 0.577350269189) : glVertex3d( RADIUS, 0.0, 0.0) : glVertex3d( 0.0, RADIUS, 0.0) : glVertex3d(0.0, 0.0, RADIUS)
      glNormal3d( 0.577350269189, 0.577350269189,-0.577350269189) : glVertex3d( RADIUS, 0.0, 0.0) : glVertex3d( 0.0, 0.0,-RADIUS) : glVertex3d(0.0, RADIUS, 0.0)
      glNormal3d( 0.577350269189,-0.577350269189, 0.577350269189) : glVertex3d( RADIUS, 0.0, 0.0) : glVertex3d( 0.0, 0.0, RADIUS) : glVertex3d(0.0,-RADIUS, 0.0)
      glNormal3d( 0.577350269189,-0.577350269189,-0.577350269189) : glVertex3d( RADIUS, 0.0, 0.0) : glVertex3d( 0.0,-RADIUS, 0.0) : glVertex3d(0.0, 0.0,-RADIUS)
      glNormal3d(-0.577350269189, 0.577350269189, 0.577350269189) : glVertex3d(-RADIUS, 0.0, 0.0) : glVertex3d( 0.0, 0.0, RADIUS) : glVertex3d(0.0, RADIUS, 0.0)
      glNormal3d(-0.577350269189, 0.577350269189,-0.577350269189) : glVertex3d(-RADIUS, 0.0, 0.0) : glVertex3d( 0.0, RADIUS, 0.0) : glVertex3d(0.0, 0.0,-RADIUS)
      glNormal3d(-0.577350269189,-0.577350269189, 0.577350269189) : glVertex3d(-RADIUS, 0.0, 0.0) : glVertex3d( 0.0,-RADIUS, 0.0) : glVertex3d(0.0, 0.0, RADIUS)
      glNormal3d(-0.577350269189,-0.577350269189,-0.577350269189) : glVertex3d(-RADIUS, 0.0, 0.0) : glVertex3d( 0.0, 0.0,-RADIUS) : glVertex3d(0.0,-RADIUS, 0.0)
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' * Compute lookup table of cos and sin values forming a cirle.
' ========================================================================================
SUB AfxGlutCircleTable (sint() AS DOUBLE, cost() AS DOUBLE, BYVAL n AS LONG)

   ' /* Table size, the sign of n flips the circle direction */
   DIM nSize AS LONG
   nSize = ABS(n)

   ' /* Determine the angle between samples */
   DIM angle AS DOUBLE
   IF n = 0 THEN n = 1
   angle = 2 * 3.14159265358979323846 / IIF(n = 0, 1, n)

   ' /* Allocate memory for n samples, plus duplicate of first entry at the end */
   REDIM sint(nSize) AS DOUBLE
   REDIM cost(nSize) AS DOUBLE

   ' /* Compute cos and sin around the circle */
   sint(0) = 0.0
   cost(0) = 1.0

   FOR i AS LONG = 1 TO nSize - 1
      sint(i) = SIN(angle * i)
      cost(i) = COS(angle * i)
   NEXT

   ' /* Last sample is duplicate of the first */
   sint(nSize) = sint(0)
   cost(nSize) = cost(0)

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a solid sphere.
' The AfxGlutSolidSphere procedure renders a sphere centered at the origin of the modeling
' coordinate system. The north and south poles of the sphere are on the positive and negative
' Z-axes respectively and the prime meridian crosses the positive X-axis.
' Parameters:
' * radius: The desired radius of the sphere.
' * slices: The desired number of slices (divisions in the longitudinal direction) in the sphere.
' * stacks: The desired number of stacks (divisions in the latitudinal direction) in the
'   sphere. The number of points in this direction, including the north and south poles,
'   is stacks+1.
' ========================================================================================
SUB AfxGlutSolidSphere (BYVAL radius AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)

   ' /* Adjust z and radius as stacks are drawn. */
   DIM AS DOUBLE z0, z1
   DIM AS DOUBLE r0, r1

   ' /* Pre-computed circle */
   DIM sint1(ANY) AS DOUBLE
   DIM cost1(ANY) AS DOUBLE
   DIM sint2(ANY) AS DOUBLE
   DIM cost2(ANY) AS DOUBLE

   AfxGlutCircleTable(sint1(), cost1(), -slices)
   AfxGlutCircleTable(sint2(), cost2(), stacks * 2)

   ' /* The top stack is covered with a triangle fan */

   z0 = 1.0
   z1 = cost2(IIF(stacks > 0, 1, 0))
   r0 = 0.0
   r1 = sint2(IIF(stacks > 0, 1, 0))

   glBegin(GL_TRIANGLE_FAN)
      glNormal3d(0, 0, 1)
      glVertex3d(0, 0, radius)
      FOR j AS LONG = slices TO 0 STEP -1
         glNormal3d(cost1(j)*r1,        sint1(j)*r1,        z1       )
         glVertex3d(cost1(j)*r1*radius, sint1(j)*r1*radius, z1*radius)
      NEXT
   glEnd

   ' /* Cover each stack with a quad strip, except the top and bottom stacks */

   FOR i AS LONG = 1 TO stacks - 2
      z0 = z1 : z1 = cost2(i + 1)
      r0 = r1 : r1 = sint2(i + 1)
      glBegin(GL_QUAD_STRIP)
         FOR j AS LONG = 0 TO slices
            glNormal3d(cost1(j)*r1,        sint1(j)*r1,        z1       )
            glVertex3d(cost1(j)*r1*radius, sint1(j)*r1*radius, z1*radius)
            glNormal3d(cost1(j)*r0,        sint1(j)*r0,        z0       )
            glVertex3d(cost1(j)*r0*radius, sint1(j)*r0*radius, z0*radius)
         NEXT
      glEnd
   NEXT

   ' /* The bottom stack is covered with a triangle fan */

   z0 = z1
   r0 = r1

   glBegin(GL_TRIANGLE_FAN)
      glNormal3d(0, 0, -1)
      glVertex3d(0, 0, -radius)
      FOR j AS LONG = 0 TO slices
         glNormal3d(cost1(j)*r0,        sint1(j)*r0,        z0       )
         glVertex3d(cost1(j)*r0*radius, sint1(j)*r0*radius, z0*radius)
      NEXT
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a wire sphere.
' The AfxGlutWireSphere procedure renders a sphere centered at the origin of the modeling
' coordinate system. The north and south poles of the sphere are on the positive and negative
' Z-axes respectively and the prime meridian crosses the positive X-axis.
' Parameters:
' * radius: The desired radius of the sphere.
' * slices: The desired number of slices (divisions in the longitudinal direction) in the sphere.
' * stacks: The desired number of stacks (divisions in the latitudinal direction) in the
'   sphere. The number of points in this direction, including the north and south poles,
'   is stacks+1.
' ========================================================================================
SUB AfxGlutWireSphere (BYVAL radius AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)

   ' // Adjust z and radius as stacks and slices are drawn.
   DIM AS DOUBLE r, x, y, z

   ' // Pre-computed circle

   DIM sint1(ANY) AS DOUBLE
   DIM cost1(ANY) AS DOUBLE
   DIM sint2(ANY) AS DOUBLE
   DIM cost2(ANY) AS DOUBLE

   AfxGlutCircleTable(sint1(), cost1(), -slices)
   AfxGlutCircleTable(sint2(), cost2(), stacks * 2)

   ' // Draw a line loop for each stack
   FOR i AS LONG = 1 TO stacks - 1
      z = cost2(i)
      r = sint2(i)
      glBegin(GL_LINE_LOOP)
         FOR j AS LONG = 0 TO slices
            x = cost1(j)
            y = sint1(j)
            glNormal3d(x, y, z)
            glVertex3d(x*r*radius, y*r*radius,z*radius)
         NEXT
      glEnd
   NEXT

   ' // Draw a line loop for each slice
   FOR i AS LONG = 0 TO slices - 1
      glBegin(GL_LINE_STRIP)
         FOR j AS LONG = 0 TO stacks
            x = cost1(i)*sint2(j)
            y = sint1(i)*sint2(j)
            z = cost2(j)
            glNormal3d(x,y,z)
            glVertex3d(x*radius, y*radius,z*radius)
         NEXT
      glEnd
   NEXT

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a solid cone.
' The AfxGlutSolidCone procedure renders a right circular cone with a base centered at the
' origin and in the X-Y plane and its tip on the positive Z-axis. The wire cone is rendered
' with triangular elements.
' Parameters:
' * dbase: The desired radius of the base of the cone
' * height: The desired height of the cone
' * slices: The desired number of slices around the cone
' * stacks: The desired number of segments between the base and the tip of the cone
'   (the number of points, including the tip, is stacks + 1)
' ========================================================================================
SUB AfxGlutSolidCone (BYVAL dbase AS DOUBLE, BYVAL height AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)

   ' /* Step in z and radius as stacks are drawn. */
   DIM AS DOUBLE z0, z1
   DIM AS DOUBLE r0, r1

   DIM AS DOUBLE zStep, rStep
   zStep = height / IIF(stacks > 0, stacks, 1)
   rStep = dbase / IIF(stacks > 0, stacks, 1)

   ' /* Scaling factors for vertex normals */
   DIM AS DOUBLE cosn, sinn
   cosn = (height / SQR(height * height + dbase * dbase))
   sinn = (dbase  / SQR(height * height + dbase * dbase))

   ' /* Pre-computed circle */
   DIM sint(ANY) AS DOUBLE
   DIM cost(ANY) AS DOUBLE

   AfxGlutCircleTable(sint(), cost(), -slices)

   ' /* Cover the circular base with a triangle fan... */

   z0 = 0.0
   z1 = zStep

   r0 = dbase
   r1 = r0 - rStep

   glBegin(GL_TRIANGLE_FAN)
      glNormal3d(0.0, 0.0, -1.0)
      glVertex3d(0.0, 0.0, z0)
      FOR j AS LONG = 0 TO slices
         glVertex3d(cost(j)*r0, sint(j)*r0, z0)
      NEXT
   glEnd

   ' /* Cover each stack with a quad strip, except the top stack */

   FOR i AS LONG = 0 TO stacks - 2
      glBegin(GL_QUAD_STRIP)
         FOR j AS LONG = 0 TO slices
            glNormal3d(cost(j)*sinn, sint(j)*sinn, cosn)
            glVertex3d(cost(j)*r0,   sint(j)*r0,   z0  )
            glVertex3d(cost(j)*r1,   sint(j)*r1,   z1  )
         NEXT
         z0 = z1 : z1 += zStep
         r0 = r1 : r1 -= rStep
      glEnd
   NEXT

   ' /* The top stack is covered with individual triangles */

   glBegin(GL_TRIANGLES)
      glNormal3d(cost(0)*sinn, sint(0)*sinn, cosn)
      FOR j AS LONG = 0 TO slices - 1
         glVertex3d(cost(j+0)*r0,   sint(j+0)*r0,   z0    )
         glVertex3d(0,              0,              height)
         glNormal3d(cost(j+1)*sinn, sint(j+1)*sinn, cosn  )
         glVertex3d(cost(j+1)*r0,   sint(j+1)*r0,   z0    )
      NEXT
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a wireframe cone.
' The AfxGlutWireCone procedure renders a right circular cone with a base centered at the
' origin and in the X-Y plane and its tip on the positive Z-axis. The wire cone is rendered
' with triangular elements.
' Parameters:
' * dbase: The desired radius of the base of the cone
' * height: The desired height of the cone
' * slices: The desired number of slices around the cone
' * stacks: The desired number of segments between the base and the tip of the cone
'   (the number of points, including the tip, is stacks + 1)
' ========================================================================================
SUB AfxGlutWireCone (BYVAL dbase AS DOUBLE, BYVAL height AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)

   ' /* Step in z and radius as stacks are drawn. */
   DIM AS DOUBLE z, r
   z = 0.0
   r = dbase

   DIM AS DOUBLE zStep, rStep
   zStep = height / IIF(stacks > 0, stacks, 1)
   rStep = dbase / IIF(stacks > 0, stacks, 1)

   ' /* Scaling factors for vertex normals */
   DIM AS DOUBLE cosn, sinn
   cosn = (height / SQR(height * height + dbase * dbase))
   sinn = (dbase  / SQR(height * height + dbase * dbase))

   ' /* Pre-computed circle */
   DIM sint(ANY) AS DOUBLE
   DIM cost(ANY) AS DOUBLE

   AfxGlutCircleTable(sint(), cost(), -slices)

   ' /* Draw the stacks... */

   FOR i AS LONG = 0 TO stacks - 1
      glBegin(GL_LINE_LOOP)
         FOR j AS LONG = 0 TO slices - 1
            glNormal3d(cost(j)*sinn, sint(j)*sinn, cosn)
            glVertex3d(cost(j)*r,    sint(j)*r,    z   )
         NEXT
      glEnd
      z += zStep
      r -= rStep
   NEXT

   ' /* Draw the slices */

   r = dbase

   glBegin(GL_LINES)
      FOR j AS LONG = 0 TO slices - 1
         glNormal3d(cost(j)*sinn, sint(j)*sinn, cosn  )
         glVertex3d(cost(j)*r,    sint(j)*r,    0.0   )
         glVertex3d(0.0,          0.0,          height)
      NEXT
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a solid cylinder.
' Parameters:
' * radius: The desired radius of the cylinder
' * height: The desired height of the cylinder
' * slices: The desired number of slices around the cylinder
' * stacks: The desired number of segments between the base and the top of the cylinder
'   (the number of points, including the tip, is stacks + 1)
' ========================================================================================
SUB AfxGlutSolidCylinder (BYVAL radius AS DOUBLE, BYVAL height AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)

   ' /* Step in z and radius as stacks are drawn. */
   DIM zStep AS DOUBLE = height / IIF(stacks > 0, stacks, 1)

   ' /* Pre-computed circle */
   DIM sint(ANY) AS DOUBLE
   DIM cost(ANY) AS DOUBLE

   AfxGlutCircleTable(sint(), cost(), -slices)

   ' /* Cover the base and top */

   glBegin(GL_TRIANGLE_FAN)
      glNormal3d(0.0, 0.0, -1.0)
      glVertex3d(0.0, 0.0,  0.0)
      FOR j AS LONG = 0 TO slices
         glVertex3d(cost(j)*radius, sint(j)*radius, 0.0)
      NEXT
   glEnd

   glBegin(GL_TRIANGLE_FAN)
      glNormal3d(0.0, 0.0, 1.0)
      glVertex3d(0.0, 0.0, height)
      FOR j AS LONG = slices TO 0 STEP - 1
         glVertex3d(cost(j)*radius, sint(j)*radius, height)
      NEXT
   glEnd

   ' /* Do the stacks */
   DIM z0 AS DOUBLE = 0.0
   DIM z1 AS DOUBLE = zStep

   FOR i AS LONG = 1 TO stacks
      IF i = stacks THEN z1 = height
      glBegin(GL_QUAD_STRIP)
         FOR j AS LONG = 0 TO slices
            glNormal3d(cost(j),        sint(j),        0.0)
            glVertex3d(cost(j)*radius, sint(j)*radius, z0 )
            glVertex3d(cost(j)*radius, sint(j)*radius, z1 )
         NEXT
      glEnd
      z0 = z1
      z1 += zStep
   NEXT

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a wireframe cylinder.
' Parameters:
' * radius: The desired radius of the cylinder
' * height: The desired height of the cylinder
' * slices: The desired number of slices around the cylinder
' * stacks: The desired number of segments between the base and the top of the cylinder
'   (the number of points, including the tip, is stacks + 1)
' ========================================================================================
SUB AfxGlutWireCylinder (BYVAL radius AS DOUBLE, BYVAL height AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)

   ' /* Step in z and radius as stacks are drawn. */
   DIM z AS DOUBLE
   DIM zStep AS DOUBLE = height / IIF(stacks > 0, stacks, 1)

   ' /* Pre-computed circle */
   DIM sint(ANY) AS DOUBLE
   DIM cost(ANY) AS DOUBLE

   AfxGlutCircleTable(sint(), cost(), -slices)

   ' /* Draw the stacks... */
   FOR i AS LONG = 0 TO stacks
      IF i = stacks THEN z = height
      glBegin(GL_LINE_LOOP)
         FOR j AS LONG = 0 TO slices - 1
            glNormal3d(cost(j),        sint(j),        0.0)
            glVertex3d(cost(j)*radius, sint(j)*radius, z  )
         NEXT
      glEnd
      z += zStep
   NEXT

   ' /* Draw the slices */
   glBegin(GL_LINES)
      FOR j AS LONG = 0 TO slices - 1
         glNormal3d(cost(j),        sint(j),        0.0   )
         glVertex3d(cost(j)*radius, sint(j)*radius, 0.0   )
         glVertex3d(cost(j)*radius, sint(j)*radius, height)
      NEXT
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' * Renders a solid dodecahedron (12-sided Platonic solid).
' The AfxGlutsolidDodecahedron procedure renders a dodecahedron whose corners are each a
' distance of sqrt(3) from the origin. The length of each side is sqrt(5)-1. There are
' twenty corners; interestingly enough, eight of them coincide with the corners of a cube
' with sizes of length 2.
' ========================================================================================
SUB AfxGlutSolidDodecahedron

   '/* Magic Numbers:  It is possible to create a dodecahedron by attaching two pentagons to each face of
   ' * of a cube.  The coordinates of the points are:
   ' *  (+-x,0, z);(+-1, 1, 1);(0, z, x)
   ' * where x =(-1 + sqrt(5))/2, z =(1 + sqrt(5))/2 or
   ' *       x = 0.61803398875 and z = 1.61803398875.
   '*/

   glBegin(GL_POLYGON)
      glNormal3d( 0.0, 0.525731112119, 0.850650808354)
      glVertex3d( 0.0, 1.61803398875, 0.61803398875)
      glVertex3d(-1.0, 1.0, 1.0)
      glVertex3d(-0.61803398875, 0.0, 1.61803398875)
      glVertex3d( 0.61803398875, 0.0, 1.61803398875)
      glVertex3d( 1.0, 1.0, 1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d( 0.0, 0.525731112119, -0.850650808354)
      glVertex3d( 0.0, 1.61803398875, -0.61803398875)
      glVertex3d( 1.0, 1.0, -1.0)
      glVertex3d( 0.61803398875, 0.0, -1.61803398875)
      glVertex3d(-0.61803398875, 0.0, -1.61803398875)
      glVertex3d(-1.0, 1.0, -1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d( 0.0, -0.525731112119, 0.850650808354)
      glVertex3d( 0.0, -1.61803398875, 0.61803398875)
      glVertex3d( 1.0, -1.0, 1.0)
      glVertex3d( 0.61803398875, 0.0, 1.61803398875)
      glVertex3d(-0.61803398875, 0.0, 1.61803398875)
      glVertex3d(-1.0, -1.0, 1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d( 0.0, -0.525731112119, -0.850650808354)
      glVertex3d( 0.0, -1.61803398875, -0.61803398875)
      glVertex3d(-1.0, -1.0, -1.0)
      glVertex3d(-0.61803398875, 0.0, -1.61803398875)
      glVertex3d( 0.61803398875, 0.0, -1.61803398875)
      glVertex3d( 1.0, -1.0, -1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d( 0.850650808354, 0.0, 0.525731112119)
      glVertex3d( 0.61803398875, 0.0, 1.61803398875)
      glVertex3d( 1.0, -1.0, 1.0)
      glVertex3d( 1.61803398875, -0.61803398875, 0.0)
      glVertex3d( 1.61803398875, 0.61803398875, 0.0)
      glVertex3d( 1.0, 1.0, 1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d(-0.850650808354, 0.0, 0.525731112119)
      glVertex3d(-0.61803398875, 0.0, 1.61803398875)
      glVertex3d(-1.0, 1.0, 1.0)
      glVertex3d(-1.61803398875, 0.61803398875, 0.0)
      glVertex3d(-1.61803398875, -0.61803398875, 0.0)
      glVertex3d(-1.0, -1.0, 1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d( 0.850650808354, 0.0, -0.525731112119)
      glVertex3d( 0.61803398875, 0.0, -1.61803398875)
      glVertex3d( 1.0, 1.0, -1.0)
      glVertex3d( 1.61803398875, 0.61803398875, 0.0)
      glVertex3d( 1.61803398875, -0.61803398875, 0.0)
      glVertex3d( 1.0, -1.0, -1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d(-0.850650808354, 0.0, -0.525731112119)
      glVertex3d(-0.61803398875, 0.0, -1.61803398875)
      glVertex3d(-1.0, -1.0, -1.0)
      glVertex3d(-1.61803398875, -0.61803398875, 0.0)
      glVertex3d(-1.61803398875, 0.61803398875, 0.0)
      glVertex3d(-1.0, 1.0, -1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d( 0.525731112119, 0.850650808354, 0.0)
      glVertex3d( 1.61803398875, 0.61803398875, 0.0)
      glVertex3d( 1.0, 1.0, -1.0)
      glVertex3d(0.0, 1.61803398875, -0.61803398875)
      glVertex3d(0.0, 1.61803398875, 0.61803398875)
      glVertex3d( 1.0, 1.0, 1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d( 0.525731112119, -0.850650808354, 0.0)
      glVertex3d( 1.61803398875, -0.61803398875, 0.0)
      glVertex3d( 1.0, -1.0, 1.0)
      glVertex3d(0.0, -1.61803398875, 0.61803398875)
      glVertex3d(0.0, -1.61803398875, -0.61803398875)
      glVertex3d( 1.0, -1.0, -1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d(-0.525731112119, 0.850650808354, 0.0)
      glVertex3d(-1.61803398875, 0.61803398875, 0.0)
      glVertex3d(-1.0, 1.0, 1.0)
      glVertex3d(0.0, 1.61803398875, 0.61803398875)
      glVertex3d(0.0, 1.61803398875, -0.61803398875)
      glVertex3d(-1.0, 1.0, -1.0)
   glEnd

   glBegin(GL_POLYGON)
      glNormal3d(-0.525731112119, -0.850650808354, 0.0)
      glVertex3d(-1.61803398875, -0.61803398875, 0.0)
      glVertex3d(-1.0, -1.0, -1.0)
      glVertex3d(0.0, -1.61803398875, -0.61803398875)
      glVertex3d(0.0, -1.61803398875, 0.61803398875)
      glVertex3d(-1.0, -1.0, 1.0)
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' * Renders a wireframe dodecahedron (12-sided Platonic solid).
' The AfxGlutWireDodecahedron procedure renders a dodecahedron whose corners are each a
' distance of sqrt(3) from the origin. The length of each side is sqrt(5)-1. There are
' twenty corners; interestingly enough, eight of them coincide with the corners of a cube
' with sizes of length 2.
' ========================================================================================
SUB AfxGlutWireDodecahedron

   '/* Magic Numbers:  It is possible to create a dodecahedron by attaching two pentagons to each face of
   ' * of a cube.  The coordinates of the points are:
   ' *   (+-x,0, z); (+-1, 1, 1); (0, z, x )
   ' * where x = (-1 + sqrt(5))/2, z = (1 + sqrt(5))/2  or
   ' *       x = 0.61803398875 and z = 1.61803398875.
   ' */
   glBegin(GL_LINE_LOOP)
      glNormal3d(0.0, 0.525731112119, 0.850650808354)
      glVertex3d (0.0, 1.61803398875, 0.61803398875)
      glVertex3d(-1.0, 1.0, 1.0 )
      glVertex3d(-0.61803398875, 0.0, 1.61803398875)
      glVertex3d(0.61803398875, 0.0, 1.61803398875)
      glVertex3d(1.0, 1.0, 1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(0.0, 0.525731112119, -0.850650808354)
      glVertex3d(0.0, 1.61803398875, -0.61803398875)
      glVertex3d(1.0, 1.0, -1.0)
      glVertex3d(0.61803398875, 0.0, -1.61803398875)
      glVertex3d(-0.61803398875, 0.0, -1.61803398875)
      glVertex3d(-1.0, 1.0, -1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(0.0, -0.525731112119, 0.850650808354)
      glVertex3d(0.0, -1.61803398875, 0.61803398875)
      glVertex3d(1.0, -1.0, 1.0)
      glVertex3d(0.61803398875, 0.0, 1.61803398875)
      glVertex3d( -0.61803398875, 0.0, 1.61803398875)
      glVertex3d(-1.0, -1.0, 1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(0.0, -0.525731112119, -0.850650808354)
      glVertex3d(0.0, -1.61803398875, -0.61803398875)
      glVertex3d(-1.0, -1.0, -1.0)
      glVertex3d(-0.61803398875, 0.0, -1.61803398875)
      glVertex3d(0.61803398875, 0.0, -1.61803398875)
      glVertex3d(1.0, -1.0, -1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(0.850650808354, 0.0, 0.525731112119)
      glVertex3d(0.61803398875, 0.0, 1.61803398875)
      glVertex3d(1.0, -1.0, 1.0)
      glVertex3d(1.61803398875, -0.61803398875, 0.0)
      glVertex3d(1.61803398875, 0.61803398875, 0.0)
      glVertex3d(1.0, 1.0, 1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(-0.850650808354, 0.0, 0.525731112119)
      glVertex3d(-0.61803398875, 0.0, 1.61803398875)
      glVertex3d(-1.0, 1.0, 1.0)
      glVertex3d(-1.61803398875, 0.61803398875, 0.0)
      glVertex3d(-1.61803398875, -0.61803398875, 0.0)
      glVertex3d(-1.0, -1.0, 1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(0.850650808354, 0.0, -0.525731112119)
      glVertex3d(0.61803398875, 0.0, -1.61803398875)
      glVertex3d(1.0, 1.0, -1.0)
      glVertex3d(1.61803398875, 0.61803398875, 0.0)
      glVertex3d(1.61803398875, -0.61803398875, 0.0)
      glVertex3d(1.0, -1.0, -1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(-0.850650808354, 0.0, -0.525731112119)
      glVertex3d(-0.61803398875, 0.0, -1.61803398875)
      glVertex3d(-1.0, -1.0, -1.0)
      glVertex3d(-1.61803398875, -0.61803398875, 0.0)
      glVertex3d(-1.61803398875,  0.61803398875, 0.0)
      glVertex3d(-1.0, 1.0, -1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(0.525731112119, 0.850650808354, 0.0)
      glVertex3d(1.61803398875, 0.61803398875, 0.0)
      glVertex3d(1.0, 1.0, -1.0)
      glVertex3d(0.0, 1.61803398875, -0.61803398875)
      glVertex3d(0.0, 1.61803398875, 0.61803398875)
      glVertex3d(1.0, 1.0, 1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(0.525731112119, -0.850650808354, 0.0)
      glVertex3d(1.61803398875, -0.61803398875, 0.0)
      glVertex3d(1.0, -1.0, 1.0)
      glVertex3d(0.0, -1.61803398875, 0.61803398875)
      glVertex3d(0.0, -1.61803398875, -0.61803398875)
      glVertex3d(1.0, -1.0, -1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(-0.525731112119, 0.850650808354, 0.0)
      glVertex3d(-1.61803398875, 0.61803398875, 0.0)
      glVertex3d(-1.0, 1.0, 1.0)
      glVertex3d(0.0, 1.61803398875, 0.61803398875)
      glVertex3d(0.0, 1.61803398875, -0.61803398875)
      glVertex3d(-1.0, 1.0, -1.0)
   glEnd

   glBegin(GL_LINE_LOOP)
      glNormal3d(-0.525731112119, -0.850650808354, 0.0)
      glVertex3d(-1.61803398875, -0.61803398875, 0.0)
      glVertex3d(-1.0, -1.0, -1.0)
      glVertex3d(0.0, -1.61803398875, -0.61803398875)
      glVertex3d(0.0, -1.61803398875, 0.61803398875)
      glVertex3d(-1.0, -1.0, 1.0)
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' Renders a solid tetrahedron (4-sided Platonic solid).
' The AfxGlutSolidTetrahedron procedure renders a tetrahedron whose corners are each a
' distance of one from the origin. The length of each side is 2/3 sqrt(6). One corner is
' on the positive X-axis and another is in the X-Y plane with a positive Y-coordinate.
' ========================================================================================
SUB AfxGlutSolidTetrahedron

   STATIC tet_r(3, 2) AS DOUBLE = { _
    {             1.0,             0.0,             0.0 }, _
    { -0.333333333333,  0.942809041582,             0.0 }, _
    { -0.333333333333, -0.471404520791,  0.816496580928 }, _
    { -0.333333333333, -0.471404520791, -0.816496580928 } _
    }

   glBegin(GL_TRIANGLES)
      glNormal3d(-tet_r(0, 0), -tet_r(0, 1), -tet_r(0, 2))
      glVertex3d(tet_r(1, 0), tet_r(1, 1), tet_r(1, 2))
      glVertex3d(tet_r(3, 0), tet_r(3, 1), tet_r(3, 2))
      glVertex3d(tet_r(2, 0), tet_r(2, 1), tet_r(2, 2))

      glNormal3d(-tet_r(1, 0), -tet_r(1, 1), -tet_r(1, 2))
      glVertex3d(tet_r(0, 0), tet_r(0, 1), tet_r(0, 2))
      glVertex3d(tet_r(2, 0), tet_r(2, 1), tet_r(2, 2))
      glVertex3d(tet_r(3, 0), tet_r(3, 1), tet_r(3, 2))

      glNormal3d(-tet_r(2, 0), -tet_r(2, 1), -tet_r(2, 2))
      glVertex3d(tet_r(0, 0), tet_r(0, 1), tet_r(0, 2))
      glVertex3d(tet_r(3, 0), tet_r(3, 1), tet_r(3, 2))
      glVertex3d(tet_r(1, 0), tet_r(1, 1), tet_r(1, 2))

      glNormal3d(-tet_r(3, 0), -tet_r(3, 1), -tet_r(3, 2))
      glVertex3d(tet_r(0, 0), tet_r(0, 1), tet_r(0, 2))
      glVertex3d(tet_r(1, 0), tet_r(1, 1), tet_r(1, 2))
      glVertex3d(tet_r(2, 0), tet_r(2, 1), tet_r(2, 2))
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' Renders a wireframe tetrahedron (4-sided Platonic solid).
' The AfxGlutWireTetrahedron procedure renders a tetrahedron whose corners are each a
' distance of one from the origin. The length of each side is 2/3 sqrt(6). One corner is
' on the positive X-axis and another is in the X-Y plane with a positive Y-coordinate.
' ========================================================================================
SUB AfxGlutWireTetrahedron

   STATIC tet_r(3, 2) AS DOUBLE = { _
    {             1.0,             0.0,             0.0 }, _
    { -0.333333333333,  0.942809041582,             0.0 }, _
    { -0.333333333333, -0.471404520791,  0.816496580928 }, _
    { -0.333333333333, -0.471404520791, -0.816496580928 } _
    }

   glBegin(GL_LINE_LOOP)
      glNormal3d(-tet_r(0, 0), -tet_r(0, 1), -tet_r(0, 2))
      glVertex3d(tet_r(1, 0), tet_r(1, 1), tet_r(1, 2))
      glVertex3d(tet_r(3, 0), tet_r(3, 1), tet_r(3, 2))
      glVertex3d(tet_r(2, 0), tet_r(2, 1), tet_r(2, 2))

      glNormal3d(-tet_r(1, 0), -tet_r(1, 1), -tet_r(1, 2))
      glVertex3d(tet_r(0, 0), tet_r(0, 1), tet_r(0, 2))
      glVertex3d(tet_r(2, 0), tet_r(2, 1), tet_r(2, 2))
      glVertex3d(tet_r(3, 0), tet_r(3, 1), tet_r(3, 2))

      glNormal3d(-tet_r(2, 0), -tet_r(2, 1), -tet_r(2, 2))
      glVertex3d(tet_r(0, 0), tet_r(0, 1), tet_r(0, 2))
      glVertex3d(tet_r(3, 0), tet_r(3, 1), tet_r(3, 2))
      glVertex3d(tet_r(1, 0), tet_r(1, 1), tet_r(1, 2))

      glNormal3d(-tet_r(3, 0), -tet_r(3, 1), -tet_r(3, 2))
      glVertex3d(tet_r(0, 0), tet_r(0, 1), tet_r(0, 2))
      glVertex3d(tet_r(1, 0), tet_r(1, 1), tet_r(1, 2))
      glVertex3d(tet_r(2, 0), tet_r(2, 1), tet_r(2, 2))
   glEnd

END SUB
' ========================================================================================

' ========================================================================================
' * Renders a torus (doughnut)
' Parameters:
' * innerRadius: Inner radius of the torus.
' * outerRadius: Outer radius of the torus.
' * nsides: Number of sides for each radial section.
' * rings: Number of radial divisions for the torus.
' ========================================================================================
SUB AfxGlutDoughnut (BYVAL innerRadius AS SINGLE, BYVAL outerRadius AS SINGLE , BYVAL nsides AS LONG, BYVAL rings AS LONG)

  DIM theta AS SINGLE
  DIM phi AS SINGLE
  DIM theta1 AS SINGLE
  DIM cosTheta AS SINGLE
  DIM sinTheta AS SINGLE
  DIM cosTheta1 AS SINGLE
  DIM sinTheta1 AS SINGLE
  DIM ringDelta AS SINGLE
  DIM sideDelta AS SINGLE  '// Double

  ringDelta = 2.0 * 3.14159265358979323846 / rings
  sideDelta = 2.0 * 3.14159265358979323846 / nsides
  theta = 0.0
  cosTheta = 1.0
  sinTheta = 0.0

  FOR i AS LONG = rings - 1 TO 0 STEP -1
    theta1 = theta + ringDelta
    cosTheta1 = COS(theta1)
    sinTheta1 = SIN(theta1)
    glBegin(GL_QUAD_STRIP)
    phi = 0.0
    FOR j AS LONG = nsides TO 0 STEP -1
      DIM cosPhi AS SINGLE
      DIM sinPhi AS SINGLE
      DIM dist AS SINGLE
      phi += sideDelta
      cosPhi = COS(phi)
      sinPhi = SIN(phi)
      dist = outerRadius + innerRadius * cosPhi
      glNormal3f(cosTheta1 * cosPhi, -sinTheta1 * cosPhi, sinPhi)
      glVertex3f(cosTheta1 * dist, -sinTheta1 * dist, innerRadius * sinPhi)
      glNormal3f(cosTheta * cosPhi, -sinTheta * cosPhi, sinPhi)
      glVertex3f(cosTheta * dist, -sinTheta * dist, innerRadius * sinPhi)
    NEXT
    glEnd
    theta = theta1
    cosTheta = cosTheta1
    sinTheta = sinTheta1
  NEXT
END SUB
' ========================================================================================

' ========================================================================================
' * Renders a wireframe torus (doughnut shape).
' The AfxGlutWireTorus procedure renders a torus centered at the origin of the modeling
' coordinate system. The torus is circularly symmetric about the Z-axis and starts at the
' positive X-axis.
' ========================================================================================
SUB AfxGlutWireTorus (BYVAL innerRadius AS DOUBLE, BYVAL outerRadius AS DOUBLE, BYVAL nsides AS LONG, BYVAL rings AS LONG)
   glPushAttrib(GL_POLYGON_BIT)
   glPolygonMode(GL_FRONT_AND_BACK, GL_LINE)
   AfxGlutDoughnut(innerRadius, outerRadius, nsides, rings)
   glPopAttrib
END SUB
' ========================================================================================

' ========================================================================================
' * Renders a solid torus (doughnut shape).
' The AfxGlutSolidTorus procedure renders a torus centered at the origin of the modeling
' coordinate system. The torus is circularly symmetric about the Z-axis and starts at the
' positive X-axis.
' ========================================================================================
SUB AfxGlutSolidTorus (BYVAL innerRadius AS DOUBLE, BYVAL outerRadius AS DOUBLE, BYVAL nsides AS LONG, BYVAL rings AS LONG)
   AfxGlutDoughnut(innerRadius, outerRadius, nsides, rings)
END SUB
' ========================================================================================

' ========================================================================================
' * Renders wireframe icosahedron (twenty-sided Platonic solid).
' The AfxGlutWireIcosahedron procedure renders an icosahedron whose corners are each a unit
' distance from the origin. The length of each side is slightly greater than one. Two of
' the corners lie on the positive and negative X-axes.
' ========================================================================================
SUB AfxGlutWireIcosahedron

   STATIC icos_r(11, 2) AS DOUBLE = { _
    {  1.0,             0.0,             0.0            }, _
    {  0.447213595500,  0.894427191000,  0.0            }, _
    {  0.447213595500,  0.276393202252,  0.850650808354 }, _
    {  0.447213595500, -0.723606797748,  0.525731112119 }, _
    {  0.447213595500, -0.723606797748, -0.525731112119 }, _
    {  0.447213595500,  0.276393202252, -0.850650808354 }, _
    { -0.447213595500, -0.894427191000,  0.0 }, _
    { -0.447213595500, -0.276393202252,  0.850650808354 }, _
    { -0.447213595500,  0.723606797748,  0.525731112119 }, _
    { -0.447213595500,  0.723606797748, -0.525731112119 }, _
    { -0.447213595500, -0.276393202252, -0.850650808354 }, _
    { -1.0,             0.0,             0.0            } _
    }

   STATIC icos_v(19, 2) AS LONG = { _
    {  0,  1,  2 }, _
    {  0,  2,  3 }, _
    {  0,  3,  4 }, _
    {  0,  4,  5 }, _
    {  0,  5,  1 }, _
    {  1,  8,  2 }, _
    {  2,  7,  3 }, _
    {  3,  6,  4 }, _
    {  4, 10,  5 }, _
    {  5,  9,  1 }, _
    {  1,  9,  8 }, _
    {  2,  8,  7 }, _
    {  3,  7,  6 }, _
    {  4,  6, 10 }, _
    {  5, 10,  9 }, _
    { 11,  9, 10 }, _
    { 11,  8,  9 }, _
    { 11,  7,  8 }, _
    { 11,  6,  7 }, _
    { 11, 10,  6 } _
    }

   FOR i AS LONG = 0 TO 19
      REDIM normal(2) AS DOUBLE
      normal(0) = (icos_r(icos_v(i,1),1) - icos_r(icos_v(i,0),1)) * _
                  (icos_r(icos_v(i,2),2) - icos_r(icos_v(i,0),2)) - _
                  (icos_r(icos_v(i,1),2) - icos_r(icos_v(i,0),2)) * _
                  (icos_r(icos_v(i,2),1) - icos_r(icos_v(i,0),1))

      normal(1) = (icos_r(icos_v(i,1),2) - icos_r(icos_v(i,0),2)) * _
                  (icos_r(icos_v(i,2),0) - icos_r(icos_v(i,0),0)) - _
                  (icos_r(icos_v(i,1),0) - icos_r(icos_v(i,0),0)) * _
                  (icos_r(icos_v(i,2),2) - icos_r(icos_v(i,0),2))

      normal(2) = (icos_r(icos_v(i,1),0) - icos_r(icos_v(i,0),0)) * _
                  (icos_r(icos_v(i,2),1) - icos_r(icos_v(i,0),1)) - _
                  (icos_r(icos_v(i,1),1) - icos_r(icos_v(i,0),1)) * _
                  (icos_r(icos_v(i,2),0) - icos_r(icos_v(i,0),0))
      glBegin(GL_LINE_LOOP)
         glNormal3dv(@normal(0))
         glVertex3dv(@icos_r(icos_v(i, 0), 0))
         glVertex3dv(@icos_r(icos_v(i, 1), 0))
         glVertex3dv(@icos_r(icos_v(i, 2), 0))
      glEnd
   NEXT

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a solid icosahedron (twenty-sided Platonic solid).
' The AfxGlutSolidIcosahedron procedure renders an icosahedron whose corners are each a unit
' distance from the origin. The length of each side is slightly greater than one. Two of
' the corners lie on the positive and negative X-axes.
' ========================================================================================
SUB AfxGlutSolidIcosahedron

   DIM i AS LONG

   STATIC icos_r(11, 2) AS DOUBLE = { _
    {  1.0,             0.0,             0.0            }, _
    {  0.447213595500,  0.894427191000,  0.0            }, _
    {  0.447213595500,  0.276393202252,  0.850650808354 }, _
    {  0.447213595500, -0.723606797748,  0.525731112119 }, _
    {  0.447213595500, -0.723606797748, -0.525731112119 }, _
    {  0.447213595500,  0.276393202252, -0.850650808354 }, _
    { -0.447213595500, -0.894427191000,  0.0 }, _
    { -0.447213595500, -0.276393202252,  0.850650808354 }, _
    { -0.447213595500,  0.723606797748,  0.525731112119 }, _
    { -0.447213595500,  0.723606797748, -0.525731112119 }, _
    { -0.447213595500, -0.276393202252, -0.850650808354 }, _
    { -1.0,             0.0,             0.0            } _
    }

   STATIC icos_v(19, 2) AS LONG = { _
    {  0,  1,  2 }, _
    {  0,  2,  3 }, _
    {  0,  3,  4 }, _
    {  0,  4,  5 }, _
    {  0,  5,  1 }, _
    {  1,  8,  2 }, _
    {  2,  7,  3 }, _
    {  3,  6,  4 }, _
    {  4, 10,  5 }, _
    {  5,  9,  1 }, _
    {  1,  9,  8 }, _
    {  2,  8,  7 }, _
    {  3,  7,  6 }, _
    {  4,  6, 10 }, _
    {  5, 10,  9 }, _
    { 11,  9, 10 }, _
    { 11,  8,  9 }, _
    { 11,  7,  8 }, _
    { 11,  6,  7 }, _
    { 11, 10,  6 } _
    }

   FOR i = 0 TO 19
      REDIM normal(2) AS DOUBLE
      normal(0) = (icos_r(icos_v(i,1),1) - icos_r(icos_v(i,0),1)) * _
                  (icos_r(icos_v(i,2),2) - icos_r(icos_v(i,0),2)) - _
                  (icos_r(icos_v(i,1),2) - icos_r(icos_v(i,0),2)) * _
                  (icos_r(icos_v(i,2),1) - icos_r(icos_v(i,0),1))

      normal(1) = (icos_r(icos_v(i,1),2) - icos_r(icos_v(i,0),2)) * _
                  (icos_r(icos_v(i,2),0) - icos_r(icos_v(i,0),0)) - _
                  (icos_r(icos_v(i,1),0) - icos_r(icos_v(i,0),0)) * _
                  (icos_r(icos_v(i,2),2) - icos_r(icos_v(i,0),2))

      normal(2) = (icos_r(icos_v(i,1),0) - icos_r(icos_v(i,0),0)) * _
                  (icos_r(icos_v(i,2),1) - icos_r(icos_v(i,0),1)) - _
                  (icos_r(icos_v(i,1),1) - icos_r(icos_v(i,0),1)) * _
                  (icos_r(icos_v(i,2),0) - icos_r(icos_v(i,0),0))
      glBegin(GL_TRIANGLES)
         glNormal3dv(@normal(0))
         glVertex3dv(@icos_r(icos_v(i, 0), 0))
         glVertex3dv(@icos_r(icos_v(i, 1), 0))
         glVertex3dv(@icos_r(icos_v(i, 2), 0))
      glEnd
   NEXT

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a wireframe rhombic dodecahedron (twelve-sided semi-regular solid).
' The AfxGlutWireRhombicDodecahedron procedure renders a rhombic dodecahedron whose corners
' are at most a distance of one from the origin. The rhombic dodecahedron has faces which
' are identical rhombuses (rhombi?) but which have some vertices at which three faces meet
' and some vertices at which four faces meet. The length of each side is sqrt(3)/2. Vertices
' at which four faces meet are found at (0, 0, +/- 1) and (+/- sqrt(2)/2, +/- sqrt(2)/2, 0).
' ========================================================================================
SUB AfxGlutWireRhombicDodecahedron

   STATIC rdod_r(13, 2) AS DOUBLE = { _
    {  0.0,             0.0,             1.0 }, _
    {  0.707106781187,  0.000000000000,  0.5 }, _
    {  0.000000000000,  0.707106781187,  0.5 }, _
    { -0.707106781187,  0.000000000000,  0.5 }, _
    {  0.000000000000, -0.707106781187,  0.5 }, _
    {  0.707106781187,  0.707106781187,  0.0 }, _
    { -0.707106781187,  0.707106781187,  0.0 }, _
    { -0.707106781187, -0.707106781187,  0.0 }, _
    {  0.707106781187, -0.707106781187,  0.0 }, _
    {  0.707106781187,  0.000000000000, -0.5 }, _
    {  0.000000000000,  0.707106781187, -0.5 }, _
    { -0.707106781187,  0.000000000000, -0.5 }, _
    {  0.000000000000, -0.707106781187, -0.5 }, _
    {  0.0,             0.0,            -1.0 } _
    }

   STATIC rdod_v(11, 3) AS DOUBLE = { _
    { 0,  1,  5,  2 }, _
    { 0,  2,  6,  3 }, _
    { 0,  3,  7,  4 }, _
    { 0,  4,  8,  1 }, _
    { 5, 10,  6,  2 }, _
    { 6, 11,  7,  3 }, _
    { 7, 12,  8,  4 }, _
    { 8,  9,  5,  1 }, _
    { 5,  9, 13, 10 }, _
    { 6, 10, 13, 11 }, _
    { 7, 11, 13, 12 }, _
    { 8, 12, 13,  9 } _
    }

   STATIC rdod_n(11, 2) AS DOUBLE = { _
    {  0.353553390594,  0.353553390594,  0.5 }, _
    { -0.353553390594,  0.353553390594,  0.5 }, _
    { -0.353553390594, -0.353553390594,  0.5 }, _
    {  0.353553390594, -0.353553390594,  0.5 }, _
    {  0.000000000000,  1.000000000000,  0.0 }, _
    { -1.000000000000,  0.000000000000,  0.0 }, _
    {  0.000000000000, -1.000000000000,  0.0 }, _
    {  1.000000000000,  0.000000000000,  0.0 }, _
    {  0.353553390594,  0.353553390594, -0.5 }, _
    { -0.353553390594,  0.353553390594, -0.5 }, _
    { -0.353553390594, -0.353553390594, -0.5 }, _
    {  0.353553390594, -0.353553390594, -0.5 } _
    }

   FOR i AS LONG = 0 TO 11
      glBegin(GL_LINE_LOOP)
         glNormal3dv(@rdod_n(0, 0))
         glVertex3dv(@rdod_r(rdod_v(i,0), 0))
         glVertex3dv(@rdod_r(rdod_v(i,1), 0))
         glVertex3dv(@rdod_r(rdod_v(i,2), 0))
         glVertex3dv(@rdod_r(rdod_v(i,3), 0))
      glEnd
   NEXT

END SUB
' ========================================================================================

' ========================================================================================
' * Draws a solid rhombic dodecahedron (twelve-sided semi-regular solid).
' The AfxGlutSolidRhombicDodecahedron procedure renders a rhombic dodecahedron whose corners
' are at most a distance of one from the origin. The rhombic dodecahedron has faces which
' are identical rhombuses (rhombi?) but which have some vertices at which three faces meet
' and some vertices at which four faces meet. The length of each side is sqrt(3)/2. Vertices
' at which four faces meet are found at (0, 0, +/- 1) and (+/- sqrt(2)/2, +/- sqrt(2)/2, 0).
' ========================================================================================
SUB AfxGlutSolidRhombicDodecahedron

   STATIC rdod_r(13, 2) AS DOUBLE = { _
    {  0.0,             0.0,             1.0 }, _
    {  0.707106781187,  0.000000000000,  0.5 }, _
    {  0.000000000000,  0.707106781187,  0.5 }, _
    { -0.707106781187,  0.000000000000,  0.5 }, _
    {  0.000000000000, -0.707106781187,  0.5 }, _
    {  0.707106781187,  0.707106781187,  0.0 }, _
    { -0.707106781187,  0.707106781187,  0.0 }, _
    { -0.707106781187, -0.707106781187,  0.0 }, _
    {  0.707106781187, -0.707106781187,  0.0 }, _
    {  0.707106781187,  0.000000000000, -0.5 }, _
    {  0.000000000000,  0.707106781187, -0.5 }, _
    { -0.707106781187,  0.000000000000, -0.5 }, _
    {  0.000000000000, -0.707106781187, -0.5 }, _
    {  0.0,             0.0,            -1.0 } _
    }

   STATIC rdod_v(11, 3) AS DOUBLE = { _
    { 0,  1,  5,  2 }, _
    { 0,  2,  6,  3 }, _
    { 0,  3,  7,  4 }, _
    { 0,  4,  8,  1 }, _
    { 5, 10,  6,  2 }, _
    { 6, 11,  7,  3 }, _
    { 7, 12,  8,  4 }, _
    { 8,  9,  5,  1 }, _
    { 5,  9, 13, 10 }, _
    { 6, 10, 13, 11 }, _
    { 7, 11, 13, 12 }, _
    { 8, 12, 13,  9 } _
    }

   STATIC rdod_n(11, 2) AS DOUBLE = { _
    {  0.353553390594,  0.353553390594,  0.5 }, _
    { -0.353553390594,  0.353553390594,  0.5 }, _
    { -0.353553390594, -0.353553390594,  0.5 }, _
    {  0.353553390594, -0.353553390594,  0.5 }, _
    {  0.000000000000,  1.000000000000,  0.0 }, _
    { -1.000000000000,  0.000000000000,  0.0 }, _
    {  0.000000000000, -1.000000000000,  0.0 }, _
    {  1.000000000000,  0.000000000000,  0.0 }, _
    {  0.353553390594,  0.353553390594, -0.5 }, _
    { -0.353553390594,  0.353553390594, -0.5 }, _
    { -0.353553390594, -0.353553390594, -0.5 }, _
    {  0.353553390594, -0.353553390594, -0.5 } _
    }

   FOR i AS LONG = 0 TO 11
      glBegin(GL_QUADS)
         glNormal3dv(@rdod_n(0, 0))
         glVertex3dv(@rdod_r(rdod_v(i,0), 0))
         glVertex3dv(@rdod_r(rdod_v(i,1), 0))
         glVertex3dv(@rdod_r(rdod_v(i,2), 0))
         glVertex3dv(@rdod_r(rdod_v(i,3), 0))
      glEnd
   NEXT

END SUB
' ========================================================================================

