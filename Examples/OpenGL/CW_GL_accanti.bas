' ########################################################################################
' Microsoft Windows
' File: CW_GL_accanti.bas
' Contents: CWindow OpenGL example
' Compiler: FreeBasic 32 & 64 bit
' Adapted in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

'/*
' * Copyright (c) 1993-1997, Silicon Graphics, Inc.
' * ALL RIGHTS RESERVED
' * Permission to use, copy, modify, and distribute this software for
' * any purpose and without fee is hereby granted, provided that the above
' * copyright notice appear in all copies and that both the copyright notice
' * and this permission notice appear in supporting documentation, and that
' * the name of Silicon Graphics, Inc. not be used in advertising
' * or publicity pertaining to distribution of the software without specific,
' * written prior permission.
' *
' * THE MATERIAL EMBODIED ON THIS SOFTWARE IS PROVIDED TO YOU "AS-IS"
' * AND WITHOUT WARRANTY OF ANY KIND, EXPRESS, IMPLIED OR OTHERWISE,
' * INCLUDING WITHOUT LIMITATION, ANY WARRANTY OF MERCHANTABILITY OR
' * FITNESS FOR A PARTICULAR PURPOSE.  IN NO EVENT SHALL SILICON
' * GRAPHICS, INC.  BE LIABLE TO YOU OR ANYONE ELSE FOR ANY DIRECT,
' * SPECIAL, INCIDENTAL, INDIRECT OR CONSEQUENTIAL DAMAGES OF ANY
' * KIND, OR ANY DAMAGES WHATSOEVER, INCLUDING WITHOUT LIMITATION,
' * LOSS OF PROFIT, LOSS OF USE, SAVINGS OR REVENUE, OR THE CLAIMS OF
' * THIRD PARTIES, WHETHER OR NOT SILICON GRAPHICS, INC.  HAS BEEN
' * ADVISED OF THE POSSIBILITY OF SUCH LOSS, HOWEVER CAUSED AND ON
' * ANY THEORY OF LIABILITY, ARISING OUT OF OR IN CONNECTION WITH THE
' * POSSESSION, USE OR PERFORMANCE OF THIS SOFTWARE.
' *
' * US Government Users Restricted Rights
' * Use, duplication, or disclosure by the Government is subject to
' * restrictions set forth in FAR 52.227.19(c)(2) or subparagraph
' * (c)(1)(ii) of the Rights in Technical Data and Computer Software
' * clause at DFARS 252.227-7013 and/or in similar or successor
' * clauses in the FAR or the DOD or NASA FAR Supplement.
' * Unpublished-- rights reserved under the copyright laws of the
' * United States.  Contractor/manufacturer is Silicon Graphics,
' * Inc., 2011 N.  Shoreline Blvd., Mountain View, CA 94039-7311.
' *
' * OpenGL(R) is a registered trademark of Silicon Graphics, Inc.
' */

'/*  accanti.c
' *  Use the accumulation buffer to do full-scene antialiasing
' *  on a scene with orthographic parallel projection.
' */

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxGlut.inc"
USING Afx

CONST GL_WINDOWWIDTH   = 600                  ' Window width
CONST GL_WINDOWHEIGHT  = 400                  ' Window height
CONST GL_WindowCaption = "OpenGL - accanti"   ' Window caption

' // jitter_point structure
TYPE jitter_point
  x AS SINGLE
  y AS SINGLE
END TYPE
CONST ACSIZE = 8

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declaration
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' =======================================================================================
' OpenGL class
' =======================================================================================
TYPE COGL

   Private:
      m_hDC AS HDC      ' // Device context handle
      m_hRC AS HGLRC    ' // Rendering context handle
      m_hwnd AS HWND    ' // Window handle

      DIM m_j8(7) AS jitter_point = { _
          (-0.334818,  0.435331), _
          ( 0.286438, -0.393495), _
          ( 0.459462,  0.141540), _
          (-0.414498, -0.192829), _
          (-0.183790,  0.082102), _
          (-0.079263, -0.317383), _
          ( 0.102254,  0.299133), _
          ( 0.164216, -0.054399) _
          }
      DIM mat_ambient(3) AS SINGLE = { 1.0, 1.0, 1.0, 1.0 }
      DIM mat_specular(3) AS SINGLE = { 1.0, 1.0, 1.0, 1.0 }
      DIM light_position(3) AS SINGLE = { 0.0, 0.0, 10.0, 1.0 }
      DIM lm_ambient(3) AS SINGLE = { 0.2, 0.2, 0.2, 1.0 }
      DIM torus_diffuse(3) AS SINGLE = { 0.7, 0.7, 0.0, 1.0 }
      DIM cube_diffuse(3) AS SINGLE = { 0.0, 0.7, 0.7, 1.0 }
      DIM sphere_diffuse(3) AS SINGLE = { 0.7, 0.0, 0.7, 1.0 }
      DIM octa_diffuse(3) AS SINGLE = { 0.7, 0.4, 0.4, 1.0 }

   Public:
      DECLARE CONSTRUCTOR (BYVAL hwnd AS HWND, BYVAL nBitsPerPel AS LONG = 32, BYVAL cDepthBits AS LONG = 24)
      DECLARE DESTRUCTOR
      DECLARE SUB SetupScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
      DECLARE SUB ResizeScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
      DECLARE SUB RenderScene
      DECLARE SUB ProcessKeystrokes (BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
      DECLARE SUB ProcessMouse (BYVAL uMsg AS UINT, BYVAL wKeyState AS DWORD, BYVAL x AS LONG, BYVAL y AS LONG)
      ' // Internal method
      DECLARE SUB displayObjects

END TYPE
' =======================================================================================

' ========================================================================================
' COGL constructor
' ========================================================================================
CONSTRUCTOR COGL (BYVAL hwnd AS HWND, BYVAL nBitsPerPel AS LONG = 32, BYVAL cDepthBits AS LONG = 24)

   DO   ' // Using a fake loop to avoid the use of GOTO or nested IFs/END IFs

      ' // Get the device context
      IF hwnd = NULL THEN EXIT DO
      m_hwnd = hwnd
      m_hDC = GetDC(m_hwnd)
      IF m_hDC = NULL THEN EXIT DO

      ' // Pixel format
      DIM pfd AS PIXELFORMATDESCRIPTOR
      pfd.nSize      = SIZEOF(PIXELFORMATDESCRIPTOR)
      pfd.nVersion   = 1
      pfd.dwFlags    = PFD_DRAW_TO_WINDOW OR PFD_SUPPORT_OPENGL OR PFD_DOUBLEBUFFER
      pfd.iPixelType = PFD_TYPE_RGBA
      pfd.cColorBits = nBitsPerPel
      pfd.cDepthBits = cDepthBits

      ' // Find a matching pixel format
      DIM pf AS LONG = ChoosePixelFormat(m_hDC, @pfd)
      IF pf = 0 THEN
         MessageBoxW hwnd, "Can't find a suitable pixel format", _
                     "Error", MB_OK OR MB_ICONEXCLAMATION
         EXIT DO
      END IF

      ' // Set the pixel format
      IF SetPixelFormat(m_hDC, pf, @pfd) = FALSE THEN
         MessageBoxW hwnd, "Can't set the pixel format", _
                     "Error", MB_OK OR MB_ICONEXCLAMATION
         EXIT DO
      END IF

      ' // Create a new OpenGL rendering context
      m_hRC = wglCreateContext(m_hDC)
      IF m_hRC = NULL THEN
         MessageBoxW hwnd, "Can't create an OpenGL rendering context", _
                     "Error", MB_OK OR MB_ICONEXCLAMATION
         EXIT DO
      END IF

      ' // Make it current
      IF wglMakeCurrent(m_hDC, m_hRC) = FALSE THEN
         MessageBoxW hwnd, "Can't activate the OpenGL rendering context", _
                     "Error", MB_OK OR MB_ICONEXCLAMATION
         EXIT DO
      END IF

      ' // Exit the fake loop
      EXIT DO
   LOOP

END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' COGL Destructor
' ========================================================================================
DESTRUCTOR COGL
   ' // Release the device and rendering contexts
   wglMakeCurrent m_hDC, NULL
   ' // Make the rendering context no longer current
   wglDeleteContext m_hRC
   ' // Release the device context
   ReleaseDC m_hwnd, m_hDC
END DESTRUCTOR
' ========================================================================================

' =======================================================================================
' All the setup goes here
' =======================================================================================
SUB COGL.SetupScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)

   glMaterialfv GL_FRONT, GL_AMBIENT, @mat_ambient(0)
   glMaterialfv GL_FRONT, GL_SPECULAR, @mat_specular(0)
   glMaterialf GL_FRONT, GL_SHININESS, 50.0
   glLightfv GL_LIGHT0, GL_POSITION, @light_position(0)
   glLightModelfv GL_LIGHT_MODEL_AMBIENT, @lm_ambient(0)

   glEnable GL_LIGHTING
   glEnable GL_LIGHT0
   glEnable GL_DEPTH_TEST
   glShadeModel GL_FLAT

   glClearColor 0.0, 0.0, 0.0, 0.0
   glClearAccum 0.0, 0.0, 0.0, 0.0

END SUB
' =======================================================================================

' =======================================================================================
' Resize the scene
' =======================================================================================
SUB COGL.ResizeScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)

   ' // Prevent divide by zero making height equal one
   IF nHeight = 0 THEN nHeight = 1
   ' // Reset the current viewport
   glViewport 0, 0, nWidth, nHeight
   ' // Select the projection matrix
   glMatrixMode GL_PROJECTION
   ' // Reset the projection matrix
   glLoadIdentity
   ' // Calculate the aspect ratio of the window
   IF nWidth <= nHeight THEN
      glOrtho -2.25, 2.25, -2.25 * nHeight / nWidth, 2.25 * nHeight / nWidth, -10.0, 10.0
   ELSE
      glOrtho -2.25 * nWidth / nHeight, 2.25 * nWidth / nHeight, -2.25, 2.25, -10.0, 10.0
   END IF
   ' // Select the model view matrix
   glMatrixMode GL_MODELVIEW
   ' // Reset the model view matrix
   glLoadIdentity

END SUB
' =======================================================================================

' =======================================================================================
' Auxiliary procedure
' =======================================================================================
SUB COGL.displayObjects

   glPushMatrix
   glRotatef 30.0, 1.0, 0.0, 0.0

   glPushMatrix
   glTranslatef -0.80, 0.35, 0.0
   glRotatef 100.0, 1.0, 0.0, 0.0
   glMaterialfv GL_FRONT, GL_DIFFUSE, @torus_diffuse(0)
   AfxGlutSolidTorus 0.275, 0.85, 16, 16
   glPopMatrix

   glPushMatrix
   glTranslatef -0.75, -0.50, 0.0
   glRotatef 45.0, 0.0, 0.0, 1.0
   glRotatef 45.0, 1.0, 0.0, 0.0
   glMaterialfv GL_FRONT, GL_DIFFUSE, @cube_diffuse(0)
   AfxGlutSolidCube 1.5
   glPopMatrix

   glPushMatrix
   glTranslatef 0.75, 0.60, 0.0
   glRotatef 30.0, 1.0, 0.0, 0.0
   glMaterialfv GL_FRONT, GL_DIFFUSE, @sphere_diffuse(0)
   AfxGlutSolidSphere 1.0, 16, 16
   glPopMatrix

   glPushMatrix
   glTranslatef 0.70, -0.90, 0.25
   glMaterialfv GL_FRONT, GL_DIFFUSE, @octa_diffuse(0)
   AfxGlutSolidOctahedron
   glPopMatrix

   glPopMatrix

END SUB
' =======================================================================================

' =======================================================================================
' Draw the scene
' =======================================================================================
SUB COGL.RenderScene

   DIM viewport(3) AS LONG
   glGetIntegerv GL_VIEWPORT, @viewport(0)

   glClear GL_ACCUM_BUFFER_BIT
   FOR jitter AS LONG = 0 TO ACSIZE - 1
      glClear GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT
      glPushMatrix
      ' // Note that 4.5 is the distance in world space between
      ' // left and right and bottom and top.
      ' // This formula converts fractional pixel movement to
      ' // world coordinates.
      glTranslatef m_j8(jitter).x * 4.5 / viewport(2), _
                   m_j8(jitter).y * 4.5 / viewport(3), 0.0
      this.displayObjects
      glPopMatrix
      glAccum GL_ACCUM, 1.0 / ACSIZE
   NEXT
   glAccum GL_RETURN, 1.0
   glFlush

   ' // Exchange the front and back buffers
   SwapBuffers m_hdc

END SUB
' =======================================================================================

' =======================================================================================
' Processes keystrokes
' Parameters:
' * uMsg = The Windows message
' * wParam = Additional message information.
' * lParam = Additional message information.
' The contents of the wParam and lParam parameters depend on the value of the uMsg parameter. 
' =======================================================================================
SUB COGL.ProcessKeystrokes (BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)

   SELECT CASE uMsg
      CASE WM_KEYDOWN   ' // A nonsystem key has been pressed
         SELECT CASE LOWORD(wParam)
            CASE VK_ESCAPE
               ' // Send a message to close the application
               SendMessageW m_hwnd, WM_CLOSE, 0, 0
         END SELECT
   END SELECT

END SUB
' =======================================================================================

' =======================================================================================
' Processes mouse clicks and movement
' Parameters:
' * wMsg      = Windows message
' * wKeyState = Indicates whether various virtual keys are down.
'               MK_CONTROL    The CTRL key is down.
'               MK_LBUTTON    The left mouse button is down.
'               MK_MBUTTON    The middle mouse button is down.
'               MK_RBUTTON    The right mouse button is down.
'               MK_SHIFT      The SHIFT key is down.
'               MK_XBUTTON1   Windows 2000/XP: The first X button is down.
'               MK_XBUTTON2   Windows 2000/XP: The second X button is down.
' * x         = x-coordinate of the cursor
' * y         = y-coordinate of the cursor
' =======================================================================================
SUB COGL.ProcessMouse (BYVAL uMsg AS UINT, BYVAL wKeyState AS DWORD, BYVAL x AS LONG, BYVAL y AS LONG)

   SELECT CASE uMsg

      CASE WM_LBUTTONDOWN   ' // Left mouse button pressed
         ' // Put your code here

      CASE WM_LBUTTONUP   ' // Left mouse button releases
         ' // Put your code here

      CASE WM_MOUSEMOVE   ' // Mouse has been moved
         ' // Put your code here

      END SELECT

END SUB
' =======================================================================================

' ========================================================================================
' Main
' ========================================================================================
FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                  BYVAL hPrevInstance AS HINSTANCE, _
                  BYVAL szCmdLine AS ZSTRING PTR, _
                  BYVAL nCmdShow AS LONG) AS LONG

   ' // Set process DPI aware
   ' // The recommended way is to use a manifest file
   AfxSetProcessDPIAware

   ' // Creates the main window
   DIM pWindow AS CWindow
   ' -or- DIM pWindow AS CWindow = "MyClassName" (use the name that you wish)
   ' // Create the window
   DIM hwndMain AS HWND = pWindow.Create(NULL, GL_WindowCaption, @WndProc)
   ' // Don't erase the background
   pWindow.ClassStyle = CS_DBLCLKS
   ' // Use a black brush
   pWindow.Brush = CreateSolidBrush(BGR(0, 0, 0))
   ' // Sizes the window by setting the wanted width and height of its client area
   pWindow.SetClientSize(GL_WINDOWWIDTH, GL_WINDOWHEIGHT)
   ' // Centers the window
   pWindow.Center

   ' // Show the window
   ShowWindow hwndMain, nCmdShow
   UpdateWindow hwndMain
 
   ' // Message loop
   DIM uMsg AS tagMsg
   WHILE GetMessageW(@uMsg, NULL, 0, 0)
      TranslateMessage @uMsg
      DispatchMessageW @uMsg
   WEND

   FUNCTION = uMsg.wParam

END FUNCTION
' ========================================================================================

' ========================================================================================
' Main window procedure
' ========================================================================================
FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

   STATIC pCOGL AS COGL PTR   ' // Pointer to the COGL class

   SELECT CASE uMsg

      CASE WM_SYSCOMMAND
         ' // Disable the Windows screensaver
         IF (wParam AND &hFFF0) = SC_SCREENSAVE THEN EXIT FUNCTION
         ' // Close the window
         IF (wParam AND &hFFF0) = SC_CLOSE THEN
            SendMessageW hwnd, WM_CLOSE, 0, 0
            EXIT FUNCTION
         END IF

      CASE WM_CREATE
         ' // Initialize the new OpenGL window
         pCOGL = NEW COGL(hwnd)
         ' // Retrieve the coordinates of the window's client area
         DIM rc AS RECT
         GetClientRect hwnd, @rc
         ' // Set up the scene
         pCOGL->SetupScene rc.Right - rc.Left, rc.Bottom - rc.Top
         ' // Set the timer (using a timer to trigger redrawing allows a smoother rendering)
         SetTimer(hwnd, 1, 0, NULL)
         EXIT FUNCTION

       CASE WM_DESTROY
         ' // Kill the timer
         KillTimer(hwnd, 1)
         ' // Delete the COGL class
         Delete pCOGL
         ' // Ends the application by sending a WM_QUIT message
         PostQuitMessage(0)
         EXIT FUNCTION

      CASE WM_TIMER
         ' // Render the scene
         pCOGL->RenderScene
         EXIT FUNCTION

      CASE WM_SIZE
         pCOGL->ResizeScene LOWORD(lParam), HIWORD(lParam)
         EXIT FUNCTION

      CASE WM_KEYDOWN
         ' // Process keystrokes
         pCOGL->ProcessKeystrokes uMsg, wParam, lParam
         EXIT FUNCTION

      CASE WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE
         ' // Process mouse movements
         pCOGL->ProcessMouse uMsg, wParam, LOWORD(lParam), HIWORD(lParam)

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================

