' ########################################################################################
' Microsoft Windows
' File: CW_GL_MoleHill.bas
' /* Copyright (c) Mark J. Kilgard, 1995. */
' /* This program is freely distributable without licensing fees
'    and is provided without guarantee or warrantee expressed or
'    implied. This program is -not- in the public domain. */
' /* molehill uses the GLU NURBS routines to draw some nice surfaces. */
' Compiler: FreeBasic 32 & 64 bit
' Adapted in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxGlut.inc"
USING Afx

CONST GL_WINDOWWIDTH   = 600               ' Window width
CONST GL_WINDOWHEIGHT  = 400               ' Window height
CONST GL_WindowCaption = "Mole Hill"       ' Window caption

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


   Public:
      DECLARE CONSTRUCTOR (BYVAL hwnd AS HWND, BYVAL nBitsPerPel AS LONG = 32, BYVAL cDepthBits AS LONG = 24)
      DECLARE DESTRUCTOR
      DECLARE SUB SetupScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
      DECLARE SUB ResizeScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
      DECLARE SUB RenderScene
      DECLARE SUB ProcessKeystrokes (BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
      DECLARE SUB ProcessMouse (BYVAL uMsg AS UINT, BYVAL wKeyState AS DWORD, BYVAL x AS LONG, BYVAL y AS LONG)

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

   DIM nurb AS GLUnurbs PTR
   DIM AS LONG u, v

   DIM mat_red_diffuse(3) AS SINGLE = {0.7, 0.0, 0.1, 1.0}
   DIM mat_green_diffuse(3) AS SINGLE = {0.0, 0.7, 0.1, 1.0}
   DIM mat_blue_diffuse(3) AS SINGLE = {0.0, 0.1, 0.7, 1.0}
   DIM mat_yellow_diffuse(3) AS SINGLE = {0.7, 0.8, 0.1, 1.0}
   DIM mat_specular(3) AS SINGLE = {1.0, 1.0, 1.0, 1.0}
   DIM mat_shininess(0) AS SINGLE = {100.0}
   DIM knots(7) AS SINGLE = {0.0, 0.0, 0.0, 0.0, 1.0, 1.0, 1.0, 1.0}

   DIM pts1(3, 3, 2) AS SINGLE
   DIM pts2(3, 3, 2) AS SINGLE
   DIM pts3(3, 3, 2) AS SINGLE
   DIM pts4(3, 3, 2) AS SINGLE

   glMaterialfv GL_FRONT, GL_SPECULAR, @mat_specular(0)
   glMaterialfv GL_FRONT, GL_SHININESS, @mat_shininess(0)
   glEnable GL_LIGHTING
   glEnable GL_LIGHT0
   glEnable GL_DEPTH_TEST
   glEnable GL_AUTO_NORMAL
   glEnable GL_NORMALIZE
   nurb = gluNewNurbsRenderer
   gluNurbsProperty nurb, GLU_SAMPLING_TOLERANCE, 25.0
   gluNurbsProperty nurb, GLU_DISPLAY_MODE, GLU_FILL

   ' /* Build control points for NURBS mole hills. */
   FOR u = 0 TO 3
      FOR v = 0 TO 3
         '/* Red. */
         pts1(u, v, 0) = 2.0 * u
         pts1(u, v, 1) = 2.0 * v
         IF  (u=1 OR u = 2) AND (v = 1 OR v = 2) THEN
            ' /* Stretch up middle. */
            pts1(u, v, 2) = 6.0
         ELSE
            pts1(u, v, 2) = 0.0
         END IF
         ' /* Green. */
         pts2(u, v, 0) = 2.0 * (u - 3.0)
         pts2(u, v, 1) = 2.0 * (v - 3.0)
         IF (u = 1 OR u = 2) AND (v = 1 OR v = 2) THEN
            IF u = 1 AND v = 1 THEN
               ' /* Pull hard on single middle square. */
               pts2(u, v, 2) = 15.0
            ELSE
               ' /* Push down on other middle squares. */
               pts2(u, v, 2) = -2.0
            END IF
         ELSE
            pts2(u, v, 2) = 0.0
         END IF
         ' /* Blue. */
         pts3(u, v, 0) = 2.0 * (u - 3.0)
         pts3(u, v, 1) = 2.0 * v
         IF (u=1 OR u = 2) AND (v = 1 OR v = 2) THEN
            IF u = 1 AND v = 2 THEN
               ' /* Pull up on single middple square. */
               pts3(u, v, 2) = 11.0
            ELSE
               ' /* Pull up slightly on other middle squares. */
               pts3(u, v, 2) = 2.0
            END IF
         ELSE
            pts3(u, v, 2) = 0.0
         END IF
         ' /* Yellow. */
         pts4(u, v, 0) = 2.0 * u
         pts4(u, v, 1) = 2.0 * (v - 3.0)
         IF (u=1 OR u = 2 OR u = 3) AND (v = 1 OR v = 2) THEN
            IF v = 1 THEN
               ' /* Push down front middle and right squares. */
               pts4(u, v, 2) = -2.0
            ELSE
               ' /* Pull up back middle and right squares. */
               pts4(u, v, 2) = 5.0
            END IF
         ELSE
            pts4(u, v, 2) = 0.0
         END IF
      NEXT
   NEXT

   ' /* Stretch up red's far right corner. */
   pts1(3, 3, 2) = 6
   ' /* Pull down green's near left corner a little. */
   pts2(0, 0, 2) = -2
   ' /* Turn up meeting of four corners. */
   pts1(0, 0, 2) = 1
   pts2(3, 3, 2) = 1
   pts3(3, 0, 2) = 1
   pts4(0, 3, 2) = 1

   glMatrixMode GL_PROJECTION
   gluPerspective 55.0, 1.0, 2.0, 24.0
   glMatrixMode GL_MODELVIEW
   glTranslatef 0.0, 0.0, -15.0
   glRotatef 330.0, 1.0, 0.0, 0.0

   glNewList 1, GL_COMPILE
   ' /* Render red hill. */
   glMaterialfv GL_FRONT, GL_DIFFUSE, @mat_red_diffuse(0)
   gluBeginSurface nurb
      gluNurbsSurface nurb, 8, @knots(0), 8, @knots(0), _
        4 * 3, 3, @pts1(0, 0, 0), _
        4, 4, GL_MAP2_VERTEX_3
   gluEndSurface nurb

   ' /* Render green hill. */
   glMaterialfv GL_FRONT, GL_DIFFUSE, @mat_green_diffuse(0)
   gluBeginSurface nurb
      gluNurbsSurface nurb, 8, @knots(0), 8, @knots(0), _
        4 * 3, 3, @pts2(0, 0, 0), _
        4, 4, GL_MAP2_VERTEX_3
   gluEndSurface nurb

   ' /* Render blue hill. */
   glMaterialfv GL_FRONT, GL_DIFFUSE, @mat_blue_diffuse(0)
   gluBeginSurface nurb
      gluNurbsSurface nurb, 8, @knots(0), 8, @knots(0), _
        4 * 3, 3, @pts3(0, 0, 0), _
        4, 4, GL_MAP2_VERTEX_3
   gluEndSurface nurb

   ' /* Render yellow hill. */
   glMaterialfv GL_FRONT, GL_DIFFUSE, @mat_yellow_diffuse(0)
   gluBeginSurface nurb
      gluNurbsSurface nurb, 8, @knots(0), 8, @knots(0), _
        4 * 3, 3, @pts4(0, 0, 0), _
        4, 4, GL_MAP2_VERTEX_3
   gluEndSurface nurb
   glEndList

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

END SUB
' =======================================================================================

' =======================================================================================
' Draw the scene
' =======================================================================================
SUB COGL.RenderScene

   ' // Clear the screen buffer
   glClear GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT

   ' // Call the list
   glCallList 1
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
