' ########################################################################################
' Microsoft Windows
' File: CW_GL_PrimitiveTypes.bas
' This sample demonstrates how to use the 10 primitive types available under OpenGL:
' points, lines, line strip, line loop, triangles triangle strips, triangle fan, quads,
' quads strips, and polygon.
' It is based on ogl_primitive_types.cpp, by Kevin Harris, 02/01/05.
' Compiler: FreeBasic 32 & 64 bit
' Adapted in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' Keyboard keys:
' F1 = Change the type
' F2 = Use a wire frame

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "GL/windows/glu.bi"
USING Afx

CONST GL_WINDOWWIDTH   = 600               ' Window width
CONST GL_WINDOWHEIGHT  = 400               ' Window height
CONST GL_WindowCaption = "OpenGL - Primitive types"   ' Window caption

' Vertex structure
TYPE Vertex
   r AS BYTE
   g AS BYTE
   b AS BYTE
   a AS BYTE
   x AS SINGLE
   y AS SINGLE
   z AS SINGLE
END TYPE

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

      DIM m_currentPrimitive AS LONG
      '// Note: We'll use the same vertices for both the line-strip and line-loop
      '// demonstrations. The difference between the two primitive types will not be
      '// noticeable until the vertices are rendered. Line-loop works exactly like
      '// line-strip except it creates a closed loop by automatically connecting
      '// the last vertex to the first with a line segment.
      DIM m_points (4) AS Vertex = { _
         (255, 0, 0, 255, 0.0, 0.0, 0.0), _
         (0, 255, 0, 255, 0.5, 0.0, 0.0), _
         (0, 0, 255, 255, -0.5, 0.0, 0.0), _
         (255, 255, 0, 255, 0.0, -0.5, 0.0), _
         (255, 0, 255, 255, 0.0, 0.5, 0.0)}
      DIM m_lines (5) AS Vertex = { _
         (255, 0, 0, 255, -1.0, 0.0, 0.0), _
         (255, 0, 0, 255, 0.0, 1.0, 0.0), _
         (0, 255, 0, 255, 0.5, 1.0, 0.0), _
         (0, 255, 0, 255, 0.5, -1.0, 0.0), _
         (0, 0, 255, 255, 1.0, -0.5, 0.0), _
         (0, 0, 255, 255, -1.0, -0.5, 0.0)}
      DIM m_lineStrip_and_lineLoop (5) AS Vertex = { _
         (255, 0, 0, 255, 0.5, 0.5, 0.0), _
         (0, 255, 0, 255, 1.0, 0.0, 0.0), _
         (0, 0, 255, 255, 0.0, -0.0, 0.0), _
         (255, 255, 0, 255, -1.0, 0.0, 0.0), _
         (255, 0, 0, 255, 0.0, 0.0, 0.0), _
         (255, 0, 0, 255, 0.0, 1.0, 0.0)}
      DIM m_triangles (5) AS Vertex = { _
         (255, 0, 0, 255, -1.0, 0.0, 0.0), _
         (0, 0, 255, 255, 1.0, 0.0, 0.0), _
         (0, 255, 0, 255, 0.0, 1.0, 0.0), _
         (255, 255, 0, 255, -0.5, 1.0, 0.0), _
         (255, 0, 0, 255, 0.5, -1.0, 0.0), _
         (0, 255, 255, 255, 0.0, -0.5, 0.0!)}
      DIM m_triangleStrip (7) AS Vertex = { _
         (255, 0, 0, 255, -2.0, 0.0, 0.0), _
         (0, 0, 255, 255, -1.0, 0.0, 0.0), _
         (0, 255, 0, 255, -1.0, 1.0, 0.0), _
         (255, 0, 255, 255, 0.0, 0.0, 0.0), _
         (255, 255, 0, 255, 0.0, 1.0, 0.0), _
         (255, 0, 0, 255, 1.0, 0.0, 0.0), _
         (0, 255, 255, 255, 1.0, 1.0, 0.0), _
         (0, 255, 0, 255, 2.0, 1.0, 0.0)}
      DIM m_triangleFan(5) AS Vertex = { _
         (255, 0, 0, 255, 0.0, -1.0, 0.0), _
         (0, 255, 255, 255, 1.0, 0.0, 0.0), _
         (255, 0, 255, 255, 0.5, 0.5, 0.0), _
         (255, 255, 0, 255, 0.0, 1.0, 0.0), _
         (0, 0, 255, 255, -0.5, 0.5, 0.0), _
         (0, 255, 0, 255, -1.0, 0.5, 0.0)}
      DIM m_quads (11) AS Vertex = { _
         (255, 0, 0, 255, -0.5, -0.5, 0.0), _
         (0, 255, 0, 255, 0.5, -0.5, 0.0), _
         (0, 0, 255, 255, 0.5, 0.5, 0.0), _
         (255, 255, 0, 255, -0.5, 0.5, 0.0), _
         (255, 0, 255, 255, -1.5, -1.5, 0.0), _
         (0, 255, 255, 255, -1.0, -1.0, 0.0), _
         (255, 0, 0, 255, -1.0, 1.5, 0.0), _
         (0, 255, 0, 255, -1.5, 1.5, 0.0), _
         (0, 0, 255, 255, 1.0, -0.2, 0.0), _
         (255, 255, 0, 255, 2.0, -0.2, 0.0), _
         (0, 255, 255, 255, 2.0, 0.2, 0.0), _
         (255, 0, 255, 255, 1.0, 0.2, 0.0)}
      DIM m_quadStrip (7) AS Vertex = { _
         (255, 0, 0, 255, -0.5, -1.5, 0.0), _
         (0, 255, 0, 255, 0.5, -1.5, 0.0), _
         (0, 0, 255, 255, -0.2, -0.5, 0.0), _
         (255, 255, 0, 255, 0.2, -0.5, 0.0), _
         (255, 0, 255, 255, -0.5, 0.5, 0.0), _
         (0, 255, 255, 255, 0.5, 0.5, 0.0), _
         (255, 0, 0, 255, -0.4, 1.5, 0.0), _
         (0, 255, 0, 255, 0.4, 1.5, 0.0)}
      DIM m_polygon (4) AS Vertex = { _
         (255, 0, 0, 255, -0.3, -1.5, 0.0), _
         (0, 255, 0, 255, 0.3, -1.5, 0.0), _
         (0, 0, 255, 255, 0.5, 0.5, 0.0), _
         (255, 255, 0, 255, 0.0, 1.5, 0.0), _
         (255, 0, 255, 255, -0.5, 0.5, 0.0)}

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

   m_currentPrimitive = GL_POLYGON
   glClearColor 0.0, 0.0, 0.0, 0.0
   glCullFace GL_BACK
   glEnable GL_CULL_FACE

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
   gluPerspective 45.0, nWidth / nHeight, 0.1, 100.0
   ' // Select the model view matrix
   glMatrixMode GL_MODELVIEW
   ' // Reset the model view matrix
   glLoadIdentity

END SUB
' =======================================================================================

' =======================================================================================
' Draw the scene
' =======================================================================================
SUB COGL.RenderScene

   ' // Clear the screen buffer
   glClear GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT
   ' // Reset the view
   glLoadIdentity

   glTranslatef 0.0, 0.0, -5.0       ' // Move into the screen 5.0

   SELECT CASE m_currentPrimitive
      CASE GL_POINTS
         glInterleavedArrays GL_C4UB_V3F, 0, @m_points(0)
         glDrawArrays GL_POINTS, 0, 5
      CASE GL_LINES
         glInterleavedArrays GL_C4UB_V3F, 0, @m_lines(0)
         glDrawArrays GL_LINES, 0, 6
      CASE GL_LINE_STRIP
         glInterleavedArrays GL_C4UB_V3F, 0, @m_lineStrip_and_lineLoop(0)
         glDrawArrays GL_LINE_STRIP, 0, 6
      CASE GL_LINE_LOOP
         glInterleavedArrays GL_C4UB_V3F, 0, @m_lineStrip_and_lineLoop(0)
         glDrawArrays GL_LINE_LOOP, 0, 6
      CASE GL_TRIANGLES
         glInterleavedArrays GL_C4UB_V3F, 0, @m_triangles(0)
         glDrawArrays GL_TRIANGLES, 0, 6
      CASE GL_TRIANGLE_STRIP
         glInterleavedArrays GL_C4UB_V3F, 0, @m_triangleStrip(0)
         glDrawArrays GL_TRIANGLE_STRIP, 0, 8
      CASE GL_TRIANGLE_FAN
         glInterleavedArrays GL_C4UB_V3F, 0, @m_triangleFan(0)
         glDrawArrays GL_TRIANGLE_FAN, 0, 6
      CASE GL_QUADS
         glInterleavedArrays GL_C4UB_V3F, 0, @m_quads(0)
         glDrawArrays GL_QUADS, 0, 12
      CASE GL_QUAD_STRIP
         glInterleavedArrays GL_C4UB_V3F, 0, @m_quadStrip(0)
         glDrawArrays GL_QUAD_STRIP, 0, 8
      CASE GL_POLYGON
         glInterleavedArrays GL_C4UB_V3F, 0, @m_polygon(0)
         glDrawArrays GL_POLYGON, 0, 5
   END SELECT

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

   STATIC bRenderInWireFrame AS BOOLEAN

   SELECT CASE uMsg
      CASE WM_KEYDOWN
         SELECT CASE LOWORD(wParam)
            CASE VK_ESCAPE
               ' // Send a message to close the application
               SendMessageW m_hwnd, WM_CLOSE, 0, 0
            CASE VK_F1
               IF m_currentPrimitive = GL_POINTS THEN
                  m_currentPrimitive = GL_LINES
               ELSEIF m_currentPrimitive = GL_LINES THEN
                  m_currentPrimitive = GL_LINE_STRIP
               ELSEIF m_currentPrimitive = GL_LINE_STRIP THEN
                  m_currentPrimitive = GL_LINE_LOOP
               ELSEIF m_currentPrimitive = GL_LINE_LOOP THEN
                  m_currentPrimitive = GL_TRIANGLES
               ELSEIF m_currentPrimitive = GL_TRIANGLES THEN
                  m_currentPrimitive = GL_TRIANGLE_STRIP
               ELSEIF m_currentPrimitive = GL_TRIANGLE_STRIP THEN
                  m_currentPrimitive = GL_TRIANGLE_FAN
               ELSEIF m_currentPrimitive = GL_TRIANGLE_FAN THEN
                  m_currentPrimitive = GL_QUADS
               ELSEIF m_currentPrimitive = GL_QUADS THEN
                  m_currentPrimitive = GL_QUAD_STRIP
               ELSEIF m_currentPrimitive = GL_QUAD_STRIP THEN
                  m_currentPrimitive = GL_POLYGON
               ELSEIF m_currentPrimitive = GL_POLYGON THEN
                  m_currentPrimitive = GL_POINTS
               END IF
            CASE VK_F2
               bRenderInWireFrame = NOT bRenderInWireFrame
               IF bRenderInWireFrame THEN
                  glPolygonMode GL_FRONT, GL_LINE
               ELSE
                  glPolygonMode GL_FRONT, GL_FILL
               END IF
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
