' ########################################################################################
' Microsoft Windows
' File: CW_GL_IndexesGeometry.bas
' Demonstrates how to optimize performance by using indexed geometry. As a demonstration,
' the sample reduces the vertex count of a simple cube from 24 to 8 by redefining the
' cube’s geometry using an indices array.
' It is based on ogl_indexed_geometry.cpp, by Kevin Harris, 02/01/05.
' Compiler: FreeBasic 32 & 64 bit
' Adapted in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' Keyboard key:
' F1 = Toggle the use of indexed geometry or not
' Left mouse:
' Rotate the cube

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "GL/windows/glu.bi"
USING Afx

CONST GL_WINDOWWIDTH   = 600   ' Window width
CONST GL_WINDOWHEIGHT  = 400   ' Window height
CONST GL_WindowCaption = "OpenGL - Indexed geometry"   ' Window caption

' Vertex structure
TYPE Vertex
   r AS SINGLE
   g AS SINGLE
   b AS SINGLE
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

      m_bUseIndexedGeometry AS BOOLEAN = TRUE
      m_fSpinX AS SINGLE
      m_fSpinY AS SINGLE

      '// To understand how indexed geometry works, we must first build something
      '// which can be optimized through the use of indices.
      '//
      '// Below, is the vertex data for a simple multi-colored cube, which is defined
      '// as 6 individual quads, one quad for each of the cube's six sides. At first,
      '// this doesn’t seem too wasteful, but trust me it is.
      '//
      '// You see, we really only need 8 vertices to define a simple cube, but since
      '// we're using a quad list, we actually have to repeat the usage of our 8
      '// vertices 3 times each. To make this more understandable, I've actually
      '// numbered the vertices below so you can see how the vertices get repeated
      '// during the cube's definition.
      '//
      '// Note how the first 8 vertices are unique. Everting else after that is just
      '// a repeat of the first 8.
      DIM m_cubeVertices(23) AS Vertex = { _
         (1.0,0.0,0.0, -1.0,-1.0, 1.0), _  ' // 0 (unique)
         (0.0,1.0,0.0,  1.0,-1.0, 1.0), _  ' // 1 (unique)
         (0.0,0.0,1.0,  1.0, 1.0, 1.0), _  ' // 2 (unique)
         (1.0,1.0,0.0, -1.0, 1.0, 1.0), _  ' // 3 (unique)
         (1.0,0.0,1.0, -1.0,-1.0,-1.0), _  ' // 4 (unique)
         (0.0,1.0,1.0, -1.0, 1.0,-1.0), _  ' // 5 (unique)
         (1.0,1.0,1.0,  1.0, 1.0,-1.0), _  ' // 6 (unique)
         (1.0,0.0,0.0,  1.0,-1.0,-1.0), _  ' // 7 (unique)
         (0.0,1.0,1.0, -1.0, 1.0,-1.0), _  ' // 5 (start repeating here)
         (1.0,1.0,0.0, -1.0, 1.0, 1.0), _  ' // 3 (repeat of vertex 3)
         (0.0,0.0,1.0,  1.0, 1.0, 1.0), _  ' // 2 (repeat of vertex 2... etc.)
         (1.0,1.0,1.0,  1.0, 1.0,-1.0), _  ' // 6
         (1.0,0.0,1.0, -1.0,-1.0,-1.0), _  ' // 4
         (1.0,0.0,0.0,  1.0,-1.0,-1.0), _  ' // 7
         (0.0,1.0,0.0,  1.0,-1.0, 1.0), _  ' // 1
         (1.0,0.0,0.0, -1.0,-1.0, 1.0), _  ' // 0
         (1.0,0.0,0.0,  1.0,-1.0,-1.0), _  ' // 7
         (1.0,1.0,1.0,  1.0, 1.0,-1.0), _  ' // 6
         (0.0,0.0,1.0,  1.0, 1.0, 1.0), _  ' // 2
         (0.0,1.0,0.0,  1.0,-1.0, 1.0), _  ' // 1
         (1.0,0.0,1.0, -1.0,-1.0,-1.0), _  ' // 4
         (1.0,0.0,0.0, -1.0,-1.0, 1.0), _  ' // 0
         (1.0,1.0,0.0, -1.0, 1.0, 1.0), _  ' // 3
         (0.0,1.0,1.0, -1.0, 1.0,-1.0)}    ' // 5

      '// Now, to save ourselves the bandwidth of passing a bunch or redundant vertices
      '// down the graphics pipeline, we shorten our vertex list and pass only the
      '// unique vertices. We then create a indices array, which contains index values
      '// that reference vertices in our vertex array.
      '//
      '// In other words, the vertex array doens't actually define our cube anymore,
      '// it only holds the unique vertices; it's the indices array that now defines
      '// the cube's geometry.
      DIM m_cubeVertices_indexed(7) AS Vertex = { _
         (1.0,0.0,0.0,  -1.0,-1.0, 1.0), _  ' // 0
         (0.0,1.0,0.0,   1.0,-1.0, 1.0), _  ' // 1
         (0.0,0.0,1.0,   1.0, 1.0, 1.0), _  ' // 2
         (1.0,1.0,0.0,  -1.0, 1.0, 1.0), _  ' // 3
         (1.0,0.0,1.0,  -1.0,-1.0,-1.0), _  ' // 4
         (0.0,1.0,1.0,  -1.0, 1.0,-1.0), _  ' // 5
         (1.0,1.0,1.0,   1.0, 1.0,-1.0), _  ' // 6
         (1.0,0.0,0.0,   1.0,-1.0,-1.0)}    ' // 7

      DIM m_cubeIndices(23) AS DWORD = { 0, 1, 2, 3, 4, 5, 6, 7, 5, 3, 2, 6, 4, 7, 1, 0, 7, 6, 2, 1, 4, 0, 3, 5 }

      '// Note: While the cube above makes for a good example of how indexed geometry
      '//       works. There are many situations which can prevent you from using
      '//       an indices array to its full potential.
      '//
      '//       For example, if our cube required normals for lighting, things would
      '//       become problematic since each vertex would be shared between three
      '//       faces of the cube. This would not give you the lighting effect that
      '//       you really want since the best you could do would be to average the
      '//       normal's value between the three faces which used it.
      '//
      '//       Another example would be texture coordinates. If our cube required
      '//       unique texture coordinates for each face, you really wouldn’t gain
      '//       much from using an indices array since each vertex would require a
      '//       different texture coordinate depending on which face it was being
      '//       used in.

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

   ' Specify clear values for the color buffers
   glClearColor 0.0, 0.0, 0.0, 0.0
   ' Specify the clear value for the depth buffer
   glClearDepth 1.0
   ' Specify the value used for depth-buffer comparisons
   glDepthFunc GL_LESS
   ' Enable depth comparisons and update the depth buffer
   glEnable GL_DEPTH_TEST
   ' Select smooth shading
   glShadeModel GL_SMOOTH

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

   glTranslatef(0.0!, 0.0!, -5.0!)
   glRotatef(-m_fSpinY, 1.0!, 0.0!, 0.0!)
   glRotatef(-m_fSpinX, 0.0!, 0.0!, 1.0!)

   IF m_bUseIndexedGeometry THEN
      glInterleavedArrays(GL_C3F_V3F, 0, @m_cubeVertices_indexed(0))
      ' Note: The original C program uses GL_UNSIGNED_BYTE, but PB only supports
      ' signed bytes, so we are using an array of DWORDs
      glDrawElements(GL_QUADS, 24, GL_UNSIGNED_INT, @m_cubeIndices(0))
   ELSE
      glInterleavedArrays(GL_C3F_V3F, 0, @m_cubeVertices(0))
      glDrawArrays(GL_QUADS, 0, 24)
   END IF

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
               m_bUseIndexedGeometry = NOT m_bUseIndexedGeometry
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

   STATIC ptLastMousePosit AS POINT
   STATIC ptCurrentMousePosit AS POINT
   STATIC bMousing AS BOOLEAN

   SELECT CASE uMsg

      CASE WM_LBUTTONDOWN
         ptLastMousePosit.x = x
         ptCurrentMousePosit.x = x
         ptLastMousePosit.y = y
         ptCurrentMousePosit.y = y
         bMousing = TRUE

      CASE WM_LBUTTONUP
         bMousing = FALSE

      CASE WM_MOUSEMOVE
         ptCurrentMousePosit.x = x
         ptCurrentMousePosit.y = y
         IF bMousing THEN
            m_fSpinX -= (ptCurrentMousePosit.x - ptLastMousePosit.x)
            m_fSpinY -= (ptCurrentMousePosit.y - ptLastMousePosit.y)
         END IF
         ptLastMousePosit.x = ptCurrentMousePosit.x
         ptLastMousePosit.y = ptCurrentMousePosit.y

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
