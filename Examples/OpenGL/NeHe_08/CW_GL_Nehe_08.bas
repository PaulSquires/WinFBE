' ########################################################################################
' Microsoft Windows
' File: CW_OGL_Nehe_08
' Contents: CWindow OpenGL - NeHe lesson 8
' Compiler: FreeBasic 32 & 64 bit
' Translated in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' Use ESC key to quit
' L key to toggle lighting (on/off)
' F key to cycle filters (Nearest, Linear, MipMapped)
' B key to toggle blending (on/off)
' PgUp key and PgDn key to Zoom in and out
' UpArrow, DnArrow, RightArrow, Left Arrow to change rotate speed of cube

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "GL/windows/glu.bi"
#INCLUDE ONCE "Afx/AfxGdiplus.inc"
USING Afx

CONST GL_WINDOWWIDTH   = 600               ' Window width
CONST GL_WINDOWHEIGHT  = 400               ' Window height
CONST GL_WindowCaption = "NeHe Lesson 8"   ' Window caption

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declarations
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

' =======================================================================================
' OpenGL class
' =======================================================================================
TYPE COGL

   Private:
      m_hDC AS HDC        ' // Device context handle
      m_hRC AS HGLRC      ' // Rendering context handle
      m_hwnd AS HWND      ' // Window handle

      TextureHandles(2) AS GLuint
      LightAmbient(3) AS SINGLE = {0.5, 0.5, 0.5, 1.0}
      LightDiffuse(3) AS SINGLE = {1.0, 1.0, 1.0, 1.0}
      LightPosition(3) AS SINGLE = {0.0, 0.0, 2.0, 1.0}

      xrot AS SINGLE
      yrot AS SINGLE
      zoom AS SINGLE
      filter AS LONG
      xspeed AS SINGLE
      yspeed AS SINGLE
      blend AS BOOLEAN

   Public:
      DECLARE CONSTRUCTOR (BYVAL hwnd AS HWND, BYVAL nBitsPerPel AS LONG = 32, BYVAL cDepthBits AS LONG = 24)
      DECLARE DESTRUCTOR
      DECLARE SUB SetupScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
      DECLARE SUB ResizeScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
      DECLARE SUB RenderScene
      DECLARE SUB ProcessKeystrokes (BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
      DECLARE SUB ProcessMouse (BYVAL uMsg AS UINT, BYVAL wKeyState AS DWORD, BYVAL x AS LONG, BYVAL y AS LONG)
      DECLARE SUB Cleanup

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

   ' // Load bitmap texture from disk
   DIM strTextureData AS STRING, TextureWidth AS LONG, TextureHeight AS LONG
   DIM hr AS LONG = AfxGdipLoadTexture("glass.bmp", TextureWidth, TextureHeight, strTextureData)
   ' // Assign an OpenGL handles to this texture
   glGenTextures 3, @TextureHandles(0)

   ' // Create Nearest Filtered Texture
   glBindTexture GL_TEXTURE_2D, TextureHandles(0)
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
   glTexImage2D GL_TEXTURE_2D, 0, 3, TextureWidth, TextureHeight, 0, _
                GL_RGBA, GL_UNSIGNED_BYTE, STRPTR(strTextureData)

   ' Create Linear Filtered Texture
   glBindTexture GL_TEXTURE_2D, TextureHandles(1)
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
   glTexImage2D GL_TEXTURE_2D, 0, 3, TextureWidth, TextureHeight, 0, _
                GL_RGBA, GL_UNSIGNED_BYTE, BYVAL STRPTR(strTextureData)

   ' Create MipMapped Texture
   glBindTexture GL_TEXTURE_2D, TextureHandles(2)
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR_MIPMAP_NEAREST
   gluBuild2DMipmaps GL_TEXTURE_2D, 3, TextureWidth, TextureHeight, _
                 GL_RGBA, GL_UNSIGNED_BYTE, BYVAL STRPTR(strTextureData)

   ' // Enable texture mapping
   glEnable GL_TEXTURE_2D

   ' // Specify clear values for the color buffers
   glClearColor 0.0, 0.0, 0.0, 0.0
   ' // Specify the clear value for the depth buffer
   glClearDepth 1.0
   ' // Specify the value used for depth-buffer comparisons
   glDepthFunc GL_LESS
   ' // Enable depth comparisons and update the depth buffer
   glEnable GL_DEPTH_TEST
   ' // Select smooth shading
   glShadeModel GL_SMOOTH

   glLightfv GL_LIGHT1, GL_AMBIENT, @LightAmbient(0)
   glLightfv GL_LIGHT1, GL_DIFFUSE, @LightDiffuse(0)
   glLightfv GL_LIGHT1, GL_POSITION, @LightPosition(0)
   glEnable GL_LIGHT1

   glColor4f 1.0, 1.0, 1.0, 0.5
   glBlendFunc  GL_SRC_ALPHA, GL_ONE
   zoom = -5.0
   blend = TRUE

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

   glTranslatef 0.0, 0.0, zoom

   glRotatef xrot, 1.0, 0.0, 0.0
   glRotatef yrot, 0.0, 1.0, 0.0

   glBindTexture GL_TEXTURE_2D, TextureHandles(filter)

   glBegin GL_QUADS
      ' Front Face
      glNormal3f   0.0, 0.0, 1.0
      glTexCoord2f 0.0, 0.0 : glVertex3f -1.0, -1.0,  1.0
      glTexCoord2f 1.0, 0.0 : glVertex3f  1.0, -1.0,  1.0
      glTexCoord2f 1.0, 1.0 : glVertex3f  1.0,  1.0,  1.0
      glTexCoord2f 0.0, 1.0 : glVertex3f -1.0,  1.0,  1.0
      ' Back Face
      glNormal3f   0.0, 0.0, -1.0
      glTexCoord2f 1.0, 0.0 : glVertex3f -1.0, -1.0, -1.0
      glTexCoord2f 1.0, 1.0 : glVertex3f -1.0,  1.0, -1.0
      glTexCoord2f 0.0, 1.0 : glVertex3f  1.0,  1.0, -1.0
      glTexCoord2f 0.0, 0.0 : glVertex3f  1.0, -1.0, -1.0
      ' Top Face
      glNormal3f   0.0, 1.0, 0.0
      glTexCoord2f 0.0, 1.0 : glVertex3f -1.0,  1.0, -1.0
      glTexCoord2f 0.0, 0.0 : glVertex3f -1.0,  1.0,  1.0
      glTexCoord2f 1.0, 0.0 : glVertex3f  1.0,  1.0,  1.0
      glTexCoord2f 1.0, 1.0 : glVertex3f  1.0,  1.0, -1.0
      ' Bottom Face
      glNormal3f   0.0,-1.0, 0.0
      glTexCoord2f 1.0, 1.0 : glVertex3f -1.0, -1.0, -1.0
      glTexCoord2f 0.0, 1.0 : glVertex3f  1.0, -1.0, -1.0
      glTexCoord2f 0.0, 0.0 : glVertex3f  1.0, -1.0,  1.0
      glTexCoord2f 1.0, 0.0 : glVertex3f -1.0, -1.0,  1.0
      ' Right face
      glNormal3f   1.0, 0.0, 0.0
      glTexCoord2f 1.0, 0.0 : glVertex3f  1.0, -1.0, -1.0
      glTexCoord2f 1.0, 1.0 : glVertex3f  1.0,  1.0, -1.0
      glTexCoord2f 0.0, 1.0 : glVertex3f  1.0,  1.0,  1.0
      glTexCoord2f 0.0, 0.0 : glVertex3f  1.0, -1.0,  1.0
      ' Left Face
      glNormal3f  -1.0,  0.0, 0.0
      glTexCoord2f 0.0, 0.0 : glVertex3f -1.0, -1.0, -1.0
      glTexCoord2f 1.0, 0.0 : glVertex3f -1.0, -1.0,  1.0
      glTexCoord2f 1.0, 1.0 : glVertex3f -1.0,  1.0,  1.0
      glTexCoord2f 0.0, 1.0 : glVertex3f -1.0,  1.0, -1.0
   glEnd

   xrot = xrot + xspeed
   yrot = yrot + yspeed

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

   STATIC AS BOOLEAN lp, fp, light

   SELECT CASE uMsg
      CASE WM_KEYDOWN   ' // A nonsystem key has been pressed
         SELECT CASE LOWORD(wParam)
            CASE VK_ESCAPE
               ' // Send a message to close the application
               SendMessageW m_hwnd, WM_CLOSE, 0, 0
            CASE VK_B
               blend = NOT blend
               IF blend = FALSE THEN
                  glEnable GL_BLEND
                  glDisable GL_DEPTH_TEST
               ELSE
                  glDisable GL_BLEND
                  glEnable GL_DEPTH_TEST
               END IF
            CASE VK_L
               light = NOT light
               IF light = FALSE THEN glDisable GL_LIGHTING ELSE glEnable GL_LIGHTING
            CASE VK_F
               filter += 1
               IF filter > 2 THEN filter = 0
            CASE VK_PRIOR   ' // page up
               zoom -= 0.02
            CASE VK_NEXT   ' // page down
               zoom += 0.02
            CASE VK_UP
               xspeed -= 0.01
            CASE VK_DOWN
               xspeed += 0.01
            CASE VK_LEFT
               yspeed -= 0.01
            CASE VK_RIGHT
               yspeed += 0.01
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

' =======================================================================================
' Cleanup
' =======================================================================================
SUB COGL.Cleanup

   ' // Delete the textures
   glDeleteTextures(3, @TextureHandles(0))

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
         ' // Clean resources
         pCOGL->Cleanup
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
         EXIT FUNCTION

   END SELECT

   ' // Default processing of Windows messages
   FUNCTION = DefWindowProcW(hwnd, uMsg, wParam, lParam)

END FUNCTION
' ========================================================================================
