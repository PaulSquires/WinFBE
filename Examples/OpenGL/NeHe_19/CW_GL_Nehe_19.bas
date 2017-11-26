' ########################################################################################
' Microsoft Windows
' File: CW_OGL_Nehe_19
' Contents: CWindow OpenGL - NeHe lesson 19
' Compiler: FreeBasic 32 & 64 bit
' Translated in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################


#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "GL/windows/glu.bi"
#INCLUDE ONCE "Afx/AfxGdiplus.inc"
USING Afx

CONST GL_WINDOWWIDTH   = 600               ' Window width
CONST GL_WINDOWHEIGHT  = 400               ' Window height
CONST GL_WindowCaption = "NeHe Lesson 19"  ' Window caption

DECLARE FUNCTION WinMain (BYVAL hInstance AS HINSTANCE, _
                          BYVAL hPrevInstance AS HINSTANCE, _
                          BYVAL szCmdLine AS ZSTRING PTR, _
                          BYVAL nCmdShow AS LONG) AS LONG

   END WinMain(GetModuleHandleW(NULL), NULL, COMMAND(), SW_NORMAL)

' // Forward declarations
DECLARE FUNCTION WndProc (BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM) AS LRESULT

CONST MAX_PARTICLES = 1000   ' Number of particles to create

' =======================================================================================
' particles structure
' =======================================================================================
TYPE particles
  active AS LONG    ' Active (Yes/No)
  life AS SINGLE    ' Particle life
  fade AS SINGLE    ' Fade speed
  r AS SINGLE       ' Red value
  g AS SINGLE       ' Green value
  b AS SINGLE       ' Blue value
  x AS SINGLE       ' X Position
  y AS SINGLE       ' Y Position
  z AS SINGLE       ' Z Position
  xi AS SINGLE      ' X Direction
  yi AS SINGLE      ' Y Direction
  zi AS SINGLE      ' Z Direction
  xg AS SINGLE      ' X Gravity
  yg AS SINGLE      ' Y Gravity
  zg AS SINGLE      ' Z Gravity
END TYPE
' =======================================================================================

' =======================================================================================
' OpenGL class
' =======================================================================================
TYPE COGL

   Private:
      m_hDC AS HDC        ' // Device context handle
      m_hRC AS HGLRC      ' // Rendering context handle
      m_hwnd AS HWND      ' // Window handle

      m_TextureHandle AS GLUint
      m_rainbow AS BOOLEAN      ' Rainbow mode?
      m_slowdown AS SINGLE      ' Slow down particles
      m_xspeed AS SINGLE        ' Base x speed (to allow keyboard direction of tail)
      m_yspeed AS SINGLE        ' Base y speed (to allow keyboard direction of tail)
      m_zoom AS SINGLE          ' Used to zoom out
      m_clr AS GLuint           ' Current color selection
      m_delay AS GLuint         ' Rainbow effec delay
      m_sp AS BOOLEAN

      m_particle(MAX_PARTICLES - 1) AS particles
      m_colors(11, 2) AS SINGLE = { _
        { 1.0f,  0.5f,  0.5f}, _
	     { 1.0f,  0.75f, 0.5f}, _
	     { 1.0f,  1.0f,  0.5f}, _
	     { 0.75f, 1.0f,  0.5f}, _
        { 0.5f,  1.0f,  0.5f}, _
	     { 0.5f,  1.0f,  0.75f}, _
	     { 0.5f,  1.0f,  1.0f}, _
	     { 0.5f,  0.75f, 1.0f}, _
        { 0.5f,  0.5f,  1.0f}, _
	     { 0.75f, 0.5f,  1.0f}, _
	     { 1.0f,  0.5f,  1.0f}, _
	     { 1.0f,  0.5f,  0.75f} }

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

      ' // Get the device context of the window
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

   m_slowdown = 2.0
   m_zoom = -40.0

   ' // Load bitmap texture from disk
   DIM strTextureData AS STRING, TextureWidth AS LONG, TextureHeight AS LONG
   DIM hr AS LONG = AfxGdipLoadTexture("particle.bmp", TextureWidth, TextureHeight, strTextureData)

   ' // Assign handle
   glGenTextures 1, @m_TextureHandle

   ' // Create linear filtered texture
   glBindTexture GL_TEXTURE_2D, m_TextureHandle
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
   glTexImage2D GL_TEXTURE_2D, 0, 3, TextureWidth, TextureHeight, 0, _
                GL_RGBA, GL_UNSIGNED_BYTE, STRPTR(strTextureData)

   ' Select smooth shading
   glShadeModel GL_SMOOTH
   ' Specify clear values for the color buffers
   glClearColor 0.0, 0.0, 0.0, 0.0
   ' Disable depth testing
   glDisable GL_DEPTH_TEST
   ' Enable blending
   glEnable GL_BLEND
   ' Type of blending to perform
   glBlendFunc GL_SRC_ALPHA, GL_ONE
   ' Really nice perspective calculations
   glHint GL_PERSPECTIVE_CORRECTION_HINT, GL_NICEST
   ' Really nice point smoothing
   glHint GL_POINT_SMOOTH_HINT, GL_NICEST
   ' Enable texture maping
   glEnable GL_TEXTURE_2D
   ' Select texture
   glBindTexture GL_TEXTURE_2D, m_TextureHandle

   ' Initialize all the particles
   FOR i AS LONG = 0 TO MAX_PARTICLES -1
      m_particle(i).active = True                               ' Make all the particles active
      m_particle(i).life = 1.0                                  ' Give all the particles full life
      m_particle(i).fade = (RND * 100) / 1000.0 + 0.003         ' Random fade speed
      m_particle(i).r = m_colors(i * (12 \ MAX_PARTICLES), 0)   ' Select red rainbow color
      m_particle(i).g = m_colors(i * (12 \ MAX_PARTICLES), 1)   ' Select red rainbow color
      m_particle(i).b = m_colors(i * (12 \ MAX_PARTICLES), 2)   ' Select red rainbow color
      m_particle(i).xi = ((RND * 50) - 26.0) * 10.0             ' Random speed on x axis
      m_particle(i).yi = ((RND * 50) - 25.0) * 10.0             ' Random speed on y axis
      m_particle(i).zi = ((RND * 50) - 25.0) * 10.0             ' Random speed on z axis
      m_particle(i).xg = 0.0                                    ' Set horizontal pull to zero
      m_particle(i).yg = -0.8                                   ' Set vertical pull downward
      m_particle(i).zg = 0.0                                    ' Set pull on z axis to zero
   NEXT

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

   DIM AS SINGLE sx, sy, sz

   ' Clear the screen buffer
   glClear GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT
   ' Reset the view
   glLoadIdentity

   ' Loop through all the particles
   FOR i AS LONG = 0 TO MAX_PARTICLES - 1
      IF m_particle(i).active THEN           ' If the particle is active
         sx = m_particle(i).x                ' Grab our particle x position
         sy = m_particle(i).y                ' Grab our particle y position
         sz = m_particle(i).z + m_zoom       ' Particle z pos + zoom

         ' draw the particle using our rgb values, fade the particle based on it's life
         glColor4f m_particle(i).r, m_particle(i).g, m_particle(i).b, m_particle(i).life
         glBegin GL_TRIANGLE_STRIP                                ' build quad from a triangle strip
            glTexCoord2d 1, 1 : glVertex3f sx + 0.5, sy + 0.5, sz  ' Top right
            glTexCoord2d 0, 1 : glVertex3f sx - 0.5, sy + 0.5, sz  ' Top left
            glTexCoord2d 1, 0 : glVertex3f sx + 0.5, sy - 0.5, sz  ' Bottom right
            glTexCoord2d 0, 0 : glVertex3f sx - 0.5, sy - 0.5, sz  ' Bottom left
         glEnd                                                     ' Done building triangle strip

         m_particle(i).x = m_particle(i).x + m_particle(i).xi / (m_slowdown * 1000)  ' Move on the x axis by x speed
         m_particle(i).y = m_particle(i).y + m_particle(i).yi / (m_slowdown * 1000)  ' Move on the y axis by y speed
         m_particle(i).z = m_particle(i).z + m_particle(i).zi / (m_slowdown * 1000)  ' Move on the z axis by z speed

         m_particle(i).xi = m_particle(i).xi + m_particle(i).xg          ' Take pull on x axis into account
         m_particle(i).yi = m_particle(i).yi + m_particle(i).yg          ' Take pull on y axis into account
         m_particle(i).zi = m_particle(i).zi + m_particle(i).zg          ' Take pull on z axis into account
         m_particle(i).life = m_particle(i).life - m_particle(i).fade    ' Reduce particles life by 'fade'

         IF m_particle(i).life < 0.0 THEN                               ' If particle is burned out
            m_particle(i).life = 1.0                                    ' Give it new life
            m_particle(i).fade = (RND * 100) / 1000.0 + 0.003          ' Random fade value
            m_particle(i).x = 0.0                                       ' Center on x axis
            m_particle(i).y = 0.0                                       ' Center on y axis
            m_particle(i).z = 0.0                                       ' Center on z axis
            m_particle(i).xi = m_xspeed + ((RND * 60) - 32.0)           ' X axis speed and direction
            m_particle(i).yi = m_yspeed + ((RND * 60) - 30.0)           ' Y axis speed and direction
            m_particle(i).zi = ((RND * 60) - 30.0)                      ' Z axis speed and direction
            m_particle(i).r = m_colors(m_clr, 0)                         ' Select red from color table
            m_particle(i).g = m_colors(m_clr, 1)                         ' Select green from color table
            m_particle(i).b = m_colors(m_clr, 2)                         ' Select blue from color table
         END IF
      END IF
   NEXT

   IF (m_rainbow AND (m_delay > 25)) THEN
      m_sp = TRUE                   ' Set flag telling us space is pressed
      m_delay = 0                    ' Reset the rainbow color cycling delay
      m_clr = m_clr + 1              ' Change the particle color
      IF m_clr > 11 THEN m_clr = 0   ' If color is too high, reset it
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

   STATIC AS BOOLEAN light

   SELECT CASE uMsg
      CASE WM_KEYDOWN   ' // A nonsystem key has been pressed
         SELECT CASE LOWORD(wParam)
            CASE VK_ESCAPE
               ' // Send a message to close the application
               SendMessageW m_hwnd, WM_CLOSE, 0, 0
            CASE VK_SUBTRACT
               IF m_slowdown < 4.0 THEN m_slowdown = m_slowdown - 0.01
            CASE VK_RETURN
               m_rainbow = NOT m_rainbow
            ' Space or rainbow mode
            CASE VK_SPACE
               IF m_rainbow = TRUE AND m_delay > 25 THEN
                  m_rainbow = FALSE              ' Disable rainbow mode
                  m_sp = TRUE                    ' Set flag telling us space is pressed
                  m_delay = 0                    ' Reset the rainbow color cycling delay
                  m_clr = m_clr + 1              ' Change the particle color
                  IF m_clr > 11 THEN m_clr = 0   ' If color is too high, reset it
               END IF
            CASE VK_PRIOR   ' // page up
               m_zoom = m_zoom - 0.1
            CASE VK_NEXT   ' // page down
               m_zoom = m_zoom + 0.1
            CASE VK_UP
               IF m_yspeed < 200 THEN m_yspeed = m_yspeed + 1.0
            CASE VK_DOWN
               IF m_yspeed > - 200 THEN m_yspeed = m_yspeed - 1.0
            CASE VK_RIGHT
               IF m_xspeed < 200 THEN m_xspeed = m_xspeed + 1.0
            CASE VK_LEFT
               IF m_xspeed > -200 THEN m_xspeed = m_xspeed - 1.0
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

   ' // Delete the texture
   glDeleteTextures(1, @m_TextureHandle)

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
