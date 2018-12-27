' ########################################################################################
' Microsoft Windows
' File: CW_OGL_Nehe_20
' Contents: CWindow OpenGL - NeHe lesson 20
' Compiler: FreeBasic 32 & 64 bit
' Translated in 2017 by José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

' Space bar to toggle bitmaps
' M key to toggle masking

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "GL/windows/glu.bi"
#INCLUDE ONCE "Afx/AfxGdiplus.inc"
USING Afx

CONST GL_WINDOWWIDTH   = 600               ' Window width
CONST GL_WINDOWHEIGHT  = 400               ' Window height
CONST GL_WindowCaption = "NeHe Lesson 20"  ' Window caption

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

      m_TextureHandles(4) AS GLuint
      m_masking AS BOOLEAN         ' Masking on/off
      m_scene AS LONG              ' Which scene to draw
      m_roll AS SINGLE             ' Rolling texture

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

   m_masking = TRUE

   ' // Load textures from disk
   DIM strTextureData(4) AS STRING, TextureWidth(4) AS LONG, TextureHeight(4) AS LONG
   AfxGdipLoadTexture("logo.bmp", TextureWidth(0), TextureHeight(0), strTextureData(0))
   AfxGdipLoadTexture("mask1.bmp", TextureWidth(1), TextureHeight(1), strTextureData(1))
   AfxGdipLoadTexture("image1.bmp", TextureWidth(2), TextureHeight(2), strTextureData(2))
   AfxGdipLoadTexture("mask2.bmp", TextureWidth(3), TextureHeight(3), strTextureData(3))
   AfxGdipLoadTexture("image2.bmp", TextureWidth(4), TextureHeight(4), strTextureData(4))

   ' // Create five textures
   glGenTextures 5, @m_TextureHandles(0)
   FOR i AS LONG = 0 TO 4
      glBindTexture GL_TEXTURE_2D, m_TextureHandles(i)
      glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
      glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
      glTexImage2D GL_TEXTURE_2D, 0, 3, TextureWidth(i), TextureHeight(i), 0, _
                   GL_RGBA, GL_UNSIGNED_BYTE, STRPTR(strTextureData(i))
   NEXT

   ' Specify clear values for the color buffers
   glClearColor 0.0, 0.0, 0.0, 0.0
   ' Enables clearing of the depth buffer
   glClearDepth 1.0
   ' Enable depth testing
   glEnable GL_DEPTH_TEST
   ' Enables smooth color shading
   glShadeModel GL_SMOOTH
   ' Enable 2D texture mapping
   glEnable GL_TEXTURE_2D

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

   ' Clear the screen buffer
   glClear GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT
   ' Reset the view
   glLoadIdentity

   ' Move into the screen 5 units
   glTranslatef 0.0, 0.0, -2.0

   ' Select the logo texture
   glBindTexture GL_TEXTURE_2D, m_TextureHandles(0)
   ' Start drawing a textured quad
   glBegin GL_QUADS
      glTexCoord2f 0.0, - m_roll + 0.0 : glVertex3f -1.1, -1.1,  0.0  ' Bottom left
      glTexCoord2f 3.0, - m_roll + 0.0 : glVertex3f  1.1, -1.1,  0.0  ' Bottom right
      glTexCoord2f 3.0, - m_roll + 3.0 : glVertex3f  1.1,  1.1,  0.0  ' Top right
      glTexCoord2f 0.0, - m_roll + 3.0 : glVertex3f -1.1,  1.1,  0.0  ' Top left
   glEnd

   glEnable GL_BLEND                            ' Enable blending
   glDisable GL_DEPTH_TEST                      ' Disable depth testing

   IF m_masking THEN                            ' Is masking enabled?
      glBlendFunc GL_DST_COLOR, GL_ZERO         ' Blend screen color with zero (black)
   END IF

   IF m_scene THEN                              ' Are we drawing the second scene?
      glTranslatef 0.0, 0.0, -1.0               ' Translate into the screen one unit
      glRotatef m_roll * 360, 0.0, 0.0, 1.0     ' Rotate on the z axis 360 degrees
      IF m_masking THEN                         ' Is masking on?
         glBindTexture GL_TEXTURE_2D, m_TextureHandles(3)   ' Select the second mask texture
         glBegin GL_QUADS                       ' Start drawing a textured quad
            glTexCoord2f 0.0, 0.0 : glVertex3f -1.1, -1.1, 0.0   ' Bottom left
            glTexCoord2f 1.0, 0.0 : glVertex3f  1.1, -1.1, 0.0   ' Bottom right
            glTexCoord2f 1.0, 1.0 : glVertex3f  1.1,  1.1, 0.0   ' Top right
            glTexCoord2f 0.0, 1.0 : glVertex3f -1.1,  1.1, 0.0   ' Top left
         glEnd                                  ' Done Drawing The Quad
      END IF

      glBlendFunc GL_ONE, GL_ONE                         ' Copy image 2 color to the screen
      glBindTexture GL_TEXTURE_2D, m_TextureHandles(4)   ' Select the second image texture
      glBegin GL_QUADS                                   ' Start drawing a textured quad
         glTexCoord2f 0.0, 0.0 : glVertex3f -1.1, -1.1,  0.0     ' Bottom left
         glTexCoord2f 1.0, 0.0 : glVertex3f  1.1, -1.1,  0.0     ' Bottom right
         glTexCoord2f 1.0, 1.0 : glVertex3f  1.1,  1.1,  0.0     ' Top right
         glTexCoord2f 0.0, 1.0 : glVertex3f -1.1,  1.1,  0.0     ' Top left
      glEnd                                              ' Done Drawing The Quad
   ELSE
      IF m_masking THEN                                    ' Is masking on?
         glBindTexture GL_TEXTURE_2D, m_TextureHandles(1)  ' Select the first mask texture
         glBegin GL_QUADS                                  ' Start drawing a textured quad
            glTexCoord2f m_roll + 0.0, 0.0 : glVertex3f -1.1, -1.1, 0.0  ' Bottom left
            glTexCoord2f m_roll + 4.0, 0.0 : glVertex3f  1.1, -1.1, 0.0  ' Bottom right
            glTexCoord2f m_roll + 4.0, 4.0 : glVertex3f  1.1,  1.1, 0.0  ' Top right
            glTexCoord2f m_roll + 0.0, 4.0 : glVertex3f -1.1,  1.1, 0.0  ' Top left
         glEnd                                            ' Done drawing the quad
      END IF

      glBlendFunc GL_ONE, GL_ONE                         ' Copy image 1 color to the screen
      glBindTexture GL_TEXTURE_2D, m_TexTureHandles(2)   ' Select the first image texture
      glBegin GL_QUADS                                   ' Start drawing a textured quad
         glTexCoord2f m_roll + 0.0, 0.0 : glVertex3f -1.1, -1.1, 0.0  ' Bottom left
         glTexCoord2f m_roll + 4.0, 0.0 : glVertex3f  1.1, -1.1, 0.0  ' Bottom right
         glTexCoord2f m_roll + 4.0, 4.0 : glVertex3f  1.1,  1.1, 0.0  ' Top right
         glTexCoord2f m_roll + 0.0, 4.0 : glVertex3f -1.1,  1.1, 0.0  ' Top left
      glEnd                                               ' Done drawing the quad
   END IF

   glEnable GL_DEPTH_TEST                    ' Enable depth testing
   glDisable GL_BLEND                        ' Disable blending

   m_roll = m_roll + 0.002                   ' Increase our texture roll bariable
   IF m_roll > 1.0 THEN                      ' Is roll greater than one
      m_roll = m_roll - 1.0                  ' Subtract 1 from roll
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
            CASE VK_SPACE
               m_scene = NOT m_scene
            CASE VK_M
               m_masking = NOT m_masking
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
   glDeleteTextures(5, @m_TextureHandles(0))

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
