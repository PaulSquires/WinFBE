' ########################################################################################
' Microsoft Windows
' File: CW_OGL_Nehe_17
' Contents: CWindow OpenGL - NeHe lesson 17
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

CONST GL_WINDOWWIDTH   = 600                ' Window width
CONST GL_WINDOWHEIGHT  = 400                ' Window height
CONST GL_WindowCaption = "NeHe Lesson 17"   ' Window caption

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

      m_TextureHandles(1) AS GLUint
      m_FontBase AS GLUint
      m_cnt1 AS SINGLE
      m_cnt2 AS SINGLE

   Public:
      DECLARE CONSTRUCTOR (BYVAL hwnd AS HWND, BYVAL nBitsPerPel AS LONG = 32, BYVAL cDepthBits AS LONG = 24)
      DECLARE DESTRUCTOR
      DECLARE SUB glPrint (BYVAL x AS LONG, BYVAL y AS LONG, BYREF szText AS ZSTRING, BYVAL nSet AS LONG, BYVAL FontBase AS GLUint, BYREF TextureHandle AS GLUint)
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

' ========================================================================================
SUB COGL.glPrint (BYVAL x AS LONG, BYVAL y AS LONG, BYREF szText AS ZSTRING, BYVAL nSet AS LONG, BYVAL FontBase AS GLUint, BYREF TextureHandle AS GLUint)

   IF nSet > 1 THEN nSet = 1

   glBindTexture GL_TEXTURE_2D, TextureHandle  ' Select our font texture
   glDisable GL_DEPTH_TEST                     ' Disables depth testing
   glMatrixMode GL_PROJECTION                  ' Select the projection matrix
   glPushMatrix                                ' Store the projection matrix
   glLoadIdentity                              ' Reset the projection matrix
   glOrtho 0, 640, 0, 480, -1, 1               ' Set up an ortho screen
   glMatrixMode GL_MODELVIEW                   ' Select the modelview matrix
   glPushMatrix                                ' Store the modelview matrix
   glLoadIdentity                              ' Reset the modelview matrix
   glTranslated x, y, 0                        ' Position the text (0,0 - bottom left)
   glListBase FontBase - 32 + (128 * nSet)     ' Choose the font set (0 or 1)
   glCallLists LEN(szText), GL_UNSIGNED_BYTE, @szText  ' Write the text to the screen
   glMatrixMode GL_PROJECTION                  ' Select the projection matrix
   glPopMatrix                                 ' Restore the old projection matrix
   glMatrixMode GL_MODELVIEW                   ' Select the modelview matrix
   glPopMatrix                                 ' Restore the old projection matrix
   glEnable GL_DEPTH_TEST                      ' Enables depth testing

END SUB
' ========================================================================================

' =======================================================================================
' All the setup goes here
' =======================================================================================
SUB COGL.SetupScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)

   ' // Load bitmap textures from disk
   DIM strTextureData AS STRING, TextureWidth AS LONG, TextureHeight AS LONG
   ' // Create two textures
   glGenTextures 2, @m_TextureHandles(0)
   ' // First texture
   AfxGdipLoadTexture("Font.bmp", TextureWidth, TextureHeight, strTextureData)
   glBindTexture GL_TEXTURE_2D, m_TextureHandles(0)
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
   glTexImage2D GL_TEXTURE_2D, 0, 3, TextureWidth, TextureHeight, 0, _
                GL_RGBA, GL_UNSIGNED_BYTE, STRPTR(strTextureData)
   ' // Second texture
   AfxGdipLoadTexture("Bumps.bmp", TextureWidth, TextureHeight, strTextureData)
   glBindTexture GL_TEXTURE_2D, m_TextureHandles(1)
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
   glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
   glTexImage2D GL_TEXTURE_2D, 0, 3, TextureWidth, TextureHeight, 0, _
                GL_RGBA, GL_UNSIGNED_BYTE, STRPTR(strTextureData)

   ' // Create 256 display lists
   m_FontBase = glGenLists(256)
   ' // Select our font texture
   glBindTexture GL_TEXTURE_2D, m_TextureHandles(0)
   FOR i AS LONG = 0 TO 255
      ' // x position of current character
      DIM cx AS SINGLE = (i MOD 16) / 16.0
      ' // y position of current character
      DIM cy AS SINGLE = (i \ 16) / 16.0
      ' // Start building a list
      glNewList m_FontBase + i, GL_COMPILE
         ' // Use a quad for each character
         glBegin GL_QUADS
            ' // Texture coordinate (bottom left)
            glTexCoord2f cx, 1 - cy - 0.0625
            ' // Vertex coordinate (bottom left)
            glVertex2i 0,0
            ' // Texture coordinate (bottom right)
            glTexCoord2f cx + 0.0625, 1 - cy - 0.0625
            ' // Vertex coordinate (bottom right)
            glVertex2i 16, 0
            ' // Texture coordinate (top right)
            glTexCoord2f cx + 0.0625, 1 - cy
            ' // Vertex coordinate (top right)
            glVertex2i 16, 16
            ' // Texture coordinate (top left)
            glTexCoord2f cx, 1 - cy
            ' // Vertex coordinate (top left)
            glVertex2i 0, 16
         ' // Done building our quad character
         glEnd
         ' // Move to the right of the character
         glTranslated 10,0,0
      ' // Done building the display list
      glEndList
   NEXT

   ' // Specify clear values for the color buffers
   glClearColor 0.0, 0.0, 0.0, 0.0
   ' // Enables clearing of the depth buffer
   glClearDepth 1.0
   ' // Type of depth test to do
   glDepthFunc GL_LEQUAL
   ' // Select the type of blending
   glBlendFunc GL_SRC_ALPHA, GL_ONE
   ' // Enables smooth color shading
   glShadeModel GL_SMOOTH
   ' // Enable 2D texture mapping
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

   ' // Clear the screen buffer
   glClear GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT
   ' // Reset the view
   glLoadIdentity

   glBindTexture GL_TEXTURE_2D, m_TextureHandles(1)  ' Select our second texture
   glTranslatef 0.0, 0.0, -5.0                       ' Move into the screen 5 units
   glRotatef 45.0, 0.0, 0.0, 1.0                     ' Rotate on the Z axis 45 degrees (clockwise)
   glRotatef m_cnt1 * 30.0, 1.0, 1.0, 0.0            ' Rotate on the X & Y axis by cnt1 (left to right)
   glDisable GL_BLEND                                ' Disable blending before we draw in 3d
   glColor3f 1.0, 1.0, 1.0                           ' Bright white
   glBegin GL_QUADS                                  ' Draw our first texture mapped quad
      glTexCoord2d 0.0,  0.0                         ' First texture coordinate
      glVertex2f  -1.0,  1.0                         ' First vertex
      glTexCoord2d 1.0,  0.0                         ' Second texture coordinate
      glVertex2f   1.0,  1.0                         ' Second vertex
      glTexCoord2d 1.0,  1.0                         ' Third texture coordinate
      glVertex2f   1.0, -1.0                         ' Third vertex
      glTexCoord2d 0.0,  1.0                         ' Fourth texture coordinate
      glVertex2f  -1.0, -1.0                         ' Fourth vertex
   glEnd                                             ' Done drawing the first quad
   glRotatef 90.0, 1.0, 1.0, 0.0                     ' Rotate on the x & y axis by 90 degrees (left to right)
   glBegin GL_QUADS                                  ' Draw our second texture mapped quad
      glTexCoord2d 0.0,  0.0                         ' First texture coordinate
      glVertex2f  -1.0,  1.0                         ' First vertex
      glTexCoord2d 1.0,  0.0                         ' Second texture coordinate
      glVertex2f   1.0,  1.0                         ' Second vertex
      glTexCoord2d 1.0,  1.0                         ' Third texture coordinate
      glVertex2f   1.0, -1.0                         ' Third vertex
      glTexCoord2d 0.0,  1.0                         ' Fourth texture coordinate
      glVertex2f  -1.0, -1.0                         ' Fourth vertex
   glEnd                                             ' Done drawing our second quad
   glEnable GL_BLEND                                 ' Enable blending

   ' // Reset the view
   glLoadIdentity

   ' // Pulsing colors based on text position
   glColor3f 1.0 * COS(m_cnt1),1.0 * SIN(m_cnt2), 1.0 - 0.5 * COS(m_cnt1 + m_cnt2)
   ' // Print GL text to the screen
   this.glPrint(280 + 250 * COS(m_cnt1), 235 + 200 * SIN(m_cnt2), "NeHe", 0, m_FontBase, m_TextureHandles(0))

   ' // Pulsing colors based on text position
   glColor3f 1.0 * SIN(m_cnt2), 1.0 - 0.5 * COS(m_cnt1 + m_cnt2), 1.0 * COS(m_cnt1)
   ' // Print GL text to the screen
   this.glPrint(280 + 230 * COS(m_cnt2), 235 + 200 * SIN(m_cnt1), "OpenGL", 1, m_FontBase, m_TextureHandles(0))

   ' // Set color to blue
   glColor3f 0.0, 0.0, 1.0
   ' // Print GL text to the screen
   this.glPrint(240 + 200 * COS((m_cnt2 + m_cnt1) / 5), 2, "Giuseppe D'Agata", 0, m_FontBase, m_TextureHandles(0))

   ' // Set color to white
   glColor3f 1.0, 1.0, 1.0
   ' // Print GL text to the screen
   this.glPrint(242 + 200 * COS((m_cnt2 + m_cnt1) / 5), 2, "Giuseppe D'Agata", 0, m_FontBase, m_TextureHandles(0))

   m_cnt1 = m_cnt1 + 0.01           ' Increase the first counter
   m_cnt2 = m_cnt2 + 0.0081         ' Increase the second counter

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

   STATIC light AS BOOLEAN

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

' =======================================================================================
' Cleanup
' =======================================================================================
SUB COGL.Cleanup

   ' // Delete the textures
   glDeleteTextures(2, @m_TextureHandles(0))
   ' // Delete the lists
   glDeleteLists m_FontBase, 256

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
