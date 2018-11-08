# OpenGL Procedures

Assorted helper OpenGL procedures.

Freeglut 2.8.1 includes twenty two routines for generating easily-recognizable 3-d geometric objects. These routines are effectively the same ones that are included in the GLUT library, and reflect the functionality available in the aux toolkit described in the OpenGL Programmer's Guide. They are included to allow programmers to create with a single line of code a three-dimensional object which can be used to test a variety of OpenGL functionality.  None of the routines generates a display list for the object which it draws. The functions generate normals appropriate for lighting but, except for the teapot functions, do not generate texture coordinates. Do note that depth testing (GL_LESS) should be enabled for the correct drawing of the nonconvex objects, i.e., the glutTorus and glutTeapot.

I have adapted these routines to FreeBasic and grouped them in the include file AfxGlut.inc

To load textures for use with OpenGL, both from disk or a resource file, you can use the overloaded wrapper function **AfxGdipLoadTexture**, included in the file AfxGdiPlus.inc.

| Name       | Description |
| ---------- | ----------- |
| [AfxGlutSolidCone](#AfxGlutCone) | Draws a solid cone. |
| [AfxGlutWireCone](#AfxGlutCone) | Draws a wireframe cone. |
| [AfxGlutSolidCube](#AfxGlutCube) | Draws a solid cube. |
| [AfxGlutWireCube](#AfxGlutCube) | Draws a wireframe cube. |
| [AfxGlutSolidCylinder](#AfxGlutCylinder) | Draws a solid cylinder. |
| [AfxGlutWireCylinder](#AfxGlutCylinder) | Draws a wireframe cylinder. |
| [AfxGlutSolidDodecahedron](#AfxGlutDodecahedron) | Draws a solid dodecahedron. |
| [AfxGlutWireDodecahedron](#AfxGlutDodecahedron) | Draws a wireframe dodecahedron. |
| [AfxGlutSolidIcosahedron](#AfxGlutIcosahedron) | Draws a solid icosahedron. |
| [AfxGlutWireIcosahedron](#AfxGlutIcosahedron) | Draws a wireframe icosahedron. |
| [AfxGlutSolidOctahedron](#AfxGlutOctahedron) | Draws a solid octahedron. |
| [AfxGlutWireOctahedron](#AfxGlutOctahedron) | Draws a wireframe octahedron. |
| [AfxGlutSolidRhombicDodecahedron](#AfxGlutRhombicDodecahedron) | Draws a solid rhombic dodecahedron. |
| [AfxGlutWireRhombicDodecahedron](#AfxGlutRhombicDodecahedron) | Draws a wireframe rhombic dodecahedron. |
| [AfxGlutSolidSphere](#AfxGlutSphere) | Renders a solid sphere centered at the origin of the modeling coordinate system. |
| [AfxGlutWireSphere](#AfxGlutSphere) | Renders a wireframe sphere centered at the origin of the modeling coordinate system. |
| [AfxGlutSolidTeapot](#AfxGlutTeapot) | Renders a solid teapot of the desired size, centered at the origin. |
| [AfxGlutWireTeapot](#AfxGlutTeapot) | Renders a wireframe teapot of the desired size, centered at the origin. |
| [AfxGlutSolidTetrahedron](#AfxGlutTetrahedron) | Renders a solid tetrahedron. |
| [AfxGlutWireTetrahedron](#AfxGlutTetrahedron) | Renders a wireframe tetrahedron. |
| [AfxGlutSolidTorus](#AfxGlutTorus) | Renders a solid torus (doughnut shape). |
| [AfxGlutWireTorus](#AfxGlutTorus) | Renders a wireframe torus (doughnut shape). |

# <a name="AfxGlutCone"></a>AfxGlutSolidCone / AfxGlutWireCone

Renders a right circular cone with a base centered at the origin and in the X-Y plane and its tip on the positive Z-axis. The wire cone is rendered with triangular elements.

```
SUB AfxGlutSolidCone (BYVAL dbase AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```
```
SUB AfxGlutWireCone (BYVAL dbase AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dbase* | The desired radius of the base of the cone. |
| *height* | The desired height of the cone. |
| *slices* | The desired number of slices around the cone. |
| *stacks* | The desired number of segments between the base and the tip of the cone (the number of points, including the tip, is stacks + 1). |

# <a name="AfxGlutCube"></a>AfxGlutSolidCube / AfxGlutWireCube

Renders a cube of the desired size, centered at the origin. Its faces are normal to the coordinate directions.

```
SUB AfxGlutSolidCube (BYVAL dbSize AS DOUBLE)
SUB AfxGlutWireCube (BYVAL dbSize AS DOUBLE)
```

| Parameter  | Description |
| ---------- | ----------- |
| *dbSize* | The desired length of an edge of the cube. |

# <a name="AfxGlutCylinder"></a>AfxGlutSolidCylinder / AfxGlutWireCylinder

Draws a cylinder.

```
SUB AfxGlutSolidCylinder (BYVAL radius AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```
```
SUB AfxGlutWireCylinder (BYVAL radius AS DOUBLE, BYVAL height AS DOUBLE, _
   BYVAL slices AS LONG, BYVAL stacks AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *radius* | The desired radius of the cylinder. |
| *height* | The desired height of the cylinder. |
| *slices* | The desired number of slices around the cylinder. |
| *stacks* | The desired number of segments between the base and the top of the cylinder (the number of points, including the tip, is stacks + 1). |

# <a name="AfxGlutDodecahedron"></a>AfxGlutSolidDodecahedron / AfxGlutWireDodecahedron

Renders a dodecahedron whose corners are each a distance of sqrt(3) from the origin. The length of each side is sqrt(5)-1. There are twenty corners; interestingly enough, eight of them coincide with the corners of a cube with sizes of length 2.

```
SUB AfxGlutSolidDodecahedron
SUB AfxGlutWireDodecahedron
```

# <a name="AfxGlutIcosahedron"></a>AfxGlutSolidIcosahedron / AfxGlutWireIcosahedron

Renders an icosahedron whose corners are each a unit distance from the origin. The length of each side is slightly greater than one. Two of the corners lie on the positive and negative X-axes.

```
SUB AfxGlutSolidIcosahedron
SUB AfxGlutWireIcosahedron
```

# <a name="AfxGlutOctahedron"></a>AfxGlutSolidOctahedron / AfxGlutWireOctahedron

Renders an octahedron whose corners are each a distance of one from the origin. The length of each side is sqrt(2). The corners are on the positive and negative coordinate axes.

```
SUB AfxGlutSolidOctahedron
SUB AfxGlutWireOctahedron
```

# <a name="AfxGlutRhombicDodecahedron"></a>AfxGlutSolidRhombicDodecahedron / AfxGlutWireRhombicDodecahedron

Renders a rhombic dodecahedron whose corners are at most a distance of one from the origin. The rhombic dodecahedron has faces which are identical rhombuses (rhombi?) but which have some vertices at which three faces meet and some vertices at which four faces meet. The length of each side is sqrt(3)/2. Vertices at which four faces meet are found at (0, 0, +/- 1) and (+/- sqrt(2)/2, +/- sqrt(2)/2, 0). 

```
SUB AfxGlutSolidRhombicDodecahedron
SUB AfxGlutWireRhombicDodecahedron
```

# <a name="AfxGlutSphere"></a>AfxGlutSolidSphere / AfxGlutWireSphere

Renders a sphere centered at the origin of the modeling coordinate system. The north and south poles of the sphere are on the positive and negative Z-axes respectively and the prime meridian crosses the positive X-axis.

```
SUB AfxGlutSolidSphere (BYVAL radius AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)
SUB AfxGlutWireSphere (BYVAL radius AS DOUBLE, BYVAL slices AS LONG, BYVAL stacks AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *radius* | The desired radius of the sphere. |
| *slices* | The desired number of slices (divisions in the longitudinal direction) in the sphere. |
| *stacks* | The desired number of stacks (divisions in the latitudinal direction) in the sphere. The number of points in this direction, including the north and south poles, is stacks+1. |

# <a name="AfxGlutTeapot"></a>AfxGlutSolidTeapot / AfxGlutWireTeapot

Renders a solid teapot of the desired size, centered at the origin.

| Parameter  | Description |
| ---------- | ----------- |
| *dbSize* | The desired size of the teapot. |

```
SUB AfxGlutSolidTeapot
SUB AfxGlutWireTeapot
```

# <a name="AfxGlutTetrahedron"></a>AfxGlutSolidTetrahedron / AfxGlutWireTetrahedron

Renders a tetrahedron whose corners are each a distance of one from the origin. The length of each side is 2/3 sqrt(6). One corner is on the positive X-axis and another is in the X-Y plane with a positive Y-coordinate.

| Parameter  | Description |
| ---------- | ----------- |
| *dbSize* | The desired size of the teapot. |

```
SUB AfxGlutSolidTetrahedron
SUB AfxGlutWireTetrahedron
```

# <a name="AfxGlutTorus"></a>AfxGlutSolidTorus / AfxGlutWireTorus

Renders a torus centered at the origin of the modeling coordinate system. The torus is circularly symmetric about the Z-axis and starts at the positive X-axis.

```
SUB AfxGlutSolidTorus (BYVAL innerRadius AS DOUBLE, BYVAL outerRadius AS DOUBLE, _
   BYVAL nsides AS LONG, BYVAL rings AS LONG)
```
```
SUB AfxGlutWireTorus (BYVAL innerRadius AS DOUBLE, BYVAL outerRadius AS DOUBLE, _
   BYVAL nsides AS LONG, BYVAL rings AS LONG)
```

| Parameter  | Description |
| ---------- | ----------- |
| *innerRadius* | Inner radius of the torus. |
| *outerRadius* | Outer radius of the torus. |
| *nsides* | Number of sides for each radial section. |
| *rings* | Number of radial divisions for the torus. |

# Example

The following example draws a teapot:

```
' ########################################################################################
' Microsoft Windows
' File: CW_GL_Teapot.bas
' Compiler: FreeBasic 32 & 64 bit
' Copyright (c) 2017 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#define UNICODE
#INCLUDE ONCE "Afx/CWindow.inc"
#INCLUDE ONCE "Afx/AfxGlut.inc"
USING Afx

CONST GL_WINDOWWIDTH   = 600                 ' Window width
CONST GL_WINDOWHEIGHT  = 400                 ' Window height
CONST GL_WindowCaption = "OpenGL - Teapot"   ' Window caption

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

   Public:
      m_hDC AS HDC      ' // Device context handle
      m_hRC AS HGLRC    ' // Rendering context handle
      m_hwnd AS HWND    ' // Window handle

   Private:
      DIM teapotList AS GLUint
      DIM ambient(3) AS SINGLE = { 0.0, 0.0, 0.0, 1.0 }
      DIM diffuse(3) AS SINGLE = { 1.0, 1.0, 1.0, 1.0 }
      DIM specular(3) AS SINGLE = { 1.0, 1.0, 1.0, 1.0 }
      DIM position(3) AS SINGLE = { 0.0, 3.0, 3.0, 0.0 }
      DIM lmodel_ambient(3) AS SINGLE = { 0.2, 0.2, 0.2, 1.0 }
      DIM local_view(0) AS SINGLE = { 0.0 }

   Public:
      DECLARE CONSTRUCTOR (BYVAL hwnd AS HWND, BYVAL nBitsPerPel AS LONG = 32, BYVAL cDepthBits AS LONG = 24)
      DECLARE DESTRUCTOR
      DECLARE SUB SetupScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
      DECLARE SUB ResizeScene (BYVAL nWidth AS LONG, BYVAL nHeight AS LONG)
      DECLARE SUB RenderScene
      DECLARE SUB ProcessKeystrokes (BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM)
      DECLARE SUB ProcessMouse (BYVAL uMsg AS UINT, BYVAL wKeyState AS DWORD, BYVAL x AS LONG, BYVAL y AS LONG)
      DECLARE SUB renderTeapot (BYVAL x AS SINGLE, BYVAL y AS SINGLE, _
         BYVAL ambr AS SINGLE, BYVAL ambg AS SINGLE, BYVAL ambb AS SINGLE, _
         BYVAL difr AS SINGLE, BYVAL difg AS SINGLE, BYVAL difb AS SINGLE, _
         BYVAL specr AS SINGLE, BYVAL specg AS SINGLE, BYVAL specb AS SINGLE, BYVAL shine AS SINGLE)

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

   glEnable(GL_LIGHTING)
   glEnable(GL_LIGHT0)
   glEnable(GL_DEPTH_TEST)

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
   ' // Reset the projection matrix
   glLoadIdentity

END SUB
' =======================================================================================

' =======================================================================================
' Draw the scene
' =======================================================================================
SUB COGL.RenderScene

   ' // Clear the screen buffer
   glClear GL_COLOR_BUFFER_BIT OR GL_DEPTH_BUFFER_BIT
   glPushMatrix
   gluLookAt(0,2.5,-7, 0,0,0, 0,1,0)
   AfxGlutSolidTeapot(2.5)
   glPopMatrix
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
```

The above example can also be used as a template for OpenGL programming with Windows. You simply have to change the code in the SetupScene, ResizeScene, RenderScene, ProcessKeystrokes and ProcessMouse methods, and the private instance variables.
