# Form (wfxForm)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AcceptButton                | Gets or sets a reference to the button that receives a click message when Enter key is pressed.        |
| AllowDrop           | Gets or sets a value (true/false) indicating whether the control will accept data that is dragged onto it.        |
| BackColor | Gets or sets the background color of the form. Refer to the **Colors** object. |
| BorderStyle | Gets or sets the border style of the form. Refer to the [FormBorderStyle](#FormBorderStyle) enum. |
| CancelButton | Gets or sets a reference to the button that receives a click message when Escape key is pressed. |
| ChildForm | Creates form with WS_CHILD style and is used when specifying child Tab Pages for Tab Control.|
| ClientSize | Gets or sets the client area of the form.  The client area of the form is the size of the form excluding the borders and the title bar. Get: returns [wfxSize](#wfxSize) object. Set: (width, height) |
| ControlBox | Gets or sets value (true/false) indicating whether a control box is displayed in the caption bar of the form.|
| CtrlType | Gets or sets the control type value. Always **ControlType.Form** and used when adding form to the application’s form collection.|
| Enabled | Gets or sets a value (true/false) indicating whether the form can respond to user interaction.|
| Height | Gets or sets the height of the form.|
| hWindow |  Gets the Windows handle (hwnd) of the form. |
| hWindowParent | Gets or sets the Windows handle (hwnd) of the parent control or form. |
| Icon | Gets or sets the icon to display in the form’s system menu box. |
| IsMainForm | Gets or sets a value (true/false) indicating the form is main and will display when application starts. When the form is closed the application also ends. |
| IsModal |  Gets a value (true/false) indicating whether the form is displayed modally.|
| KeyPreview |  Gets or sets a value (true/false) indicating whether the form will receive key events before the event is passed to the control that has focus. |
| Left | Gets or sets the distance, in pixels, between the left edge of the form and the left edge of its container's client area (normally the screen).|
| Location |Gets or sets the top and left position of the form. Get: returns [wfxPoint](#wfxPoint) object. Set: (left, top). |
| Locked | Gets or sets a value (true/false) indicating whether the control can be moved or resized.|
| MainMenu | Gets a reference to any defined MainMenu object for the form. |
| MaximizeBox | Gets or sets a value (true/false) indicating whether the maximize button is displayed in the caption bar of the form.|
| MaximumHeight | Gets or sets the maximum height of the form.|
| MaximumWidth | Gets or sets the maximum width of the form.|
| MinimizeBox | Gets or sets a value (true/false) indicating whether the minimize button is displayed in the caption bar of the form.|
| MinimumHeight | Gets or sets the minimum height of the form.|
| MinimumWidth | Gets or sets the minimum width of the form.|
| Parent | Gets or sets the parent container of the form.|
| ScaleX | Gets the scaled value of a value based on the horizontal display resolution.|
| ScaleY | Gets the scaled value of a value based on the vertical display resolution.|
| Size | Gets or sets the size of the form. Get: returns [wfxSize](#wfxSize) object. Set: (width, height)|
| StartPosition | Gets or sets the starting position of the form at run time. Refer to [FormStartPosition](#FormStartPosition) enum.|
| StatusBar | Gets a reference to any defined StatusBar object for the form. |
|Tag | Gets or sets user defined text associated with the form.|
|Text | Gets or sets the text (caption) associated with this form.|
| ToolBar | Gets a reference to any defined ToolBar object for the form. |
| Top | Gets or sets the distance, in pixels, between the top edge of the form and the top edge of its container's client area (normally the screen).|
| Visible | Gets or sets a value (true/false) indicating whether the form is displayed.|
| Width | Gets or sets the width of the form.|
| WindowState | Gets or sets a value that indicates whether form is minimized, maximized, or normal. Refer to the [FormWindowState](#FormWindowState) enum.|

### Methods
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Close | Closes the form.|
| Hide | Conceals the form from the user.|
| Refresh | Forces the form to invalidate its client area and immediately redraw itself and any child controls.|
| SetBounds | Sets the bounds of the form to the specified location and size. SetBounds(left, top, width, height).|
| Show | Creates and displays the form to the user.|
| ShowDialog | Creates and shows the form as a modal dialog box.  |
| ShowChild | Creates and shows the form as a child (WS_CHILD) of the specified parent form similar to a Panel control. |

### Events
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Activated | Occurs when the form is activated in code or by the user.|
| AllEvents | Special handler where all events are routed through. Use this handler if you prefer to use the Win32 api style messages and wParam and lParam parameters. Set the Handled element of EventArgs to true if you handle a message and do not want Windows to perform any further processing on the message.|
| Click | Occurs when the client area of the form is clicked.|
| Deactivate | Occurs when the form loses focus and is no longer the active form.|
| DropFiles | Occurs when an object is dragged and dropped onto the control and the AllowDrop property of the control is set to True.|
| FormClosed | Occurs after the form is closed.|
| FormClosing | Occurs before the form is closed.|
| Initialize | Occurs before a user loads the form. The form TYPE variable is valid but any Window handles are not valid at this point.|
| KeyDown | Occurs when a key is pressed while the form has focus.|
| KeyPress | Occurs when a character, space or backspace key is pressed while the form has focus.|
| KeyUp | Occurs when a key is released while the form has focus.|
| Load | Occurs when the user loads the form and before it is displayed for the first time. Window handles are valid at this point.|
| MouseDoubleClick | Occurs when the form is double clicked by the mouse.|
| MouseDown | Occurs when the mouse pointer is over the form and a mouse button is pressed.|
| MouseEnter | Occurs when the mouse pointer enters the form.|
| MouseHover | Occurs when the mouse pointer rests on the form.|
| MouseLeave | Occurs when the mouse pointer leaves the form.|
|MouseMove | Occurs when the mouse pointer is moved over the form.|
| MouseUp | Occurs when the mouse pointer is over the form and a mouse button is released.|
| Move | Occurs when the form is moved.|
| Resize | Occurs when the form is resized.|
| Shown | Occurs whenever the form is first displayed.|



Form events in a specific order every time a form is created and shown. 

### During form creation
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Initialize | Occurs before a user loads the form. The form TYPE variable is valid but any Window handles are not valid at this point.|
| FormLoad | The form handle and all child controls exist however the form and controls are not yet visible. Respond to this event to reposition controls or to add data to controls. For example, add rows to a Listbox or Combobox. |
| FormActivated | The form has gained input focus (similar to the GotFocus event of a control).|
| Shown | This event is only raised the first time a form is displayed; subsequently minimizing, maximizing, restoring, hiding, showing, or invalidating and repainting will not raise this event. |

### During form destruction
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| FormClosing | Event occurs as the form is closing. If you cancel this event, then the form remains open. To cancel, simply set the Cancel element of the EventArgs structure to True.|
| Deactivate | Occurs when the form loses focus and is no longer the active form (similar to the LostFocus event of a control).|
| FormClosed | Occurs after the form has closed (similar to the Destroy event of a control).|



##### ImageLayout

```
Enum ImageLayout
  None = 1
  Tile
  Center
  Stretch
  Zoom
END ENUM 
```
##### FormBorderStyle
```
Enum FormBorderStyle
   None = 1
   Sizable 
   Fixed3D 
   FixedSingle	
   FixedDialog	
   FixedToolWindow 
   SizableToolWindow 
End Enum
```
##### FormWindowState
```
Enum FormWindowState
   Maximized = 1
   Minimized
   Normal
End Enum
```
##### FormStartPosition
```
Enum FormStartPosition
   CenterParent = 1
   CenterScreen
   Manual
   WindowsDefaultLocation
End Enum
```
##### wfxSize
```
Type wfxSize
   private:
      _Width  as Long
      _Height as long 

   public:
      Declare Property Width() As Long
      Declare Property Width( ByVal nValue As Long )
      Declare Property Height() As Long
      Declare Property Height( ByVal nValue As Long )
      Declare Function IsEmpty() as Boolean
      Declare Constructor ( byval nWidth as long = 0, byval nHeight as long = 0)
End Type
```
##### wfxPoint
```
Type wfxPoint
   private:
      _x as Long
      _y as long 

   public:
      Declare Property x() As Long
      Declare Property x( ByVal nValue As Long )
      Declare Property y() As Long
      Declare Property y( ByVal nValue As Long )
      Declare Function IsEmpty() as Boolean
      Declare Constructor ( byval xPos as long = 0, byval yPos as long = 0)
End Type
```
