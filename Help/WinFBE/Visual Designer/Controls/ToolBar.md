# ToolBar (wfxToolBar)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
|  ClickIndex | Gets the zero based index of the button that was clicked. |
|  ClickDropDownIndex | Gets the zero based index of the dropdown or whole dropdown button that was clicked. |
| CtrlId | Gets or sets a value indicating the control ID of the control.|
| CtrlType | Gets or sets the control type value. Always **ControlType.ToolBar** and used when adding control to the application’s form collection. |
| Height | Gets or sets the height of the control.|
| hWindow | Gets the Windows handle (hwnd) of the control (the toolbar tself). |
| hWindowParent | Gets or sets the Windows handle (hwnd) of the parent control. |
| Button |Gets or sets properties of an Button [wfxToolBarButton](#wfxToolBarButton) object |
| Buttons | Gets or sets properties of a Buttons's [wfxToolBarButtonsCollection](#wfxToolBarButtonsCollection) collection object |
| ButtonSize | Gets or sets the button size for all buttons in the toolbar (valid values are 16, 24, 32, 48). |
| Left | Gets or sets the distance, in pixels, between the left edge of the control and the left edge of its container's client area (normally the form).|
| Location |Gets or sets the top and left position of the control. Get: returns [wfxPoint](#wfxPoint) object. Set: (left, top). |
| Locked | Gets or sets a value (true/false) indicating whether the control can be moved or resized.|
| Parent | Gets or sets the parent container of the control.|
| Size | Gets or sets the size of the form. Get: returns [wfxSize](#wfxSize) object. Set: (width, height)|
| Tag | Gets or sets user defined text associated with the form.|
| Top | Gets or sets the distance, in pixels, between the top edge of the control and the top edge of its container's client area (normally the form).|
| Visible | Gets or sets a value (true/false) indicating whether the control is displayed.|
| Width | Gets or sets the width of the control.|

### Methods
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Hide | Conceals the control from the user.|
| Refresh | Forces the form to invalidate its client area and immediately redraw itself and any child controls.|
| SetBounds | Sets the bounds of the control to the specified location and size. SetBounds(left, top, width, height).|
| Show | Creates and displays the control to the user.|

### Events
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllEvents | Special handler where all events are routed through. Use this handler if you prefer to use the Win32 api style messages and wParam and lParam parameters. Set the Handled element of EventArgs to true if you handle a message and do not want Windows to perform any further processing on the message.|
| Click | Occurs when the client area of the control is clicked.|
| Destroy | Occurs immediately before the control is about to be destroyed and all resources associated with it released.|
| MouseDoubleClick | Occurs when the control is double clicked by the mouse.|
| MouseDown | Occurs when the mouse pointer is over the control and a mouse button is pressed.|
| MouseEnter | Occurs when the mouse pointer enters the control.|
| MouseHover | Occurs when the mouse pointer rests on the control.|
| MouseLeave | Occurs when the mouse pointer leaves the control.|
|MouseMove | Occurs when the mouse pointer is moved over the control.|
| MouseUp | Occurs when the mouse pointer is over the control and a mouse button is released.|

### Buttons Collection (wfxToolBarButtonsCollection)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Add | Add a new item (and optional 32bit value) to the toolbar.|
| ByIndex | Return the wfxToolBarButton object related to the specified toolbar item index.|
| Count | Returns the total number of items in the toolbar.|

### Button (wfxToolBarButton)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Index | The index value related to the statusbar panel.|
| ButtonType | Gets or sets a value indicating the button type [ToolBarButton](#ToolBarButton) enum.|
| NormalImage | Gets or sets the normal image name (as found in the Resource file). |
| HotImage | Gets or sets the hot image name (as found in the Resource file). |
| DisabledImage | Gets or sets the disabled image name (as found in the Resource file). |
| ToolTip | The tooltip text to display when the mouse hovers over the button.|

##### ToolBarButton
```
Enum ToolBarButton
   Button
   Separator
   DropDown
   WholeDropDown
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

```