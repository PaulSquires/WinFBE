# StatusBar (wfxStatusBar)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
|  ClickIndex | Gets the zero based index of the panel that was clicked.|
| CtrlId | Gets or sets a value indicating the control ID of the control.|
| CtrlType | Gets or sets the control type value. Always **ControlType.StatusBar** and used when adding control to the application’s form collection.|
| Height | Gets or sets the height of the control.|
| hWindow |  Gets the Windows handle (hwnd) of the control. |
| hWindowParent | Gets or sets the Windows handle (hwnd) of the parent control. |
| Panel |Gets or sets properties of an Panel [wfxStatusBarPanel](#wfxStatusBarPanel) object |
| Panels | Gets or sets properties of an Panels [wfxStatusBarPanelsCollection](#wfxStatusBarPanelsCollection) collection object |
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

### Panels Collection (wfxStatusBarPanelsCollection)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Add | Add a new item (and optional 32bit value) to the statusbar.|
| ByIndex | Return the wfxStatusBarPanel object related to the specified statusbar item index.|
| Clear | Removes all items from the statusbar.|
| Count | Returns the total number of items in the statusbar.|
| Remove | Remove/delete the item identified by the index value.|

### Panel (wfxStatusBarPanel)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Index | The index value related to the statusbar panel.|
| Alignment | Gets or sets a value indicating the panel's text alignment [StatusBarPanelAlignment](#StatusBarPanelAlignment) enum.|
| AutoSize | Gets or sets a value indicating the panel's sizing style [StatusBarPanelAutoSize](#StatusBarPanelAutoSize) enum. |
| BorderStyle | Gets or sets a value indicating the panel's border style [StatusBarPanelBorderStyle](#StatusBarPanelBorderStyle) enum.|
| Icon | Gets or sets the image name (as found in the Resource file) for any applicable icon image to display in the panel. |
| Text | The text associated with this panel item.|
| ToolTip | The tooltip text to display when the panel width is too narrow to display the full panel text.|
| Width | Gets or sets the maximum panel width. |
| MinWidth | Gets or sets the minimum panel width. |

##### StatusBarPanelAlignment
```
Enum StatusBarPanelAlignment
   Left   = ES_LEFT 
   Right  = ES_RIGHT
   Center = ES_CENTER
End Enum
```
##### StatusBarPanelAutoSize
```
Enum StatusBarPanelAutoSize
   None     = 0
   Contents = 1
   Spring   = 2
End Enum
```
##### StatusBarPanelBorderStyle
```
Enum StatusBarPanelBorderStyle
   Sunken = 0
   None   = SBT_NOBORDERS
   Raised = SBT_POPOUT
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

```