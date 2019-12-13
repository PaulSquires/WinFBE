# TabControl (wfxTabControl)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllowDrop | Gets or sets a value (true/false) indicating whether the control will accept data that is dragged onto it.|
| AllowFocus | Gets or sets a value (true/false) indicating whether the control can receive keyboard focus.|
| BorderStyle | Gets or sets the border style of the control. Refer to the [ControlBorderStyle](#ControlBorderStyle) enum.|
| ButtonStyle | Gets or sets a value (true/false) indicating Tabs appear as buttons and no border is drawn around the display area of the Tab Control.|
| CtrlId | Gets or sets a value indicating the control ID of the control.|
| CtrlType | Gets or sets the control type value. Always **ControlType.TabControl** and used when adding control to the application’s form collection.|
| Enabled | Gets or sets a value (true/false) indicating whether the control can respond to user interaction.|
| FixedWidthTabs | Gets or sets a value (true/false) indicating that all tabs in a tab control are the same width (uses the TabWidth and TabHeight properties).|
| Focused | Gets or sets a value (true/false) indicating whether the control has input focus.|
| Font | Gets or sets the font for the control. Refer to the Font object.|
| ForceImageLeft | Gets or sets a value (true/false) indicating to force the tab image to the left, leaving the label centered in a tab control.|
| ForceLabelLeft | Gets or sets a value (true/false) indicating to force the tab image and tab label to the left in a tab control.|
| HotTracking | Gets or sets a value (true/false) indicating whether items under the mouse pointer are automatically highlighted.|
| Height | Gets or sets the height of the control.|
| hWindow |  Gets the Windows handle (hwnd) of the control. |
| hWindowParent | Gets or sets the Windows handle (hwnd) of the parent control. |
| MultiLine | Gets or sets a value (true/false) indicating that tabs can span more than one line.|
| TabImageSize | Gets or sets the size (in pixels) for images that will display in the tabs [ControlImageSize](#ControlImageSize) |
| TabHeight | Gets or sets the height of tabs in a tab control that have the property FixedWidthTabs equal True.|
| TabTopPadding | Gets or sets the amount of vertical padding around each tab's icon and label in a tab control.|
| TabSidePadding | Gets or sets the amount of horizontal padding around each tab's icon and label in a tab control.| 
| TabWeight | Gets or sets the width of tabs in a tab control that have the property FixedWidthTabs equal True.|
| TabPage |Gets or sets properties of a TabPage [wfxTabPage](#wfxTabPage) object |
| TabPages | Gets or sets properties of an TabPages [wfxTabPagesCollection](#wfxTabPagesCollection) collection object |
| Left | Gets or sets the distance, in pixels, between the left edge of the control and the left edge of its container's client area (normally the form).|
| Location |Gets or sets the top and left position of the control. Get: returns [wfxPoint](#wfxPoint) object. Set: (left, top). |
| Locked | Gets or sets a value (true/false) indicating whether the control can be moved or resized.|
| SelectedIndex | Gets or sets the index value of the selected tab in the tab control.|
| Size | Gets or sets the size of the form. Get: returns [wfxSize](#wfxSize) object. Set: (width, height)|
| TabIndex | Gets or sets the position that the control occupies in the TAB position.|
| TabStop | Gets or sets a value (true/false) indicating whether the user can use the TAB key to give focus to the control.|
| Tag | Gets or sets user defined text associated with the form.|
| Top | Gets or sets the distance, in pixels, between the top edge of the control and the top edge of its container's client area (normally the form).|
| Visible | Gets or sets a value (true/false) indicating whether the control is displayed.|
| Width | Gets or sets the width of the control.|

### Methods
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Hide | Conceals the control from the user.|
| Refresh | Forces the form to invalidate its client area and immediately redraw itself and any child controls.|
| SelectNextControl | Moves the input control to the next (True) or previous (False) control in the tab order. |
| SetBounds | Sets the bounds of the control to the specified location and size. SetBounds(left, top, width, height).|
| Show | Creates and displays the control to the user.|
| ShowPage | Show the specified page.|
| HidePage | Hide the specified page.|

### Events
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllEvents | Special handler where all events are routed through. Use this handler if you prefer to use the Win32 api style messages and wParam and lParam parameters. Set the Handled element of EventArgs to true if you handle a message and do not want Windows to perform any further processing on the message.|
| Click | Occurs when the client area of the control is clicked.|
| Destroy | Occurs immediately before the control is about to be destroyed and all resources associated with it released.|
| DropFiles | Occurs when an object is dragged and dropped onto the control and the AllowDrop property of the control is set to True.|
| GotFocus | Occurs when the control receives focus.|
| LostFocus | Occurs when the control loses focus.|
| KeyDown | Occurs when a key is pressed while the control has focus.|
| KeyPress | Occurs when a character, space or backspace key is pressed while the control has focus.|
| KeyUp | Occurs when a key is released while the control has focus.|
| MouseDoubleClick | Occurs when the control is double clicked by the mouse.|
| MouseDown | Occurs when the mouse pointer is over the control and a mouse button is pressed.|
| MouseEnter | Occurs when the mouse pointer enters the control.|
| MouseHover | Occurs when the mouse pointer rests on the control.|
| MouseLeave | Occurs when the mouse pointer leaves the control.|
| MouseMove | Occurs when the mouse pointer is moved over the control.|
| MouseUp | Occurs when the mouse pointer is over the control and a mouse button is released.|
| Selected | Occurs when a tab is selected.|
| Selecting | Occurs before a tab is selected, enabling a handler to cancel the tab change via setting e.Cancel = True or by simply returning TRUE from the event handler.|

### TabPages Collection (wfxTabPagesCollection)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Add | Add a new tab into the tab control.|
| ByIndex | Return the wfxTabPage object related to the specified tab control item index.|
| Clear | Removes all items from the tab control.|
| Count | Returns the total number of items in the tab control.|
| Inserts | Inserts a new tab into the tab control.|
| Remove | Remove/delete the item identified by the index value.|

### TabPage (wfxTabPage)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Index | The index value related to the tab page.|
| Selected | Gets or sets a value (true/false) indicating whether the tab page is selected.|
| Text | The text associated with this tab page.|
| ChildFormName | The Name of the form that is displayed in the tab control's client area.|
| hWndChildForm | The Windows Handle of the form that is displayed in the tab control's client area.|
| Image | The image name associated with this tab page.|
| Data32 | The optional 32bit value associated with this tab page.|


##### ControlImageSize
```
Enum ControlImageSize
   Size16 = 16
   Size24 = 24
   Size32 = 32
   Size48 = 48
End Enum
```
##### ControlBorderStyle
```
Enum ControlBorderStyle
   None	= 1
   FixedSingle	
   Fixed3D 
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
