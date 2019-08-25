# PictureBox (wfxPictureBox)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllowDrop           | Gets or sets a value (true/false) indicating whether the control will accept data that is dragged onto it.        |
| BackColor | Gets or sets the background color of the control. Refer to the **Colors** object. |
| BackColorHot | Gets or sets the background color of the control when the mouse passes over it. Refer to the **Colors** object. |
| CtrlId | Gets or sets a value indicating the control ID of the control.|
| CtrlType | Gets or sets the control type value. Always **ControlType.PictureBox** and used when adding the control to the form's control collection.|
| Enabled | Gets or sets a value (true/false) indicating whether the control can respond to user interaction.|
| Height | Gets or sets the height of the control.|
| hWindow |  Gets the Windows handle (hwnd) of the control. |
| hWindowParent | Gets or sets the Windows handle (hwnd) of the parent control. |
| Image | Gets or sets the image to display in the control. The image is retrieved from the resource file for the project. Use the Image Manager to store you images specifying the RC_DATA resource type format. |
| ImageHeight | Gets or sets the height of the image. Use the value of zero (0) to allow the control to manage the height. |
| ImageWidth | Gets or sets the width of the image. Use the value of zero (0) to allow the control to manage the width. |
| ImageScaling | Gets or sets the scale used to apply to the image.Refer to the [ImageScale](#ImageScale) enum for values. |
| Left | Gets or sets the distance, in pixels, between the left edge of the control and the left edge of its container's client area (normally the form).|
| Location |Gets or sets the top and left position of the control. Get: returns [wfxPoint](#wfxPoint) object. Set: (left, top). |
| Locked | Gets or sets a value (true/false) indicating whether the control can be moved or resized.|
| Parent | Gets or sets the parent container of the control.|
| Size | Gets or sets the size of the form. Get: returns [wfxSize](#wfxSize) object. Set: (width, height)|
| Tag | Gets or sets user defined text associated with the control.|
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
| DropFiles | Occurs when an object is dragged and dropped onto the control and the AllowDrop property of the control is set to True.|
| MouseDoubleClick | Occurs when the control is double clicked by the mouse.|
| MouseDown | Occurs when the mouse pointer is over the control and a mouse button is pressed.|
| MouseEnter | Occurs when the mouse pointer enters the control.|
| MouseHover | Occurs when the mouse pointer rests on the control.|
| MouseLeave | Occurs when the mouse pointer leaves the control.|
| MouseMove | Occurs when the mouse pointer is moved over the control.|
| MouseUp | Occurs when the mouse pointer is over the control and a mouse button is released.|

##### ImageScale
```
enum ImageScale
   None = 1
   AutoSize
   FitWidth
   FitHeight   
   Stretch
end enum
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
