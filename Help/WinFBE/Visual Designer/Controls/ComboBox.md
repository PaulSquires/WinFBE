# ComboBox (wfxComboBox)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllowDrop | Gets or sets a value (true/false) indicating whether the control will accept data that is dragged onto it.|
| CtrlId | Gets or sets a value indicating the control ID of the control.|
| CtrlType | Gets or sets the control type value. Always **ControlType.ComboBox** and used when adding control to the application’s form collection.|
| DropDownStyle | Gets or sets the style of the control. Refer to the [ComboBoxStyle](#ComboBoxStyle) Enum |
| Enabled | Gets or sets a value (true/false) indicating whether the control can respond to user interaction.|
| Focused | Gets or sets a value (true/false) indicating whether the control has input focus.|
| Font | Gets or sets the font for the control. Refer to the Font object.|
| Height | Gets or sets the height of the control.|
| hWindow |  Gets the Windows handle (hwnd) of the control. |
| hWindowParent | Gets or sets the Windows handle (hwnd) of the parent control. |
| IntegralHeight | Gets or sets a value (true/false) indicating whether the list can contain only complete items.|
| Item |Gets or sets properties of an Item [wfxComboBoxItem](#wfxComboBoxItem) object |
| Items | Gets or sets properties of an Items [wfxComboBoxItemsCollection](#wfxComboBoxItemsCollection) collection object |
| Left | Gets or sets the distance, in pixels, between the left edge of the control and the left edge of its container's client area (normally the form).|
| Location |Gets or sets the top and left position of the control. Get: returns [wfxPoint](#wfxPoint) object. Set: (left, top). |
| Locked | Gets or sets a value (true/false) indicating whether the control can be moved or resized.|
| Parent | Gets or sets the parent container of the control.|
| SelectedItem | Gets or sets the selected item in the combobox (using a [wfxComboBoxItem](#wfxComboBoxItem) object). |
| SelectedIndex | Gets or sets the index value of the selected item in the combobox.|
| Size | Gets or sets the size of the form. Get: returns [wfxSize](#wfxSize) object. Set: (width, height)|
| Sorted | Gets or sets a value (true/false) indicating if the combobox should be sorted.|
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
|MouseMove | Occurs when the mouse pointer is moved over the control.|
| MouseUp | Occurs when the mouse pointer is over the control and a mouse button is released.|
| DropDown | Occurs when the drop-down portion of the combobox is shown.|
| DropDownClosed | Indicates that the drop-down portion of the combobox has closed.|
| TextChanged | Occurs when the Text property is changed by either a programmatic modification or user interaction. Only valid for DropDownStyle equal to ComboBoxStyle.Simple |

### Items Collection (wfxComboBoxItemsCollection)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Add | Add a new item (and optional 32bit value) to the combobox.|
| ByIndex | Return the wfxComboBoxItem object related to the specified combobox item index.|
| Clear | Removes all items from the combobox.|
| Count | Returns the total number of items in the combobox.|
| Remove | Remove/delete the item identified by the index value.|
| SelectedCount | Returns the total number of selected items in the combobox.|

### Item (wfxComboBoxItem)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Index | The index value related to the combobox item.|
| Selected | Gets or sets a value (true/false) indicating whether the Item is selected.|
| Text | The text associated with this combobox item.|
| Data32 | The optional 32bit value associated with this combobox item.|

## Working with a ComboBox

Adding items (Text and 32-bit user defined value)

```
nIndex = Form1.Combo1.Items.Add("First item in ComboBox”, 12345)
```

Adding items (Text only)

```
nIndex = Form1.Combo1.Items.Add("Second item in ComboBox”)
```

Deleting an Item (the selected item)

```
Form1.Combo1.Items.Remove(Form1.Combo1.SelectedIndex)
```

Deleting an Item (based on index)

```
Form1.Combo1.Items.Remove(0)  ‘ removes the first Combobox item
```

Remove all Items from the ComboBox

```
Form1.Combo1.Items.Clear
```

Iterate all items in ComboBox

```
For i as long = 0 to Form1.Combo.Items.Count – 1
   ? Form1.Combo1.Item(i).Text, Form1.Combo1.Item(i).Data32
Next
```

#####  ComboBoxStyle
```
Enum ComboBoxStyle
   Simple = 1
   DropDown
   DropDownList
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
