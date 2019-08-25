# ListBox (wfxListBox)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllowDrop | Gets or sets a value (true/false) indicating whether the control will accept data that is dragged onto it.|
| AllowSelection | Gets a value (true/false) indicating whether the control will allow items to be selected.|
| BackColor | Gets or sets the background color of the form. Refer to the **Colors** object. |
| BackColorHot | Gets or sets the background color of the item in the control that the mouse is over. Refer to the **Colors** object. |
| BackColorSelected | Gets or sets the background color of the selected/highlighted items in the control. Refer to the **Colors** object. |
| BorderStyle | Gets or sets the border style of the control. Refer to the [ControlBorderStyle](#ControlBorderStyle) enum.|
| ColumnWidth | Gets or sets a value indicating how wide (in pixels) each column should be in a multicolumn listbox.|
| CtrlId | Gets or sets a value indicating the control ID of the control.|
| CtrlType | Gets or sets the control type value. Always **ControlType.ListBox** and used when adding control to the application’s form collection.|
| Enabled | Gets or sets a value (true/false) indicating whether the control can respond to user interaction.|
| Focused | Gets or sets a value (true/false) indicating whether the control has input focus.|
| Font | Gets or sets the font for the control. Refer to the Font object.|
| ForeColor | Gets or sets the foreground color of the control. Refer to the **Colors** object. |
| ForeColorHot | Gets or sets the foreground color of the item in the control that the mouse is over. Refer to the **Colors** object. |
| ForeColorSelected | Gets or sets the foreground color of the selected/highlighted items in the control. Refer to the **Colors** object. |
| Height | Gets or sets the height of the control.|
| HorizontalExtent | Gets or sets the width (in pixels) by which a listbox can be scrolled horizontally. Only valid if the **HorizontalScrollBars** property is True.|
| HorizontalScrollBars | Gets or sets a value (true/false) indicating whether the ListBox will display a horizontal scroll bar for items beyond the right edge of the ListBox.|
| hWindow |  Gets the Windows handle (hwnd) of the control. |
| hWindowParent | Gets or sets the Windows handle (hwnd) of the parent control. |
| ItemHeight | Gets or sets the height of a line in the control.|
| IntegralHeight | Gets or sets a value (true/false) indicating whether the list can contain only complete items.|
| Item |Gets or sets properties of an Item [wfxListBoxItem](#wfxListBoxItem) object |
| Items | Gets or sets properties of an Items [wfxListBoxItemsCollection](#wfxListBoxItemsCollection) collection object |
| Left | Gets or sets the distance, in pixels, between the left edge of the control and the left edge of its container's client area (normally the form).|
| Location |Gets or sets the top and left position of the control. Get: returns [wfxPoint](#wfxPoint) object. Set: (left, top). |
| Locked | Gets or sets a value (true/false) indicating whether the control can be moved or resized.|
| MultiColumn | Gets or sets a value (true/false) indicating if values should be displayed in columns horizontally.|
| Parent | Gets or sets the parent container of the control.|
| ScrollAlwaysVisible | Gets or sets a value (true/false) indicating if the listbox should always have a scrollbar present regardless of how many items are in it.|
| SelectedItem | Gets or sets the selected item in thelistbox (using a [wfxListBoxItem](#wfxListBoxItem) object). |
| SelectedIndex | Gets or sets the index value of the selected item in the listbox.|
| SelectionMode | Gets or sets how the listbox handles selection of items. Refer to the [ListSelectionMode](#ListSelectionMode) enum.|
| Size | Gets or sets the size of the form. Get: returns [wfxSize](#wfxSize) object. Set: (width, height)|
| Sorted | Gets or sets a value (true/false) indicating if the listbox should be sorted.|
| TabIndex | Gets or sets the position that the control occupies in the TAB position.|
| TabStop | Gets or sets a value (true/false) indicating whether the user can use the TAB key to give focus to the control.|
| Tag | Gets or sets user defined text associated with the form.|
| TextAlign | Gets or sets a value indicating the alignment of the text on a control. Refer to [LabelAlignment](#LabelAlignment) enum. |
| Top | Gets or sets the distance, in pixels, between the top edge of the control and the top edge of its container's client area (normally the form).|
| TopIndex | Gets or sets the item index of the first visible item in the displayed list.|
| UseTabStops | Gets or sets a value (true/false) indicating whether the TAB character should be expanded into full spacing.|
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

### Items Collection (wfxListBoxItemsCollection)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Add | Add a new item (and optional 32bit value) to the listbox.|
| ByIndex | Return the wfxListBoxItem object related to the specified listbox item index.|
| Clear | Removes all items from the listbox.|
| Count | Returns the total number of items in the listbox.|
| Remove | Remove/delete the item identified by the index value.|
| SelectedCount | Returns the total number of selected items in the listbox.|

### Item (wfxListBoxItem)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Index | The index value related to the listbox item.|
| Selected | Gets or sets a value (true/false) indicating whether the Item is selected.|
| Text | The text associated with this listbox item.|
| Data32 | The optional 32bit value associated with this listbox item.|


## Working with a ListBox

Adding items (Text and 32-bit user defined value)
```
nIndex = Form1.ListBox1.Items.Add("First item in ListBox”, 12345)
```
Adding items (Text only)
```
nIndex = Form1.ListBox1.Items.Add("Second item in ListBox”)
```
Deleting an Item (the selected item)
```
Form1.ListBox1.Items.Remove(Form1.ListBox1.SelectedIndex)
```
Deleting an Item (based on index)
```
Form1.ListBox1.Items.Remove(0)  ‘ removes the first listbox item
```
Remove all Items from the ListBox
```
Form1.ListBox1.Items.Clear
```
Iterate all items in ListBox
```
For i as long = 0 to Form1.ListBox1.Items.Count – 1
   ? Form1.ListBox1.Item(i).Text, Form1.ListBox1.Item(i).Data32
Next
```
Iterate all selected items in the ListBox
```
' Determine the type of listbox selection method and then display items
select case Form1.ListBox1.SelectionMode
   case ListSelectionMode.None
      ' Obviously nothing in the list can be selected

   case ListSelectionMode.One
      ' Show the selected item (single select listbox)
      ? "Method #1: "; Form1.ListBox1.Item(Form1.ListBox1.SelectedIndex).Text, _
Form1.ListBox1.Item(Form1.ListBox1.SelectedIndex).Data32
      
      ' Alternate approach, show the selected item (single select listbox)
      ? "Method #2: "; Form1.ListBox1.SelectedItem.Text, _
                       Form1.ListBox1.SelectedItem.Data32

   Case ListSelectionMode.MultiSimple, ListSelectionMode.MultiExtended
      ' Show the selected items (multiselect listbox). Must iterate
      ' through the list and test if entry is selected.
      ? "Count: "; Form1.ListBox1.Items.Count
      ? "SelectedCount: "; Form1.ListBox1.Items.SelectedCount
      for i as long = 0 to Form1.ListBox1.Items.Count - 1  '<- Count not SelectedCount
         if Form1.ListBox1.Item(i).Selected then
            ? Form1.ListBox1.Item(i).Text, Form1.ListBox1.Item(i).Data32
         end if   
      next
End Select
```


##### ControlBorderStyle
```
Enum ControlBorderStyle
   None	= 1
   FixedSingle	
   Fixed3D 
End Enum
```
##### ListSelectionMode
```
Enum ListSelectionMode
   None = 1
   One
   MultiSimple
   MultiExtended
END ENUM
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
