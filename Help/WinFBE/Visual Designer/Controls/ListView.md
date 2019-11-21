# ListView (wfxListView)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllowDrop | Gets or sets a value (true/false) indicating whether the control will accept data that is dragged onto it.|
| BorderStyle | Gets or sets the border style of the control. Refer to the [ControlBorderStyle](#ControlBorderStyle) enum.|
| CheckBoxes | Gets or sets a value indicating whether a check box appears next to each item in the control. |
| Column |Gets or sets properties of a Column [wfxListViewColumn](#wfxListViewColumn) object |
| Columns | Gets or sets properties of a Columns [wfxListViewColumnsCollection](#wfxListViewColumnsCollection) collection object |
| CtrlId | Gets or sets a value indicating the control ID of the control.|
| CtrlType | Gets or sets the control type value. Always **ControlType.ListView** and used when adding control to the application’s form collection.|
| Enabled | Gets or sets a value (true/false) indicating whether the control can respond to user interaction.|
| Focused | Gets or sets a value (true/false) indicating whether the control has input focus.|
| Font | Gets or sets the font for the control. Refer to the Font object.|
| FullRowSelect | Gets or sets a value indicating whether clicking an item selects all its subitems. |
| GridLines | Gets or sets a value indicating whether grid lines appear between the rows and columns containing the items and subitems in the control. |
| HeaderStyle | Gets or sets the column header style. |
| HeaderHeight | Gets or sets the column header height |
| HeaderBackColor | The backgroundground color of the ListView column display text. |
| HeaderForeColor | The foreground color of the ListView column display text. |
| Height | Gets or sets the height of the control.|
| HideSelection  | Indicates that the selection should be hidden when the edit control loses focus. |
| hWindow |  Gets the Windows handle (hwnd) of the control. |
| hWindowParent | Gets or sets the Windows handle (hwnd) of the parent control. |
| Item |Gets or sets properties of an Item [wfxListViewItem](#wfxListViewItem) object |
| Items | Gets or sets properties of an Items [wfxListViewItemsCollection](#wfxListViewItemsCollection) collection object |
| Left | Gets or sets the distance, in pixels, between the left edge of the control and the left edge of its container's client area (normally the form).|
| Location |Gets or sets the top and left position of the control. Get: returns [wfxPoint](#wfxPoint) object. Set: (left, top). |
| Locked | Gets or sets a value (true/false) indicating whether the control can be moved or resized.|
| MultiSelec  | Gets or sets a value indicating whether multiple items can be selected. |
| Parent | Gets or sets the parent container of the control.|
| SelectedItem | Gets or sets the selected item in thelistbox (using a [wfxListViewItem](#wfxListViewItem) object). |
| SelectedIndex | Gets or sets the index value of the selected item in the listview.|
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
| BeginUpdate | Prevents control updating during adding/removing of a bulk number of items |
| EndUpdate | Enables control updating after adding/removing of a bulk number of items |
| Hide | Conceals the control from the user.|
| HitTest | Determines what (if any) ListView Item and/or SubItem has been clicked on. HitTest( iItem, iSubItem ) |
| Refresh | Forces the form to invalidate its client area and immediately redraw itself and any child controls.|
| SelectNextControl | Moves the input control to the next (True) or previous (False) control in the tab order. |
| SetBounds | Sets the bounds of the control to the specified location and size. SetBounds(left, top, width, height).|
| Show | Creates and displays the control to the user.|

### Events
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllEvents | Special handler where all events are routed through. Use this handler if you prefer to use the Win32 api style messages and wParam and lParam parameters. Set the Handled element of EventArgs to true if you handle a message and do not want Windows to perform any further processing on the message.|
| Click | Occurs when the client area of the control is clicked. e.ListViewRow contains clicked Item index. e.ListViewColumn contains clicked SubItem index. This event will also fire if ENTER is pressed.|
| ColumnClick | Occurs when a column header is clicked. e.ListViewColumn contains index of clicked Column header.|
| DoubleClick | Occurs when the client area of the control is double clicked. e.ListViewRow contains clicked Item index. e.ListViewColumn contains clicked SubItem index.|
| RightClick | Occurs when the client area of the control is right clicked. e.ListViewRow contains clicked Item index. e.ListViewColumn contains clicked SubItem index. |
| ItemSelectionChanged | Occurs when the selected Item has changed. e.ListViewRow contains current selected Item index.|
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

### Items Collection (wfxListViewItemsCollection)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Add | Add a new item (and optional 32bit value) to the listview. ListView.Items.Add( Text, 32bitValue)|
| ByIndex | Return the wfxListViewItem object related to the specified listview item index.|
| Clear | Removes all items from the listview.|
| Count | Returns the total number of items in the listview.|
| Insert | Add a new item (and optional 32bit value) to the listview at a specific position. ListView.Items.Insert( Index, Text, 32bitValue)|
| Remove | Remove/delete the item identified by the index value.|
| SelectedCount | Returns the total number of selected items in the listview.|
| SortByColumn | Sort by the specified SubItem column index. ListView.Items.SortByColumn( nSortColumn, bSortAscend ) |

### Item (wfxListViewItem)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Index | The index value related to the listview item.|
| Checked | Gets or sets a value (true/false) indicating whether the Item is checked.|
| Selected | Gets or sets a value (true/false) indicating whether the Item is selected.|
| Data32 | The optional 32bit value associated with this listview item.|
| Text | The text associated with this listview item.|
| BackColor | Gets or sets the foreground color of the listview item. Refer to the **Colors** object. |
| ForeColor | Gets or sets the foreground color of the listview item. Refer to the **Colors** object. |
| SubItem | A specific wfxListViewSubItem of a listview item. |
| SubItems | A collection of subitems of a listview item. |

### SubItem (wfxListViewSubItem)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Index | The index value related to the listview subitem.|
| Text | The text associated with this listview item.|
| BackColor | Gets or sets the foreground color of the listview subitem. Refer to the **Colors** object. |
| ForeColor | Gets or sets the foreground color of the listview subitem. Refer to the **Colors** object. |

### SubItems Collection (wfxListViewSubItemsCollection)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Add | Add a new subitem to the listview. ListView.Items(i).SubItems.Add( Text )|
| ByIndex | Return the wfxListViewSubItem object related to the specified listview item index.|
| Clear | Removes all subitems from the listview item.|
| Count | Returns the total number of subitems for the listview item.|
| Remove | Removes a specific subitem from a listview item.|

### Column (wfxListViewColumn)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Index | The index value related to the listview column.|
| Text | The text associated with this listview column.|
| Width | The width of the listview column.|
| TextAlign | Set the text alignment of the column text.|

### Columns Collection (wfxListViewColumnsCollection)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Add | Add a new column to the listview. ListView.Columns.Add( Text, Width(pixels), TextAlign as TextAlignment )|
| ByIndex | Return the wfxListViewColumn object related to the specified column index.|
| Clear | Removes all columnss from the listview.|
| Count | Returns the total number of columns for the listview.|
| Remove | Removes a specific column from a listview.|


## Working with a ListView
```
' Row count = frmMain.ListView1.Items.Count
' Selected count = frmMain.ListView1.Items.SelectedCount
' Current row = frmMain.ListView1.SelectedIndex

' Add 3 columns each 100 pixels wide and each using a different text alignment
frmMain.ListView1.Columns.Add( "Column 1", 100, TextAlignment.Left )
frmMain.ListView1.Columns.Add( "Column 2", 100, TextAlignment.Center )
frmMain.ListView1.Columns.Add( "Column 3", 100, TextAlignment.Right )

' Add a bulk number of Items and SubItems (basically rows to the ListView)
' Prevent slow ListView updating by wrapping the operation in BeginUpdate/EndUpdate

frmMain.ListView1.BeginUpdate
dim i as long
for ii as long = 0 to 99
   i = frmMain.ListView1.Items.Add( "Line " & ii )
   frmMain.ListView1.Item(i).SubItems.Add( "L" & ii & "Sub1" )
   frmMain.ListView1.Item(i).SubItems.Add( "L" & ii & "Sub2" )
next
frmMain.ListView1.EndUpdate

' Change the foreground colors of a couple of SubItems
frmMain.ListView1.Item(0).SubItem(0).ForeColor = colors.Red
frmMain.ListView1.Item(0).SubItem(2).ForeColor = colors.Red

' Change the foreground/background colors of a couple of SubItems
frmMain.ListView1.Item(2).ForeColor = colors.Blue
frmMain.ListView1.Item(2).BackColor = colors.Yellow
frmMain.ListView1.Item(2).SubItem(1).ForeColor = colors.Blue

' Set the initial Item to be selected
frmMain.ListView1.SelectedIndex = 0
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
