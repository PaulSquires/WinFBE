# Date Time Picker(wfxDateTimePicker)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllowDrop           | Gets or sets a value (true/false) indicating whether the control will accept data that is dragged onto it.        |
| CalendarRightAlign | Indicates whether the drop-down calendar will be right aligned with the control instead of left aligned which is the default.|
| DateFormat | Gets or sets the format of the date and time displayed in the control. Refer to the DateTimePickerFormat enum.|
| CtrlId | Gets or sets a value indicating the control ID of the control.|
| CtrlType | Gets or sets the control type value. Always **ControlType.DateTimePicker** and used when adding form to the application's form collection.|
| Enabled | Gets or sets a value (true/false) indicating whether the control can respond to user interaction.|
| Font | Gets or sets the font for the control. Refer to the Font object.|
| FormatCustom | Gets or sets the custom format when the DateFormat property is set to CustomFormat. |
| Height | Gets or sets the height of the control.|
| hWindow |  Gets the Windows handle (hwnd) of the control. |
| hWindowParent | Gets or sets the Windows handle (hwnd) of the parent control. |
| Left | Gets or sets the distance, in pixels, between the left edge of the control and the left edge of its container's client area (normally the form).|
| Location |Gets or sets the top and left position of the control. Get: returns [wfxPoint](#wfxPoint) object. Set: (left, top). |
| Locked | Gets or sets a value (true/false) indicating whether the control can be moved or resized.|
| SelectedDate | Gets or sets the selected date in the format YYYMMDD. |
| SelectedTime | Gets or sets the selected time in the format HHMMSS. |
| Parent | Gets or sets the parent container of the control.|
| Size | Gets or sets the size of the form. Get: returns [wfxSize](#wfxSize) object. Set: (width, height)|
| TabIndex | Gets or sets the position that the control occupies in the TAB position.|
| TabStop | Gets or sets a value (true/false) indicating whether the user can use the TAB key to give focus to the control.|
| Tag | Gets or sets user defined text associated with the form.|
| ShowUpDown | Gets or set a value indicating whether an up-down control is used to adjust the date/time value. |
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
| DateTimeChanged | Indicates that a change (date or time has changed its value) has occured in the control. |
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


##### DateTimePickerFormat 
```
Enum DateTimePickerFormat 
   LongDate = 1
   ShortDate
   ShortDateCentury
   TimeFormat
   CustomFormat
End Enum
```

If The DateFormat property is set to DateTimePickerFormat = CustomFormat, then the control will use the value of the FormatCustom property to display the date and/or time in the control.

It is acceptable to include extra characters within the format string to produce a more rich display. However, any nonformat characters must be enclosed within single quotes. For example, the format string "'Today is: 'hh':'m':'s ddddMMMdd', 'yyy" would produce output like "Today is: 04:22:31 Tuesday Mar 23, 1996". Note: When entering such a custom format string from within the WinFBE Visual Designer, you do not enclose the string in double quotes. WinFBE will automatically enclose the string in double quotes when code is generated at compile time.

A custom format string consists of a series of elements that represent a particular piece of information and define its display format. The elements will be displayed in the order they appear in the format string. Date and time format elements will be replaced by the actual date and time. They are defined by the following groups of characters.

| Element |Description |
| ------- | ---------- |
| "d" | The one- or two-digit day. |
| "dd" | The two-digit day. Single-digit day values are preceded by a zero.|
| "ddd" | The three-character weekday abbreviation. |
| "dddd" | The full weekday name. |
| "h" | The one- or two-digit hour in 12-hour format. |
| "hh" | The two-digit hour in 12-hour format. Single-digit values are preceded by a zero.|
| "H" | The one- or two-digit hour in 24-hour format.|
| "HH"| The two-digit hour in 24-hour format. Single-digit values are preceded by a zero.|
| "m" | The one- or two-digit minute.|
| "mm" | The two-digit minute. Single-digit values are preceded by a zero.|
| "M" | The one- or two-digit month number.|
| "MM" | The two-digit month number. Single-digit values are preceded by a zero.|
| "MMM" |The three-character month abbreviation.|
| "MMMM" |The full month name.|
| "t" | The one-letter AM/PM abbreviation (that is, AM is displayed as "A").|
| "tt" | The two-letter AM/PM abbreviation (that is, AM is displayed as "AM").|
| "yy" | The last two digits of the year (that is, 1996 would be displayed as "96"). |
| "yyyy" | The full year (that is, 1996 would be displayed as "1996").|

 


