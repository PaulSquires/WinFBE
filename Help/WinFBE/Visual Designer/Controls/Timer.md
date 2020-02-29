# Timer (wfxTimer)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AutoReset | Gets or sets a value indicating whether the Timer should raise the Elapsed event only once (false) or repeatedly (true).|
| CtrlType | Gets or sets the control type value. Always **ControlType.Timer** and used when adding control to the application’s form collection.|
| Enabled | Gets or sets a value (true/false) indicating whether the Timer will start (true) or stop (false).|
| Interval | Gets or sets the interval, expressed in milliseconds, at which to raise the Elapsed event.|
| Parent | Gets or sets the parent container of the control.|
| Tag | Gets or sets user defined text associated with the form.|
| TimerId | Gets or sets a value indicating the Timer ID of the control.|

### Methods
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| StartTimer | Starts the Timer and begins the countdown based on the Interval property value.|
| StopTimer | Stops the Timer and begins the countdown based on the Interval property value.|

### Events
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllEvents | Special handler where all events are routed through. Use this handler if you prefer to use the Win32 api style messages and wParam and lParam parameters. Set the Handled element of EventArgs to true if you handle a message and do not want Windows to perform any further processing on the message.|
| Destroy | Occurs immediately before the control is about to be destroyed and all resources associated with it released.|
| Elapsed | Occurs when the Timer interval elapses.|


###
You do not have to kill or destroy a Timer. You simply need to StartTimer and StopTimer. The Timer resources are freed when StopTimer is called or the Form containing the Timer is closed.

