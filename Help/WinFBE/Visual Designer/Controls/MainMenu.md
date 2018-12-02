# MainMenu (wfxMainMenu)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| CtrlType | Gets or sets the control type value. Always **ControlType.MainMenu** and used when adding control to the application’s form collection.|
| Handle | Gets or sets a value indicating the HMENU handle of the control.|
| IsLoading | Gets or sets a value (true/false) indicating whether the control is currently being created.|
| MenuItem | Gets or sets values for a specific MenuItem belonging to the MainMenu (refer to [wfxMenuItem](#wfxMenuItem)).|
| MenuItems | Gets or sets values for the collection of MenuItems that belong to the MainMenu (refer to [wfxMenuItemsCollection](#wfxMenuItemsCollection)).|

### Methods
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| ByMenuID | Get a reference to a MenuItem [wfxMenuItem](#wfxMenuItem) belonging to the MainMenu based on a MenuItem menu ID.|
| ByMenuName | Get a reference to a MenuItem [wfxMenuItem](#wfxMenuItem) belonging to the MainMenu based on a MenuItem name.|
| ByPopupMenuHandle | Get a reference to a MenuItem [wfxMenuItem](#wfxMenuItem) belonging to the MainMenu based on a MenuItem popup menu handle.|
| Refresh | Forces the form to invalidate its client area and immediately redraw itself and any child controls.|
| Show | Creates and displays the control to the user.|

### Events
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllEvents | Special handler where all events are routed through. Use this handler if you prefer to use the Win32 api style messages and wParam and lParam parameters. Set the Handled element of EventArgs to true if you handle a message and do not want Windows to perform any further processing on the message.|
| Click | Occurs when a menu item is clicked.|
| Popup | Occurs when a drop-down menu or submenu is about to become active. This allows an application to modify the menu before it is displayed without changing the entire menu.|

### MenuItems Collection (wfxMenuItemsCollection)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Add | Add a new MenuItem to the MainMenu.|
| ByIndex | Return the wfxMenuItem object related to the specified MainMenu menu item index.|
| Clear | Removes all items from the MainMenu.|
| Count | Returns the total number of items in the MainMenu.|
| Remove | Remove/delete the item identified by the index value.|

### Item (wfxMenuItem)
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Checked | Gets or sets a value (true/false) indicating whether to display a check mark next to the MenuItem.|
| Grayed | Gets or sets a value (true/false) indicating whether to display the MenuItem as enabled or disabled.|
| MenuItem | Gets or sets values for a specific MenuItem belonging to the MainMenu (refer to [wfxMenuItem](#wfxMenuItem)).|
| MenuItems | Gets or sets values for the collection of MenuItems that belong to the MainMenu (refer to [wfxMenuItemsCollection](#wfxMenuItemsCollection)).|
| Handle | Gets or sets a value indicating the HMENU handle of the control.|
| Index | Gets or sets the index value related to the MenuItem.|
| IsParent | Gets or sets a value (true/false) indicating that this MenuItem has child popup MenuItems.|
| MenuID | Gets or sets the menu ID related to the MenuItem.|
| Name | Gets or sets the name related to the MenuItem.|
| Shortcut | Gets or sets the shortcut accelerator text associated with this MenuItem.|
| Text | Gets or sets the text associated with this MenuItem.|
| PopupMenuHandle | Gets or sets the handle to the child popup menu of this MenuItem |
| MenuParent | Gets or sets a _wfxMainMenu ptr_ to the MainMenu object that this MenuItem belongs to.|


## Working with a MainMenu

### Creating the MainMenu

```
   this.MainMenu.Parent = @this

   dim mnuFile   as wfxMenuItem = wfxMenuItem("File", "mnuFile")
   dim mnuNew    as wfxMenuItem = wfxMenuItem("New",  "mnuNew", "Ctrl+DEL")
   DIM mnuOpen   as wfxMenuItem = wfxMenuItem("Open", "mnuOpen")
   DIM mnuSave   as wfxMenuItem = wfxMenuItem("Save", "mnuSave", "Ctrl+S")
   DIM mnuSaveAs as wfxMenuItem = wfxMenuItem("Save As", "mnuSaveAs", "F12")
   DIM mnuSep1   as wfxMenuItem = wfxMenuItem("-",    "mnuSep1")
   DIM mnuExit   as wfxMenuItem = wfxMenuItem("Exit", "mnuExit")
   mnuFile.MenuItems.Add(mnuNew)
   mnuFile.MenuItems.Add(mnuOpen)
   mnuFile.MenuItems.Add(mnuSave)
   mnuFile.MenuItems.Add(mnuSaveAs)
   mnuFile.MenuItems.Add(mnuSep1)
   mnuFile.MenuItems.Add(mnuExit)
   this.MainMenu.MenuItems.Add(mnuFile)
   
   DIM mnuEdit  as wfxMenuItem = wfxMenuItem("Edit", "mnuEdit")
   DIM mnuCut   as wfxMenuItem = wfxMenuItem("Cut",  "mnuCut")
   DIM mnuCopy  as wfxMenuItem = wfxMenuItem("Copy", "mnuCopy")
   DIM mnuPaste as wfxMenuItem = wfxMenuItem("Paste", "mnuPaste")
   mnuEdit.MenuItems.Add(mnuCut)
   mnuEdit.MenuItems.Add(mnuCopy)
   mnuEdit.MenuItems.Add(mnuPaste)
   this.MainMenu.MenuItems.Add(mnuEdit)
   
   this.MainMenu.OnPopup = @Form1_MainMenu_Popup
   this.MainMenu.OnClick = @Form1_MainMenu_Click

   this.Controls.Add(ControlType.MainMenu, @this.MainMenu)
```
As you can see from the example above, you must create each individual MenuItem and then assign them to the parent MenuItem that is the parent for the specific MenuItem. Every MenuItem can have children MenuItems - these are the submenus that popup when a MenuItem is clicked.

In the example, the _this_ identifier was used because the example was created in the constructor fo the Form that contained the MainMenu. If created elsewhere, you would use the specifically declared variable (eg. Form1, Form2, frmMain, etc).

Also, notice that the MenuItem constructor can be called using one or more values. 
```
declare constructor( ByRef wszText     As WString = "", _
                     ByRef wszName     As WString = "", _ 
                     Byref wszShortcut As WString = "", _
                     ByVal bChecked    As boolean = false, _
                     ByVal bGrayed     As boolean = false )
```
### Shortcuts
A shortcut is a keyboard accelerator key combination that will activate a specific menu entry. It can be any combination of the Ctrl, Alt, Shift keys and another key. The text of the shortcut key is displayed to the right of the menu text.

### Retrieving Reference to a specific MenuItem
You can search for a specific MenuItem using several built in methods of the MainMenu class. The key is that you should declare your variable as *ByRef* so that the returned reference is assigned to it and points to the original MenuItem (much like a *pointer* does).
```
Dim ByRef iMenuItem = Form1.MainMenu.ByMenuName("mnuSave")
Dim ByRef iMenuItem = Form1.MainMenu.ByMenuID(110)
Dim ByRef iMenuItem = Form1.MainMenu.ByMenuID( GetMenuItemID(Form1.MainMenu.Handle, 5) )
Dim ByRef iMenuItem = Form1.MainMenu.ByPopupMenuHandle( GetSubMenu(Form1.MainMenu.Handle, 5) )

' Retrieve info on all top level MenuItems that show on the menu bar
Dim ByRef iMenuItem 
For i as long = 0 to Form1.MainMenu.MenuItems.Count – 1
   iMenuItem = Form1.MainMenu.Menutem(i)
   ? iMenuItem.Text, iMenuItem.Name
Next

' Rename all of the items in the File menu
Dim byref mnuFile as wfxMenuItem = Form1.MainMenu.ByMenuName("mnuFile")
for i as long = 0 to mnuFile.MenuItems.Count - 1
   mnuFile.MenuItem(i).Text = "Paul"
next
```

### Iterate all MenuItems in the entire MainMenu
This is not as straightforward as you might expect because every MenuItem can itself have many children MenuItems. There is no one single collection that contains all of the MenuItem objects for the entire MainMenu. You need to look at each MenuItem and process that object individually. The best way to perform this type of operation is through the use of recursion.

The following example iterates over all MenuItems and disables every MenuItems that is not a root item that displays directly on the MainMenu's menu bar.  
```
function IterateMenuItems( byref tMenuItem as wfxMenuItem ) as long
   
   ' Internal function that recursive processes all of the menu items. 
   ' It takes into account that submenus can exist for each MenuItem.

   select case ucase(tMenuItem.Name)
      CASE "MNUFILE", "MNUEDIT", "MNUVIEW", "MNUHELP"
         ' Do not modify these. These are on the MainMenu menu bar.
      CASE ELSE
         tMenuItem.Grayed = True
   end select

   ' Process any children MenuItems recursively
   For i As Long = 0 To tMenuItem.MenuItems.Count - 1
      if tMenuItem.MenuItem(i).IsParent then
         IterateMenuItems( tMenuItem.MenuItem(i) )
      end if   
   Next
   
   function = 0
end function


function Form1_Button1_Click( byref sender as wfxButton, byref e as EventArgs ) as LRESULT
   ' Iterate all MenuItems in the MainMenu and set each one to Grayed.
   ' We start at the root MenuItem that is dsplayed on the MainMenu's menu bar
   ' (eg. File, Edit, View, Help) and use recursion to visit every sub-MenuItem. 
   For i As Long = 0 To Form1.MainMenu.MenuItems.Count - 1
      IterateMenuItems( Form1.MainMenu.MenuItem(i) )
   Next

   function = 0
end function
```

### Responding to a Menu Click
When a user clicks on a menu item in the main menu then a *Click* event is fired that the user can respond to. 
```
function Form1_MainMenu_Click( byref sender as wfxMenuItem, byref e as EventArgs ) as LRESULT

   select case ucase(sender.Name)
      CASE "MNUFILE"
      CASE "MNUEDIT"
      CASE "MNUVIEW"
      CASE "MNUHELP"
   end select
   
   function = 0
end function
```

### Initializing Menu Items before they are shown
Immediately prior to a popup menu being shown on the screen, a *Popup* event is fired that the user can respond to. This allows the user and opportunity to initialize MenuItems based on any state information for the application. For example, in a text editor application you may gray out items such as Cut, Copy, Paste if such operations are not valid at the time the menu is displayed.
```
function Form1_MainMenu_Popup( byref sender as wfxMenuItem, byref e as EventArgs ) as LRESULT

   select case ucase(sender.Name)
      CASE "MNUFILE"
      CASE "MNUEDIT"
         Form1.MainMenu.ByMenuName("mnuCut").Grayed = bCanCut
         Form1.MainMenu.ByMenuName("mnuCopy").Grayed = bCanCopy
         Form1.MainMenu.ByMenuName("mnuPaste").Grayed = bCanPaste
      CASE "MNUVIEW"
      CASE "MNUHELP"
   end select
   
   function = 0
end function
```
