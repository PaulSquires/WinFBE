# PopupMenu (wfxPopupMenu)

### Methods

| Name         | Description                                                  |
| ------------ | ------------------------------------------------------------ |
| AddItem      | Add an item to the menu. Additem( Text, ResultCode). The ResultCode is the value returned from the Show method if this menu option is selected. |
| AddSeparator | Add a separator bar to the menu.                             |
| Handle       | Retrieves the menu handle (HMENU). This handle can be used with regular Windows API functions. |
| Create       | Creates and initializes an empty popup menu. The menu is not displayed until the Show method is called. |
| Show         | Show the popup menu. Specify the parent window and the position (screen coordinates) to display the menu.( byval hParentWindow as HWND, byval xpos as long, byval ypos as long) |
| Show         | Show the popup menu. Specify the parent window and the position (screen coordinates) to display the menu.( byval hParentWindow as HWND, byval pt as POINT) |
| Show         | Show the popup menu. Specify the parent window and the position (screen coordinates) to display the menu.( byval hParentWindow as HWND, byval xpos as long, byval wpt as wfxPoint) |


##### Example: Right click popup menu on TreeNode
```
''
''
Function frmMain_TreeView1_MouseDown( ByRef sender As wfxTreeView, ByRef e As EventArgs ) As LRESULT
   
   if e.RButton then
      ' Determine if we have clicked on a TreeNode
      dim as wfxPoint wpt = e.pt
      wpt.Convert( frmMain.TreeView1.hWindow )
     
      ' Only continue if the TreeNode is valid
      dim byref TreeNode as wfxTreeNode = sender.GetNodeAt( wpt )
      if TreeNode.hNode = 0 then exit function
     
      ' Ensure that the node right clicked on is now selected
      TreeNode.Selected = true
     
      dim as wfxPopupMenu PopMenu
      PopMenu.AddItem( "Menu Option 1", 1 )
      PopMenu.AddItem( "Menu Option 2", 2 )
      PopMenu.AddItem( "Menu Option 3", 3 )
      PopMenu.AddSeparator
      PopMenu.AddItem( "Menu Option 4", 4 )
     
      ' Show the popup menu (need to use screen coordinates rather than client coordinates)
      dim as long nResult = PopMenu.Show( sender.hWindow, e.x, e.y )
     
      AfxMsg( "You selected menu option: " & nResult )
     
      select case nResult
         case 1
         case 2
         case 3
         case 4   
      end select
   end if
   
   Function = 0
End Function

```
