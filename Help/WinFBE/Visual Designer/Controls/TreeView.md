# TreeView (wfxTreeView)
#### (Note: Image support for TreeView has not yet been implemented)
### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllowDrop           | Gets or sets a value (true/false) indicating whether the control will accept data that is dragged onto it.        |
| Anchor | Specifies how a control anchors to the edges of its Form. |
| BackColor | Gets or sets the background color of the control. Refer to the **Colors** object. |
| BorderStyle | Gets or sets the border style of the control. Refer to the [ControlBorderStyle](#ControlBorderStyle) enum. |
| CheckBoxes | Gets or sets a value (true/false) indicating whether a check box appears next to each item in the control.|
| CtrlId | Gets or sets a value indicating the control ID of the control.|
| CtrlType | Gets or sets the control type value. Always **ControlType.TreeView** and used when adding control to the application’s form collection.|
| Enabled | Gets or sets a value (true/false) indicating whether the control can respond to user interaction.|
| FadeButtons | Fade expand/collapse buttons in or out when the mouse moves away or into a state of hovering over the control.|
| Focused | Gets or sets a value (true/false) indicating whether the control has input focus.|
| Font | Gets or sets the font for the control. Refer to the Font object.|
| ForeColor | Gets or sets the foreground color of the control. Refer to the **Colors** object. |
| FullRowSelect | Gets or sets a value indicating whether clicking an item selects all its subitems (for TreeViews this style cannot be used in conjunction with the ShowLines property.)|
| Height | Gets or sets the height of the control.|
| HideSelection | Indicates that the selection should be hidden when the control loses focus. |
| HotTracking | Determines whether items under the mouse pointer are automatically highlighted.|
| hWindow |  Gets the Windows handle (hwnd) of the control. |
| hWindowParent | Gets or sets the Windows handle (hwnd) of the parent control. |
| ItemHeight | Determines the height of an individual item in a control (in pixels). |
| Left | Gets or sets the distance, in pixels, between the left edge of the control and the left edge of its container's client area (normally the form).|
| Location |Gets or sets the top and left position of the control. Get: returns [wfxPoint](#wfxPoint) object. Set: (left, top). |
| Locked | Gets or sets a value (true/false) indicating whether the control can be moved or resized.|
| Node |Gets or sets properties of an TreeNode [wfxTreeNode](#wfxTreeNode) object |
| Nodes | Gets or sets properties of an TreeNode [wfxTreeNodeCollection](#wfxTreeNodeCollection) collection object |
| Parent | Gets or sets the parent container of the control.|
| Scrollable | Gets or sets a value (true/false) indicating whether the control displays scroll bars when they are needed.|
| ShowLines | Gets or sets a value (true/false) indicating whether lines are drawn between the tree nodes in the control.|
| ShowPlusMinus | Gets or sets a value (true/false) indicating whether the expand/collapse image is displayed next to tree nodes that contain child tree nodes.|
| ShowRootLines | Gets or sets a value (true/false) indicating whether lines are drawn between the tree nodes that are at the root of the control.|
| Size | Gets or sets the size of the form. Get: returns [wfxSize](#wfxSize) object. Set: (width, height)|
| Sorted | Gets or sets a value (true/false) indicating whether the nodes are automatically sorted alphabetically based on the value of the Text property.|
| TabIndex | Gets or sets the position that the control occupies in the TAB position.|
| TabStop | Gets or sets a value (true/false) indicating whether the user can use the TAB key to give focus to the control.|
| Tag | Gets or sets user defined text associated with the form.|
| Text | Gets or sets the text (caption) associated with this form.|
| Top | Gets or sets the distance, in pixels, between the top edge of the control and the top edge of its container's client area (normally the form).|
| TreeNode | The node that is currently being acted upon for messages such as BeforeExpand, AfterExpand, BeforeCollapse, AfterCollapse, BeforeCheck, AfterCheck, etc|
| Width | Gets or sets the width of the control.|

### Methods
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| BeginUpdate | Maintains performance while items are added to the control one at a time by preventing the control from drawing until the EndUpdate method is called.|
| EndUpdate | Resumes painting the control after painting had been suspended by the BeginUpdate method.|
| ExpandAll | Expands all the tree nodes. |
| CollapseAll | Collapses all the tree nodes. |
| GetNodeAt | Retrieves the tree node that is at the specified location. GetNodeAt(Point) or GetNodeAt(xpos, ypos)|
| Hide | Conceals the control from the user.|
| Refresh | Forces the form to invalidate its client area and immediately redraw itself and any child controls.|
| SelectNextControl | Moves the input control to the next (True) or previous (False) control in the tab order. |
| SetBounds | Sets the bounds of the control to the specified location and size. SetBounds(left, top, width, height).|
| Show | Creates and displays the control to the user.|
| Sort | Sorts the items in TreeView control.|

### Events
| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllEvents | Special handler where all events are routed through. Use this handler if you prefer to use the Win32 api style messages and wParam and lParam parameters. Set the Handled element of EventArgs to true if you handle a message and do not want Windows to perform any further processing on the message.|
| Click | Occurs when a node in the TreeView is clicked on or when the current selected node changes.|
| Destroy | Occurs immediately before the control is about to be destroyed and all resources associated with it released.|
| DropFiles | Occurs when an object is dragged and dropped onto the control and the AllowDrop property of the control is set to True.|
| BeforeExpand | Occurs before the tree node is expanded.|
| AfterExpand | Occurs after the tree node is expanded.|
| BeforeCollapse | Occurs before the tree node is collapsed.|
| AfterCollapse | Occurs after the tree node is collapsed.|
| BeforeSelect | Occurs before the tree node is selected.|
| AfterSelect | Occurs after the tree node is selected.|
| BeforeCheck | Occurs before the tree node check box is checked.|
| AfterCheck | Occurs after the tree node check box is checked.|
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


# TreeNode (wfxTreeNode)
### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| hNode | The node handle (HTREEITEM) used by Win32 api functions.|
| Index | The position of the node within the node collection.|
| Checked | Gets or sets a value (true/false) indicating that the node's checkmark is checked.|
| Selected | Gets or sets a value (true/false) indicating that the node is the currently selected node in the control.|
| Text | Gets or sets the text that displays for the node.|
| Data32 | Gets or sets a value (32-bit) associated with the node. This value can be used by the user for whatever purpose necessary.|
| ToolTipText | Gets or sets ToolTip text that will display when the mouse pointer moves over this TreeNode. |
| pTreeView | A pointer to the main wfxTreeView class.|
| Node |Gets or sets properties of an TreeNode [wfxTreeNode](#wfxTreeNode) object |
| Nodes | Gets or sets properties of an TreeNode [wfxTreeSubNodeCollection](#wfxTreeSubNodeCollection) collection object 
### Methods

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Remove | Removes the tree node from the control.|
| Expand | Expands the tree node.|
| ExpandAll | Expands all the child tree nodes.|
| Collapse | Collapses the TreeNode.|
| EnsureVisible | Ensures that the tree node is visible, expanding tree nodes and scrolling the tree view control as necessary.|
| SortChildren | Sorts the immediate child nodes of the TreeNode.|
| Nodes | Collection of child nodes of this TreeNode (wfxTreeSubNodeCollection.|

# TreeNodeCollection (wfxTreeNodeCollection)
# TreeSubNodeCollection (wfxTreeSubNodeCollection)
### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| Add | Add a new node (and optional 32bit value) to the treeview. TreeView.Nodes.Add( Text, 32bitValue) or TreeNode.Nodes.Add( Text, 32bitValue)|
| ByIndex | Return the wfxTreeNode object related to the specified TreeNode item index.|
| Clear | Removes all items from the collection.|
| Count | Returns the total number of nodes in the collection.|
| Insert | Add a new node (and optional 32bit value) to the treeview at a specific position. TreeView.Nodes.Insert( Index, Text, 32bitValue)|
| Remove | Remove/delete the node identified by the index value.|
| pTreeView | A pointer to the main wfxTreeView class.|


# How to use the TreeView

## Adding root and children nodes
The following example will create four root nodes. The fourth root node (Project 4), will have several children nodes. If the TreeView has the Sorted property set to True then the nodes and children nodes will automatically be sorted when they are added to the control.
```
dim as long root, node, subnode, subsubnode
root = frmMain.TreeView1.Nodes.Add("Project 1")
root = frmMain.TreeView1.Nodes.Add("Project 2")
root = frmMain.TreeView1.Nodes.Add("Project 13")
    
node = frmMain.TreeView1.Node(root).Nodes.Add("Project 6")
for i as long = 1 to 3
   subnode = frmMain.TreeView1.Node(root).Node(node).Nodes.Add("Project File" & i)
   for i as long = 1 to 2
      subsubnode = frmMain.TreeView1.Node(root).Node(node).Node(subnode).Nodes.Add("Sub Project File" & i)
      for i as long = 1 to 5
         frmMain.TreeView1.Node(root).Node(node).Node(subnode).Node(subsubnode).Nodes.Add("Sub Sub Project File" & i)
      next
   next
next
```
The above example just adds a text description to each new TreeNode. You can also specifiy an optional 32-bit Long integer that you can use for whatever purpose you wish.
```
frmMain.TreeView1.GetSelectedNode.Nodes.Add( "New Node", 10232 )
```

## Deleting/Removing a TreeNode
```
' The following removes the first node off of the root using the collection:
frmMain.TreeView1.Nodes.Remove(0)

' Remove the node using the TreeNode Remove method:
frmMain.TreeView1.Node(0).Remove

' -or- (notice the use of ByRef)
Dim ByRef TreeNode = frmMain.TreeView1.Node(0)
TreeNode.Remove

' Every node also has child node collections, so to remove the first child node of the first root node then do this:
frmMain.TreeView1.Node(0).Nodes.Remove(0)

' Remove the currently selected node
frmMain.TreeView1.SelectedNode.Remove
```

## Display a Popup Menu when Right Click on a TreeNode
```
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

## TreeView Sorting
You can specify that the all nodes in the TreeView are automatically sorted whenever a node is added or inserted into the tree.
```
frmMain.TreeView1.Sorted = true
```
If the tree is not sorted, you can manually make a call to the sort method to sort all of the nodes.
```
frmMain.TreeView1.Sort
```
If the tree's sorting is False, you can manually sort individual nodes (and their children nodes).
```
frmMain.TreeView1.GetSelectedNode.SortChildren
' -or-
frmMain.TreeView1.Node(nIndex).SortChildren
```

## Actions on a TreeNode
```
' Remember, you can retrieve a reference to the TreeNode rather than always having to access the node through the collection
Dim ByRef TreeNode = frmMain.TreeView1.Node(nIndex)

' Set the text for a node
frmMain.TreeView1.Node(nIndex).Text = "New Node Text"

' Set the ToolTip text that displays when the mouse passes over the node
frmMain.TreeView1.Node(nIndex).TooTipText = "This is the tooltip"

' Make the specified node the Selected Node
frmMain.TreeView1.Node(nIndex).Selected = true

' Check the checkbox for the node (if CheckBox style is enabled for the TreeVie)
frmMain.TreeView1.Node(nIndex).Checked = true

' Expand the TreeNode
frmMain.TreeView1.Node(nIndex).Expand

' Expand the TreeNode and all of it's children nodes
frmMain.TreeView1.Node(nIndex).ExpandAll

' Ensure that the TreeNode is visible
frmMain.TreeView1.Node(nIndex).EnsureVisible

' Retrieve the handle to the node (needed if you want to manipulate the node using Windows API functions)
Dim As HTREEITEM hNode = frmMain.TreeView1.Node(nIndex).hNode 
```

## Searching entire TreeView
Because a TreeView is nothing more than a collection of nodes and child nodes, you can use a recursive approach to search the entire TreeView.
```
function SearchTreeView( byref ParentNode as wfxTreeNode, byval wszSearchText as CWSTR ) as boolean
   ' Do we have a match?
   if ParentNode.Text = wszSearchText then 
      ' Text was matched. Do whatever we need to do to the node.
      ' For example, set the node to be the selected node.
      ParentNode.Selected = true
      ParentNode.EnsureVisible 
      return true   ' cancel any further searching
   end if
   
   ' Recursively search any subnodes of this TreeNode
   for i as long = 0 to ParentNode.Nodes.count - 1
      SearchTreeView( ParentNode.Node(i), wszSearchText )
   next      
   function = 0
end function

' Start our search by iterating the root nodes
dim as CWSTR wszSearchText = "Sub Sub Project File1"
for i as long = 0 to frmMain.TreeView1.Nodes.count - 1
   if SearchTreeView( frmMain.TreeView1.Node(i), wszSearchText ) then exit for
next
```

## Reducing screen flicker
You can reduce screen flicker when adding bulk amount of nodes by wrapping the code in BeginUpdate and EndUpdate.
```
frmMain.TreeView1.BeginUpdate
   
dim as long root
For i as long = 1 to 10000
   root = frmMain.TreeView1.Nodes.Add("Node #" & i)
Next
   
frmMain.TreeView1.EndUpdate
```

## Canceling an action on a TreeView
Some actions on a TreeView are allowed to be canceled. For example, when navigating from node to node, events are fired. In your code, you have the ability to stop the movement from the current node to the new node. To cancel an action you simply need to set e.Cancel = True prior to returning from the event handler (alternatively, you can set the event handler itself to return False). You can only cancel an action in a "Before" handler.

Examples:
```
Function frmMain_TreeView1_BeforeSelect( ByRef sender As wfxTreeView, ByRef e As EventArgs ) As LRESULT
   ? "BeforeSelect: "; sender.TreeNode.Text
   e.Cancel = true
   Function = 0
End Function

Function frmMain_TreeView1_BeforeExpand( ByRef sender As wfxTreeView, ByRef e As EventArgs ) As LRESULT
   ? "BeforeExpand: "; sender.TreeNode.Text
   ' e.Cancel = true
   Function = true  ' <-- same effect as e.Cancel = True
End Function

Function frmMain_TreeView1_BeforeCollapse( ByRef sender As wfxTreeView, ByRef e As EventArgs ) As LRESULT
   ? "BeforeCollapse: "; sender.TreeNode.Text
   e.Cancel = true
   Function = 0
End Function
```
In the above examples, notice the use of the property "TreeNode" of the sender object. This is a special property that represents the TreeNode that is currently being acted upon. This can be different than the SelectedNode TreeNode (especially if a node is being expanded or collapsed programmatically).





