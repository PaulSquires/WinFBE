# TreeView (wfxTreeView)

### Properties

| Name                            | Description                    |
| ------------------------------- | ------------------------------ |
| AllowDrop           | Gets or sets a value (true/false) indicating whether the control will accept data that is dragged onto it.        |
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
| pTreeView | A pointer to the main wfxTreeView class.|

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


##### ControlBorderStyle
```
Enum ControlBorderStyle
   None  = 1
   FixedSingle  
   Fixed3D 
End Enum
```

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

## 



