# Size (wfxSize)

This class allows you to easily retrieve the height and width of a control. You can not use this class to set the height and width of a control.

#### Creating and assigning a new size

```
Dim sz As wfxSize
sz.Width = 100
sz.Height = 400

Dim sz As wfxSize = frmMain.TreeView1.hWindow
' sz.Height and sz.Width will now reflect the height and width of the control.

Dim sz As wfxSize
sz.Convert( frmMain.TreeView1.hWindow )
' sz.Height and sz.Width will now reflect the height and width of the control.

```

#### Convert between Coordinates  

```
Function frmMain_TreeView1_MouseDown( ByRef sender As wfxTreeView, ByRef e As EventArgs ) As LRESULT
   
   ' The incoming e.pt is screen coordinates based. Use the Convert method
   ' of the class to change the point relative to the TreeView window.
   
   dim as wfxPoint wpt = e.pt
   wpt.Convert( frmMain.TreeView1.hWindow )

   Function = 0
End Function   
```

