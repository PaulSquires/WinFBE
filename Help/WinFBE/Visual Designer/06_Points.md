# Points (wfxPoint)

As a programmer, you have to manage a POINT many times throughout the life of your programs. A POINT is simply an x, y coordinate that many Windows API's require. Sometimes the POINT needs to be based on screen coordinates while at other times it needs to be relative to the client area of a particular control depending on what function is being called. The wfxPoint class makes handling points and converting from one coordinate to another easy.

#### Creating and assigning a new point

```
Dim wpt As wfxPoint = pt
Dim wpt As wfxPoint = wfxPoint(x,y)

Dim wpt As wfxPoint
wpt = pt
wpt.x = x
wpt.y = y

' Determine if the point is empty
If wpt.IsEmpty Then
   ' Point has no values
End If   
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

