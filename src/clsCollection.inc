'    WinFBE - Programmer's Code Editor for the FreeBASIC Compiler
'    Copyright (C) 2016-2018 Paul Squires, PlanetSquires Software
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT any WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS for A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

''
''  Control Collection
''

constructor clsCollection
   '_arrControls
end constructor

destructor clsCollection
   '_arrControls
   for i as long = lbound(_arrControls) to ubound(_arrControls)
      delete _arrControls(i)
   NEXT
end destructor

property clsCollection.Count() as long
   return ubound(_arrControls) - lbound(_arrControls) + 1
end property

property clsCollection.ItemFirst() as long
   return lbound(_arrControls)
end property

property clsCollection.ItemLast() as long
   return Ubound(_arrControls)
end property

function clsCollection.ItemAt( byval nIndex as long ) as clsControl ptr
   return _arrControls(nIndex)
end function

'function clsCollection.ItemByName( byref wszName as wstring ) as clsControl ptr
'   for i as long = lbound(_arrControls) to ubound(_arrControls)
'       if ucase(_arrControls(i)->Name) = ucase(wszName) THEN
'         return _arrControls(i) 
'       END IF
'   next
'   function = 0   ' null pointer not found
'end function

function clsCollection.Add( byval pCtrl as clsControl ptr ) as long
   dim as long ub = ubound(_arrControls) + 1
   redim preserve _arrControls(ub) as clsControl ptr
   _arrControls(ub) = pCtrl
   function = ub
end function

function clsCollection.Remove( byval pCtrl as clsControl ptr ) as long
   dim as long idx = -1
   for i as long = lbound(_arrControls) to ubound(_arrControls)
      if _arrControls(i) = pCtrl THEN
         idx = i: exit for
      END IF
   NEXT
   if idx = -1 THEN exit function
      
   for i as long = idx to ubound(_arrControls) - 1
      _arrControls(i) = _arrControls(i+1)
   next
   
   ' Destroy control and unallocate the memory.
   if pCtrl->AfxButtonPtr then delete pCtrl->AfxButtonPtr
   DestroyWindow(pCtrl->hWindow)
   delete pCtrl
   
   if ubound(_arrControls) = -1 THEN
      erase _arrControls
   else
      redim preserve _arrControls(ubound(_arrControls)-1) as clsControl ptr
   END IF
   
   function = idx
end function

function clsCollection.DeselectAllControls() as long
   ' Deselect the selected state of all controls
   for i as long = lbound(_arrControls) to ubound(_arrControls)
      _arrControls(i)->IsSelected = false
   NEXT
   function = 0
end function

function clsCollection.SelectedControlsCount() as long
   ' Return number of controls with selected status (except for form itself)
   dim as Long nCount = 0
   for i as long = lbound(_arrControls) to ubound(_arrControls)
      if _arrControls(i)->ControlType <> CTRL_FORM THEN
         if _arrControls(i)->IsSelected then nCount = nCount + 1
      END IF
   NEXT
   function = nCount
end function

function clsCollection.SelectAllControls() as long
   ' Select the selected state of all controls (except for the form itself)
   dim as long idxActive = -1
   for i as long = lbound(_arrControls) to ubound(_arrControls)
      if _arrControls(i)->ControlType = CTRL_FORM then
         _arrControls(i)->IsSelected = false
      else
         if _arrControls(i)->IsActive THEN idxActive = i
         _arrControls(i)->IsSelected = true
      end if   
   NEXT
   ' If there was no active control already set before selecting all
   ' of the controls then we need to set one now.
   if idxActive = -1 THEN this.SetActiveControl(0)
   function = 0
end function

function clsCollection.SelectControl( byval hWndCtrl as hwnd )as long
   for i as long = lbound(_arrControls) to ubound(_arrControls)
      if _arrControls(i)->hWindow = hWndCtrl then
         _arrControls(i)->IsSelected = true
         exit for    
      END IF
   NEXT
   function = 0
end function

function clsCollection.SetActiveControl( byval hWndCtrl as hwnd) as long
   dim as long idxActive = -1
   for i as long = lbound(_arrControls) to ubound(_arrControls)
      _arrControls(i)->IsActive = iif(_arrControls(i)->hWindow = hWndCtrl, true, false)
      if _arrControls(i)->IsActive THEN idxActive = i
   NEXT
   ' Ensure that at least one control is active. Only controls that are
   ' selected can also be an active control.
   if idxActive = -1 THEN
      idxActive = 0   ' default that the form will be selected
      for i as long = lbound(_arrControls) to ubound(_arrControls)
         if _arrControls(i)->IsSelected then idxActive = i
      NEXT
   END IF
   _arrControls(idxActive)->IsActive = true
   _arrControls(idxActive)->IsSelected = true
   function = 0
end function

function clsCollection.GetActiveControl() as clsControl ptr
   for i as long = lbound(_arrControls) to ubound(_arrControls)
      if _arrControls(i)->IsActive then
         return _arrControls(i)
      end if   
   NEXT
   function = 0
end function

function clsCollection.GetCtrlPtr( byval hWndCtrl as hwnd) as clsControl ptr
   for i as long = lbound(_arrControls) to ubound(_arrControls)
      if _arrControls(i)->hWindow = hWndCtrl then
         return _arrControls(i)
      END IF
   NEXT
   function = 0
end function

function clsCollection.Debug() as long
   ? "Control Count = "; this.Count
   for i as long = lbound(_arrControls) to ubound(_arrControls)
      ? GetControlName(_arrControls(i)->ControlType), _arrControls(i)
   NEXT
   ? ""
   function = 0
end function
