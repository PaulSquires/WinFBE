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
   
   if ubound(_arrControls)-1 THEN
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
   for i as long = lbound(_arrControls) to ubound(_arrControls)
? _arrControls(i)->ControlType, _arrControls(i)->IsSelected, _arrControls(i)->IsActive, "";
? _arrControls(i)->rcHandles(GRAB_TOP).Left, _arrControls(i)->rcHandles(GRAB_TOP).top, _arrControls(i)->rcHandles(GRAB_TOP).right, _arrControls(i)->rcHandles(GRAB_TOP).bottom
   NEXT
? ""
   function = 0
end function
