''
''  Control Collection
''

constructor clsCollection
   '_arrControls
end constructor

destructor clsCollection
   '_arrControls
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


