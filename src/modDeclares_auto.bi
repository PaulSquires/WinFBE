Declare Function CreateTempResourceFile() as boolean
Declare Function ResetScintillaCursors() as Long
Declare Function code_Compile( ByVal wID As Long ) As BOOLEAN
Declare Function CreateChildProcess( byref zCmdLine AS wstring, byref zParams AS wstring, byval hChildStdOutWrite AS HANDLE ) AS BOOLEAN
Declare Sub RedirConsoleToFile( byref szExe AS wstring, byref szCmdLine AS wstring, byref sConsoleText AS string )
Declare Function RunEXE( ByRef wszFileExe As CWSTR, ByRef wszParam As CWSTR ) As Long
Declare Function SetDocumentErrorPosition( ByVal hLV As HWnd, Byval wID as long ) As Long
Declare Function ParseLogForError( byref wsLogSt as CWSTR, byval bAllowSuccessMessage as Boolean, Byval wID as long, byval fCompile64 as Boolean ) as Boolean
Declare Function PositionExplorerWindows( ByVal HWnd As HWnd ) As LRESULT
Declare Function AddFunctionsToExplorerTreeview( ByVal pDoc As clsDocument Ptr, ByVal fUpdateNodes As BOOLEAN ) As Long
Declare Function GetFormName( byval pDoc as clsDocument ptr ) as CWSTR
Declare Function SetColorListBoxSelection( byref wszPropValue as wstring ) as Long
Declare Function KeyboardMoveControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long
Declare Function KeyboardResizeControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long
Declare Function KeyboardCycleActiveControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long
Declare Function IsControlNameExists( byval pDoc as clsDocument ptr, byref wszControlName as wstring ) as boolean
Declare Function IsControlLocked( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr ) as boolean
Declare Function IsFormNameExists( byref wszFormName as wstring ) as boolean
Declare Function GenerateDefaultFormName() as CWSTR
Declare Function GenerateDefaultControlName( byval pDoc as clsDocument ptr, byval nControlType as long ) as CWSTR
Declare Function GenerateDefaultTabIndex( byval pDoc as clsDocument ptr ) as CWSTR
Declare Function CreateToolboxControl( byval pDoc as clsDocument ptr, byval ControlType as long, byref rcDraw as RECT ) as clsControl ptr
Declare Function GetImagesTypePtr( byref wszImageName as wstring ) As IMAGES_TYPE ptr
Declare Function GetActiveToolboxControlName() as CWSTR
Declare Function GetActiveToolboxControlClassName() as CWSTR
Declare Function GetActiveToolboxControlType() as Long
Declare Function GetActivePropertyPtr() as clsProperty ptr
Declare Function GetActiveEventPtr() as clsEvent ptr
Declare Function GetFormCtrlPtr( byval pDoc as clsDocument ptr ) as clsControl ptr
Declare Function SetActiveToolboxControl( byval ControlType as long ) as Long
Declare Function GetControlClassName( byval pCtrl as clsControl ptr ) as CWSTR
Declare Function GetToolBoxName( byval nControlType as long ) as CWSTR
Declare Function GetControlName( byval nControlType as long ) as CWSTR
Declare Function GetControlType( byval wszControlName as CWSTR ) as long
Declare Function GetWinformsXClassName( byval nControlType as long ) as CWSTR
Declare Function GetControlRECT( byval pCtrl as clsControl ptr ) as RECT
Declare Function ShowDefaultPropertyForControl( byval pCtrl as clsControl ptr ) as Long
Declare Function DisplayPropertyList( byval pDoc as clsDocument ptr ) as Long
Declare Function HidePropertyListControls() as long
Declare Function InitiatePropertyValueChange() as Long
Declare Function ShowPropertyListControl() as Long
Declare Function IsDesignerView( byval pDoc as clsDocument ptr ) as Boolean
Declare Function ResetClosestControls( byval pDoc as clsDocument ptr ) as long
Declare Function SetBlueSnapLine( byval pDoc as clsDocument ptr, byval x1 as long, byval y1 as long, byval x2 as long, byval y2 as long ) as long
Declare Function SetClosestControls( byval pDoc as clsDocument ptr, byval pCtrlActive as clsControl ptr ) As clsControl ptr
Declare Function DisplayControlPopupMenu( byval hwnd as hwnd, byval xpos as long, byval ypos as long) As Long
Declare Function DisplayFormPopupMenu( byval hwnd as hwnd, byval xpos as long, byval ypos as long) As Long
Declare Function SetMouseCursor() As Long
Declare Function SetGrabHandleMouseCursor( byval pDoc as clsDocument ptr, byval x as long, byval y as long, byref pCtrlAction as clsControl Ptr ) as LRESULT
Declare Function CalculateGrabHandles( byval pDoc as clsDocument ptr) as long
Declare Function DrawGrabHandles( byval hDC as HDC, byval pDoc as clsDocument ptr, byval bFormOnly as Boolean ) as long
Declare Function HandleDesignerLButtonDoubleClick( ByVal HWnd As HWnd ) as LRESULT
Declare Function HandleDesignerLButtonDown( ByVal HWnd As HWnd ) as LRESULT
Declare Function HandleDesignerLButtonUp( ByVal HWnd As HWnd ) as LRESULT
Declare Function HandleDesignerRButtonDown( ByVal HWnd As HWnd ) as LRESULT
Declare Function HandleDesignerMouseMove( ByVal HWnd As HWnd ) as LRESULT
Declare Function DesignerForm_OnPaint( ByVal HWnd As HWnd) As LRESULT
Declare Function DesignerForm_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function DesignerForm_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function DesignerForm_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function Control_SubclassProc( BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM, BYVAL uIdSubclass AS UINT_PTR, BYVAL dwRefData AS DWORD_PTR ) AS LRESULT
Declare Function FormatCodetip( byval strCodeTip as string ) as STRING
Declare Function MaskStringCharacters( byval st as string) as string
Declare Function RemoveComments( byval st as string) as string
Declare Function DetermineWithVariable( byval pDoc as clsDocument ptr) as string
Declare Function DereferenceLine( byval pDoc as clsDocument ptr, byval sTrigger as String ) as DB2_DATA ptr
Declare Function ShowCodetip( byval pDoc as clsDocument ptr) as BOOLEAN
Declare Function ShowAutoCompletePopup( byval hEdit as hwnd, byref sList as string ) as Long
Declare Function ExistsInAutocompleteList( byref sList as string, byref sMatch as string ) as long
Declare Function ShowAutocompleteList( byval Notification as long = 0) as BOOLEAN
Declare Function GetMRUMenuHandle() as HMENU
Declare Function OpenMRUFile( ByVal HWnd As HWnd, ByVal wID As Long ) As Long
Declare Function ClearMRUlist( ByVal wID As Long ) As Long
Declare Function CreateMRUpopup() As HMENU
Declare Function UpdateMRUMenu( ByVal hMenu As HMENU ) As Long
Declare Function UpdateMRUList( Byref wzFilename As WString ) As Long
Declare Function GetMRUProjectMenuHandle() As HMENU
Declare Function OpenMRUProjectFile( ByVal HWnd As HWnd, ByVal wID As Long) As Long
Declare Function UpdateMRUProjectMenu( ByVal hMenu As HMENU ) As Long
Declare Function UpdateMRUProjectList( Byval wszFilename As CWSTR ) As Long
Declare Function LoadHelpTreeview( byval hTreeview as hwnd, byval sPath as CWSTR, byval hParent as HTREEITEM ) as long
Declare Function ShowContextHelp( byval id as long ) As Long
