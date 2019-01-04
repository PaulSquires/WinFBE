Declare Function frmMain_GotoDefinition( ByVal pDoc As clsDocument Ptr ) As Long
Declare Function frmMain_GotoLastPosition() As Long
Declare Function frmMain_UpdateLineCol( ByVal HWnd As HWnd ) As Long
Declare Function frmMain_ProcessCommandLine( ByVal HWnd As HWnd ) As Long
Declare Function Scintilla_OnNotify( ByVal HWnd As HWnd, ByVal pNSC As SCNOTIFICATION Ptr ) As Long
Declare Function frmMain_SetFocusToCurrentCodeWindow() As Long
Declare Function frmMain_OnPaint( ByVal HWnd As HWnd ) As LRESULT
Declare Function frmMain_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
Declare Function frmMain_OpenProjectSafely( ByVal HWnd As HWnd, byref wszProjectFileName as const WString ) as Boolean
Declare Function frmMain_OpenFileSafely( ByVal HWnd As HWnd, ByVal bIsNewFile As BOOLEAN, ByVal bIsTemplate As BOOLEAN, ByVal bShowInTab As BOOLEAN, byval bIsInclude as BOOLEAN, ByVal pwszName As WString Ptr, ByVal pDocIn As clsDocument Ptr, byval bIsDesigner as Boolean = false ) As clsDocument Ptr
Declare Function frmMain_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmMain_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
Declare Function frmMain_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmMain_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
Declare Function frmMain_OnActivateApp(ByVal HWnd As HWnd, ByVal fActivate As BOOLEAN, ByVal dwThreadId As DWORD) As LRESULT
Declare Function frmMain_OnContextMenu( ByVal HWnd As HWnd, ByVal hwndContext As HWnd, ByVal xPos As Long, ByVal yPos As Long ) As LRESULT
Declare Function frmMain_OnDropFiles( ByVal HWnd As HWnd, ByVal hDrop As HDROP ) As LRESULT
Declare Function frmMain_OnClose(ByVal HWnd As HWnd) As LRESULT
Declare Function frmMain_OnDestroy(byval HWnd As HWnd) As LRESULT
Declare Function frmMain_TabCtl_SubclassProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM, ByVal uIdSubclass As UINT_PTR, ByVal dwRefData As DWORD_PTR ) As LRESULT
Declare Function frmMain_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmMain_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function SetCompileStatusBarMessage( byref wszText as const wstring, byval hIconCompile as HICON ) as LRESULT
Declare Function CreateTempResourceFile() as boolean
Declare Function ResetScintillaCursors() as Long
Declare Function code_Compile( ByVal wID As Long ) As BOOLEAN
Declare Function CreateChildProcess( byref zCmdLine AS wstring, byref zParams AS wstring, byval hChildStdOutWrite AS HANDLE ) AS BOOLEAN
Declare Sub RedirConsoleToFile( byref szExe AS wstring, byref szCmdLine AS wstring, byref sConsoleText AS string )
Declare Function RunEXE( ByRef wszFileExe As CWSTR, ByRef wszParam As CWSTR ) As Long
Declare Function SetDocumentErrorPosition( ByVal hLV As HWnd, Byval wID as long ) As Long
Declare Function ParseLogForError( byref wsLogSt as CWSTR, byval bAllowSuccessMessage as Boolean, Byval wID as long, byval fCompile64 as Boolean ) as Boolean
Declare Function JulianDateNow() as long
Declare Function ConvertWinFBEversion( byref wszVersion as wstring ) as long
Declare Function DisableAllModeless() as long
Declare Function EnableAllModeless() as long
Declare Function GetTemporaryFilename( byref wszFolder as wstring, BYREF wszExtension AS wSTRING) AS string
Declare Function ComboBox_ReplaceString(BYVAL hComboBox AS HWND, BYVAL index AS LONG, BYVAL pwszNewText AS WSTRING PTR, BYVAL pNewData AS LONG_PTR = 0) AS LONG
Declare Function GetFontCharSetID(ByREF wzCharsetName As CWSTR ) As Long
Declare Function RemoveDuplicateSpaces( byref sText as const string) as string
Declare Function ConvertCase( byval sText as string) as string
Declare Function Utf8ToAscii(byref strUtf8 AS STRING) AS STRING
Declare Function AnsiToUtf8(BYREF sAnsi AS STRING) AS STRING
Declare Function Utf8ToUnicode(BYREF ansiStr AS CONST STRING) AS STRING
Declare Function UnicodeToUtf8(byval pswzUnicode as wstring ptr) AS STRING
Declare Function FileEncodingTextDescription(byval FileEncoding as long) as CWSTR
Declare Function GetFileToString( byref wszFilename as const wstring, byref txtBuffer as string, byval pDoc as clsDocument ptr ) as boolean
Declare Function ConvertTextBuffer( byval pDoc as clsDocument ptr, byval FileEncoding as long ) as Long
Declare Function IsCurrentLineIncludeFilename() as Boolean
Declare Function OpenSelectedDocument( byref wszFilename as wstring, byref wszFunctionName as WSTRING = "", byval nLineNumber as long = 0 ) as clsDocument ptr
Declare Function ProcessToCurdrive( Byval wzFilename As CWSTR ) As CWSTR
Declare Function ProcessFromCurdrive( Byval wzFilename As CWSTR ) As CWSTR
Declare Function Treeview_RemoveCheckBox( byval hTree as hwnd, byval hNode as HTREEITEM) as long
Declare Function FF_TreeView_InsertItem( ByVal hWndControl As HWnd, ByVal hParent As HANDLE, ByRef TheText As WString, ByVal lParam As LPARAM = 0, ByVal iImage As Long = 0, ByVal iSelectedImage As Long = 0 ) As HANDLE
Declare Function FF_TreeView_GetlParam( ByVal hWndControl As HWnd, ByVal hItem As HANDLE ) As Long
Declare Function FF_TreeView_SetCheckState( ByVal hWndControl As HWnd, ByVal hItem As HANDLE, ByVal fCheck As Boolean ) As BOOLEAN
Declare Function FF_TreeView_GetCheckState( ByVal hWndControl As HWnd, ByVal hItem As HANDLE ) As BOOLEAN
Declare Function FF_TreeView_SetlParam(ByVal hWndControl as HWnd, ByVal hItem as HANDLE, ByVal lParam as Long) as Long
Declare Function AfxIFileOpenDialogW( ByVal hwndOwner As HWnd, ByVal idButton As Long) As WString Ptr
Declare Function AfxIFileOpenDialogMultiple( ByVal hwndOwner As HWnd, ByVal idButton As Long) As IShellItemArray Ptr
Declare Function AfxIFileSaveDialog( ByVal hwndOwner As HWnd, ByVal pwszFileName As WString Ptr, ByVal pwszDefExt As WString Ptr, ByVal id As Long = 0, ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH ) As WString Ptr
Declare Function FF_Toolbar_EnableButton(ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
Declare Function FF_Toolbar_DisableButton(ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
Declare Function FF_ListView_InsertItem( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal lParam As LPARAM = 0 ) As BOOLEAN
Declare Function FF_ListView_GetItemText( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal nTextMax As Long ) As BOOLEAN
Declare Function FF_ListView_SetItemText( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal nTextMax As Long ) As Long
Declare Function FF_ListView_GetlParam( ByVal hWndControl As HWnd, ByVal iRow As Long ) As LPARAM
Declare Function FF_ListView_SetlParam( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal ilParam As LPARAM ) As long
Declare Function LoadLocalizationFile( Byref wszFileName As CWSTR, byval IsEnglish as boolean = false ) As BOOLEAN
Declare Function GetProcessImageName( ByVal pe32w As PROCESSENTRY32W Ptr, ByVal pwszExeName As WString Ptr ) As Long
Declare Function IsProcessRunning( ByVal pwszExeFileName As WString Ptr ) As BOOLEAN
Declare Function GetRunExecutableFilename() as CWSTR
Declare Function LoadPNGfromRes(BYVAL hInstance AS HINSTANCE, BYREF wszImageName AS WSTRING) as any ptr
Declare Function AfxStrRemoveWithMark(BYREF wszMainStr AS CONST WSTRING, BYREF wszDelim1 AS CONST WSTRING, BYREF wszDelim2 AS CONST WSTRING, BYREF MarkKeys AS CONST WSTRING="", BYVAL fRemoveAll AS BOOLEAN = FALSE, BYVAL IsInstrRev AS BOOLEAN = FALSE) AS CWSTR
Declare Function ParseDocument( byval pDoc as clsDocument ptr, byval wszFilename as CWSTR ) As Long
Declare Function frmMain_BuildMenu( ByVal pWindow As CWindow Ptr ) As HMENU
Declare Function frmMain_ChangeTopMenuStates() As Long
Declare Function AddProjectFileTypesToMenu( ByVal hPopUpMenu As HMENU, ByVal pDoc As clsDocument Ptr, byval fNoSeparator as BOOLEAN = false ) As Long
Declare Function CreateStatusBarFileTypeContextMenu() As HMENU
Declare Function CreateStatusBarFileEncodingContextMenu() As HMENU
Declare Function CreateTopTabCtlContextMenu( ByVal idx As Long ) As HMENU
Declare Function CreateExplorerContextMenu( ByVal pDoc As clsDocument Ptr ) As HMENU
Declare Function CreateScintillaContextMenu() As HMENU
Declare Function PositionExplorerWindows( ByVal HWnd As HWnd ) As LRESULT
Declare Function AddFunctionsToExplorerTreeview( ByVal pDoc As clsDocument Ptr, ByVal fUpdateNodes As BOOLEAN ) As Long
Declare Function DoCheckForUpdates( byval hWndParent as hwnd, byval bSilentCheck as Boolean = false ) as long
Declare Function DisplayPropertyDetails() as Long
Declare Function DisplayEventDetails() as Long
Declare Function GetRGBColorFromProperty( byval pProp as clsProperty ptr ) as COLORREF
Declare Function SetFontClassFromPropValue( byref wszPropValue as wstring ) as CWSTR
Declare Function SetLogFontFromPropValue( byref wszPropValue as wstring ) as LOGFONT
Declare Function SetPropValueFromLogFont( byref lf as LOGFONT ) as CWSTR
Declare Function ChooseFontForProperty( byval pProp as clsProperty ptr ) as long
Declare Function CreateDefaultFontPropValue() as CWSTR
Declare Function IsPropertyExists( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR ) as boolean
Declare Function IsEventExists( byval pCtrl as clsControl ptr, byval wszEventName as CWSTR ) as boolean
Declare Function GetControlPropertyPtr( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR ) as clsProperty Ptr
Declare Function GetControlProperty( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR ) as CWSTR
Declare Function SetControlProperty( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR, byval wszPropValue as CWSTR ) as long
Declare Function AddControlEvent( byval pCtrl as clsControl ptr, byref wszEventName as CWSTR, byval bIsSelected as boolean = false ) as Long
Declare Function SetControlEvent( byval pCtrl as clsControl ptr, byval wszEventName as CWSTR, byval bIsSelected as boolean ) as long
Declare Function AttachDefaultControlEvents( byval pCtrl as clsControl ptr ) as Long
Declare Function AddControlProperty( byval pCtrl as clsControl ptr, byref wszPropName as CWSTR, byref wszPropValue as CWSTR, byval nPropType as PropertyType ) as Long
Declare Function AttachDefaultControlProperties( byval pCtrl as clsControl ptr ) as Long
Declare Function LoadPropertyComboListbox( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function Form_ApplyProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function Button_ApplyProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function Label_ApplyProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function CheckBox_ApplyProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function OptionButton_ApplyProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function ListBox_ApplyProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function ComboBox_ApplyProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function TextBox_ApplyProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function MaskedEdit_ApplyProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function Frame_ApplyProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr, byval pProp as clsProperty ptr ) as Long
Declare Function ApplyControlProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr ) as long
Declare Function GetFormName( byval pDoc as clsDocument ptr ) as CWSTR
Declare Function GenerateFormMetaData( byval pDoc as clsDocument ptr ) as long
Declare Function GenerateMainMenuCode( byval pDoc as clsDocument ptr ) as CWSTR
Declare Function GenerateFormCode( byval pDoc as clsDocument ptr ) as long
Declare Function SetColorListBoxSelection( byref wszPropValue as wstring ) as Long
Declare Function frmVDColors_Show( ByVal hWndParent As HWnd, byref wszPropValue as wstring ) as Long
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
