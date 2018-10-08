declare function IsPropertyExists( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR) as boolean
declare function GetActivePropertyPtr() as clsProperty ptr
declare function GetActiveEventPtr() as clsEvent ptr
declare function GenerateFormMetaData( byval pDoc as clsDocument ptr ) as long
declare function GenerateFormCode( byval pDoc as clsDocument ptr ) as long
declare function AttachDefaultControlProperties( byval pCtrl as clsControl ptr ) as Long
declare function GetControlType( byval wszControlName as CWSTR ) as long
declare function GetControlName( byval nControlType as long ) as CWSTR
declare function IsDesignerView( byval pDoc as clsDocument ptr ) as Boolean
declare function RemoveComments( byval st as string) as string
declare function MaskStringCharacters( byval st as string) as string
declare function ChangeCommentCharacters( byval st as string) as string
declare Function DisplayPropertyList( byval pDoc as clsDocument ptr ) as Long
declare function KeyboardResizeControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long
declare function KeyboardMoveControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long
declare function KeyboardCycleActiveControls( byval pDoc as clsDocument ptr, byval vKeycode as long ) as Long
declare function GetControlRECT( byval pCtrl as clsControl ptr ) as RECT
declare function ApplyControlProperties( byval pDoc as clsDocument ptr, byval pCtrl as clsControl ptr ) as Long
declare function GetControlProperty( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR) as CWSTR
declare function SetControlProperty( byval pCtrl as clsControl ptr, byval wszPropName as CWSTR, byval wszPropValue as CWSTR) as long
declare function SetControlEvent( byval pCtrl as clsControl ptr, byval wszEventName as CWSTR, byval bIsSelected as boolean ) as long
declare function GetActiveToolboxControlName() as CWSTR
declare function GetActiveToolboxControlClassName() as CWSTR
declare function GetActiveToolboxControlType() as long
declare function CreateToolboxControl( byval pDoc as clsDocument ptr, byval ControlType as long, byref rcDraw as RECT) as clsControl ptr
declare function GetActiveToolboxControl() as Long
declare function SetActiveToolboxControl( byval ControlType as long ) as Long
declare Function frmVDToolbox_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0) As Long
declare FUNCTION Control_SubclassProc( BYVAL hwnd AS HWND, BYVAL uMsg AS UINT, BYVAL wParam AS WPARAM, BYVAL lParam AS LPARAM, BYVAL uIdSubclass AS UINT_PTR, BYVAL dwRefData AS DWORD_PTR ) AS LRESULT
declare Function DesignerMain_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As LRESULT
declare Function DesignerFrame_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As LRESULT
declare Function DesignerForm_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As LRESULT
declare Function Lasso_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As LRESULT
declare function ParseLogForError( byref wsLogSt as CWSTR, byval bAllowSuccessMessage as Boolean, Byval wID as long) as Boolean
declare SUB RedirConsoleToFile( byref szExe AS wstring, byref szCmdLine AS wstring, byref sConsoleText AS string)
Declare Function GetTemporaryFilename( byref wszFolder as wstring, BYREF wszExtension AS wSTRING) AS string
Declare Function GetFontCharSetID(ByREF wzCharsetName As CWSTR ) As Long
Declare Function RemoveDuplicateSpaces( byref sText as const string) as string
Declare Function ConvertCase( byval sText as string) as string
Declare Function Utf8ToAscii(byref strUtf8 AS STRING) AS STRING
Declare Function AnsiToUtf8(BYREF sAnsi AS STRING) AS STRING
Declare Function Utf8ToUnicode(BYREF ansiStr AS CONST STRING) AS STRING
Declare Function UnicodeToUtf8(byval pswzUnicode as wstring ptr) AS STRING
Declare Function FileEncodingTextDescription(byval FileEncoding as long) as CWSTR
Declare Function GetFileToString( byref wszFilename as const wstring, byref txtBuffer as string, byval pDoc as clsDocument ptr) as boolean
Declare Function ConvertTextBuffer( byval pDoc as clsDocument ptr, byval FileEncoding as long ) as Long
Declare Function OpenSelectedDocument( byval pDoc as clsDocument ptr, byref wszName as WSTRING, ByVal hCurrentNode As HTREEITEM =0) as long
Declare Function ProcessToCurdrive( ByRef wzFilename As CWSTR ) As CWSTR
Declare Function ProcessFromCurdrive( ByRef wzFilename As CWSTR ) As CWSTR
Declare Function Treeview_RemoveCheckBox( byval hTree as hwnd, byval hNode as HTREEITEM) as long
Declare Function FF_TreeView_InsertItem( ByVal hWndControl As HWnd, ByVal hParent As HANDLE, ByRef TheText As WString, ByVal lParam As LPARAM = 0, ByVal iImage As Long = 0, ByVal iSelectedImage As Long = 0 ) As HANDLE
Declare Function FF_TreeView_GetlParam( ByVal hWndControl As HWnd, ByVal hItem As HANDLE ) As Long
Declare Function FF_TreeView_GetiImage( ByVal hWndControl As HWnd, ByVal hItem As HANDLE ) As Long
Declare Function FF_TreeView_SetCheckState( ByVal hWndControl As HWnd, ByVal hItem As HANDLE, ByVal fCheck As Boolean ) As BOOLEAN
Declare Function FF_TreeView_GetCheckState( ByVal hWndControl As HWnd, ByVal hItem As HANDLE ) As BOOLEAN
Declare Function FF_TreeView_SetlParam (ByVal hWndControl as HWnd, ByVal hItem as HANDLE, ByVal lParam as Long) as Long
Declare Function AfxIFileOpenDialogW( ByVal hwndOwner As HWnd, ByVal idButton As Long) As WString Ptr
Declare Function AfxIFileOpenDialogMultiple( ByVal hwndOwner As HWnd, ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH) As IShellItemArray Ptr
Declare Function AfxIFileSaveDialog( ByVal hwndOwner As HWnd, ByVal pwszFileName As WString Ptr, ByVal pwszDefExt As WString Ptr, ByVal id As Long = 0, ByVal sigdnName As SIGDN = SIGDN_FILESYSPATH ) As WString Ptr
Declare Function FF_Toolbar_EnableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
Declare Function FF_Toolbar_DisableButton (ByVal hToolBar As HWnd, ByVal idButton As Long) As BOOLEAN
Declare Function FF_ListView_InsertItem( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal lParam As LPARAM = 0 ) As BOOLEAN
Declare Function FF_ListView_GetItemText( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal nTextMax As Long ) As BOOLEAN
Declare Function FF_ListView_SetItemText( ByVal hWndControl As HWnd, ByVal iRow As Long, ByVal iColumn As Long, ByVal pwszText As WString Ptr, ByVal nTextMax As Long ) As Long
Declare Function FF_ListView_GetlParam( ByVal hWndControl As HWnd, ByVal iRow As Long ) As LPARAM
Declare Function LoadLocalizationFile( Byref wszFileName As CWSTR ) As BOOLEAN
Declare Function GetProcessImageName( ByVal pe32w As PROCESSENTRY32W Ptr, ByVal pwszExeName As WString Ptr ) As Long
Declare Function IsProcessRunning( ByVal pwszExeFileName As WString Ptr ) As BOOLEAN
Declare Function GetRunExecutableFilename() as CWSTR
Declare Function Splitter_OnLButtonDown() As BOOLEAN
Declare Function Splitter_OnLButtonUp() As BOOLEAN
Declare Function Splitter_OnMouseMove() As BOOLEAN
Declare Function IsPreparsedFile( byref sFilename as string ) as boolean
Declare Function GetIncludeFilename( byref sFilename as CWSTR, byref sLine as string ) as CWSTR
Declare Function ScintillaGetLine( byval pDoc as clsDocument ptr, ByVal nLine As Long ) As String
Declare Function ParseDocument( byval idx as long, byval pDoc as clsDocument ptr, byval sFilename as CWSTR) As Long
Declare Function SetCompileStatusBarMessage( byref wszText as const wstring, byval hIconCompile as HICON ) as LRESULT
Declare Function code_Compile( ByVal wID As Long ) As BOOLEAN
Declare Function RunEXE( ByRef wszFileExe As CWSTR, ByRef wszParam As CWSTR ) As Long
Declare Function SetDocumentErrorPosition( ByVal hLV As HWnd, Byval wID as long ) As Long
Declare Function frmMain_BuildMenu ( ByVal pWindow As CWindow Ptr ) As HMENU
Declare Function frmMain_ChangeTopMenuStates() As Long
Declare Function AddProjectFileTypesToMenu( ByVal hPopUpMenu As HMENU, ByVal pDoc As clsDocument Ptr, byval fNoSeparator as BOOLEAN = false ) As Long
Declare Function CreateStatusBarFileTypeContextMenu() As HMENU
Declare Function CreateStatusBarFileEncodingContextMenu() As HMENU
Declare Function CreateTopTabCtlContextMenu( ByVal idx As Long ) As HMENU
Declare Function CreateExplorerContextMenu( ByVal pDoc As clsDocument Ptr ) As HMENU
Declare Function CreateScintillaContextMenu() As HMENU
Declare Function frmMain_CreateToolbar( ByVal pWindow As CWindow Ptr ) As BOOLEAN
Declare Function frmMain_DisableToolbarButtons() As Long
Declare Function frmMain_ChangeToolbarButtonsState() As Long
Declare Function GetMRUMenuHandle() as HMENU
Declare Function OpenMRUFile( ByVal HWnd As HWnd, ByVal wID As Long ) As Long
Declare Function ClearMRUlist( ByVal wID As Long ) As Long
Declare Function CreateMRUpopup() As HMENU
Declare Function UpdateMRUMenu( ByVal hMenu As HMENU ) As Long
Declare Function UpdateMRUList( ByVal wzFilename As WString Ptr ) As Long
Declare Function GetMRUProjectMenuHandle() As HMENU
Declare Function OpenMRUProjectFile( ByVal HWnd As HWnd, ByVal wID As Long, ByVal pwszFilename As WString Ptr = 0 ) As Long
Declare Function UpdateMRUProjectMenu( ByVal hMenu As HMENU ) As Long
Declare Function UpdateMRUProjectList( ByVal wzFilename As WString Ptr ) As Long
Declare Function FormatCodetip( byval strCodeTip as string ) as STRING
declare function ShowCodetip( byval pDoc as clsDocument ptr) as BOOLEAN
Declare Function ShowAutocompleteList(byval Notification as long = 0) as BOOLEAN
Declare Function frmHotImgBtn_IsEnabled( ByVal HWnd As HWnd ) As boolean
Declare Function frmHotImgBtn_Enabled( ByVal HWnd As HWnd, ByVal fEnabled as BOOLEAN ) As Long
Declare Function frmHotImgBtn_SetBackColors( ByVal HWnd As HWnd, ByVal clrNormal As COLORREF, byval clrHot as COLORREF ) As Long
Declare Function frmHotImgBtn_SetImages( ByVal HWnd As HWnd, ByRef wszImgNormal As WString = "", ByRef wszImgHot As WString = "", byval bIsGDI as Boolean = false ) As Long
Declare Function frmHotImgBtn_SetSelected( ByVal HWnd As HWnd, ByVal fSelected As BOOLEAN) As Long
Declare Function frmHotImgBtn_GetSelected( ByVal HWnd As HWnd ) As BOOLEAN
Declare Function frmHotImgBtn_OnLButtonUp(ByVal HWnd As HWnd, ByVal x As Long, ByVal y As Long, ByVal keyFlags As Long ) As LRESULT
Declare Function frmHotImgBtn_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmHotImgBtn_OnPaint( ByVal HWnd As HWnd) As LRESULT
Declare Function frmHotImgBtn_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmHotImgBtn( ByVal hWndParent As HWnd, ByVal cID As Long = 0, ByRef wszImgNormal As WString = "", ByRef wszImgHot As WString = "", ByRef wszTooltip As WString = "", ByVal nImgWidth As Long = 24, ByVal nImgHeight As Long = 24, ByVal clrBg As COLORREF = 0, ByVal clrBgHot As COLORREF = 0, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0, byval fSelectedFrame as BOOLEAN = true, byval fEnabled as BOOLEAN = true, byval fIsGDI as Boolean = false ) As HWND
Declare Function frmHotTxtBtn_OnLButtonUp(ByVal HWnd As HWnd, ByVal x As Long, ByVal y As Long, ByVal keyFlags As Long ) As LRESULT
Declare Function frmHotTxtBtn_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmHotTxtBtn_OnPaint( ByVal HWnd As HWnd) As LRESULT
Declare Function frmHotTxtBtn_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmHotTxtBtn( ByVal hWndParent As HWnd, ByVal cID As Long = 0, ByRef wszText As WString = "", ByRef wszTooltip As WString = "", ByVal clrFg As COLORREF = 0, ByVal clrBg As COLORREF = 0, ByVal clrFgHot As COLORREF = 0, ByVal clrBgHot As COLORREF = 0, ByVal nOrientation As Long = 1, BYVAL x AS LONG = 0, BYVAL y AS LONG = 0, BYVAL nWidth AS LONG = 0, BYVAL nHeight AS LONG = 0 ) As HWND
Declare Function LoadRecentFilesTreeview( ByVal HWnd As HWnd ) As LRESULT
Declare Function PositionRecentWindows( ByVal HWnd As HWnd ) As LRESULT
Declare Function frmRecent_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
Declare Function frmRecent_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmRecent_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmRecent_Tree_SubclassProc ( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM, ByVal uIdSubclass As UINT_PTR, ByVal dwRefData As DWORD_PTR ) As LRESULT
Declare Function frmRecent_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmRecent_Show( ByVal hWndParent As HWnd ) As LRESULT
Declare Function PositionExplorerWindows( ByVal HWnd As HWnd ) As LRESULT
Declare Function AddFunctionsToExplorerTreeview( ByVal pDoc As clsDocument Ptr, ByVal fUpdateNodes As BOOLEAN ) As Long
Declare Function frmExplorer_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
Declare Function frmExplorer_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmExplorer_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
Declare Function frmExplorer_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmExplorer_Tree_SubclassProc ( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM, ByVal uIdSubclass As UINT_PTR, ByVal dwRefData As DWORD_PTR ) As LRESULT
Declare Function frmExplorer_OnPaint( ByVal HWnd As HWnd) As LRESULT
Declare Function frmExplorer_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmExplorer_Show( ByVal hWndParent As HWnd ) As LRESULT
Declare Function ExecuteUserTool( ByVal nToolNum As Long ) As Long
Declare Function CreateUserTools_AcceleratorTable() As Long
Declare Function LoadToolsListBox( byval hParent as hwnd ) as Long
Declare Function SwapToolsListBoxItems( byval Item1 as long, Byval Item2 as long) as Long
Declare Function SetUserToolsTextboxes() as long
Declare Function frmUserTools_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmUserTools_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmUserTools_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmUserTools_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmUserTools_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmUserTools_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function LoadBuildListBox( byval hParent as hwnd ) as Long
Declare Function GetSelectedBuildGUID() as String
Declare Function GetDefaultBuildGUID() as String
Declare Function LoadBuildComboBox() as Long
Declare Function SwapBuildListBoxItems( byval Item1 as long, Byval Item2 as long) as Long
Declare Function SetCompilerConfigTextboxes() as long
Declare Function frmCompileConfig_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmCompileConfig_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmCompileConfig_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmCompileConfig_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmCompileConfig_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmCompileConfig_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function frmOutput_ShowNotes() as long
Declare Function UpdateToDoListview() as long
Declare Function DoSearchFileSelected() as LONG
Declare Function ShowHideOutputControls( ByVal HWnd As HWnd ) As LRESULT
Declare Function PositionOutputWindows( ByVal HWnd As HWnd ) As LRESULT
Declare Function frmOutput_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
Declare Function frmOutput_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmOutput_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
Declare Function frmOutput_OnPaint( ByVal HWnd As HWnd) As LRESULT
Declare Function frmOutput_Listview_SubclassProc ( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM, ByVal uIdSubclass As UINT_PTR, ByVal dwRefData As DWORD_PTR ) As LRESULT
Declare Function frmOutput_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOutput_Show( ByVal hWndParent As HWnd ) As LRESULT
Declare Function frmOptionsGeneral_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptionsGeneral_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function frmOptionsEditor_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptionsEditor_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function EnumFontName(lf As LOGFONTW, tm As TEXTMETRIC, ByVal FontType As Long, HWnd As HWnd) As Long
Declare Sub FillFontCombo( ByVal HWnd As HWnd)
Declare Function DrawFontCombo(ByVal HWnd As HWnd, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As Long
Declare Sub FillFontSizeCombo( ByVal hCb As HWnd, ByVal strFontName As WString Ptr )
Declare Sub FillFontCharSets( ByVal hCtl As HWnd )
Declare Function SetTempColorSelection( ByVal HWnd As HWnd, ByVal nCtrlID As Long ) As Long
Declare Function ShowComboColors( ByVal HWnd As HWnd ) As Long
Declare Function frmOptionsColors_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmOptionsColors_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptionsColors_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function frmOptionsCompiler_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmOptionsCompiler_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptionsCompiler_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Sub SaveEditorOptions()
Declare Function frmOptions_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmOptions_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmOptions_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
Declare Function frmOptions_OnCtlColorStatic(HWnd As HWnd, hdc As HDC, hWndChild As HWnd, nType As Long) As HBRUSH
Declare Function frmOptions_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmOptions_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmOptions_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmOptions_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function frmTemplates_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmTemplates_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmTemplates_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmTemplates_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmTemplates_WndProc (ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As LRESULT
Declare Function frmTemplates_Show (ByVal hParent As HWnd, ByVal x As Long, ByVal y As Long) As Long
Declare Function frmFnList_SetListBoxPosition() As Long
Declare Function frmFnList_UpdateListBox() As Long
Declare Function frmFnList_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmFnList_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
Declare Function frmFnList_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
Declare Function frmFnList_OnClose( ByVal HWnd As HWnd ) As LRESULT
Declare Function frmFnList_OnMeasureItem( ByVal HWnd As HWnd, ByVal lpmis As MEASUREITEMSTRUCT Ptr ) As Long
Declare Function frmFnList_OnDrawItem( ByVal HWnd As HWnd, ByVal lpdis As Const DRAWITEMSTRUCT Ptr ) As Long
Declare Function frmFnList_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmFnList_ListBox_SubclassProc ( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM, ByVal uIdSubclass As UINT_PTR, ByVal dwRefData As DWORD_PTR ) As LRESULT
Declare Function frmFnList_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmFnList_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function frmGoto_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmGoto_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmGoto_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmGoto_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmGoto_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmGoto_Show( ByVal hWndParent As HWnd ) As Long
Declare Function frmFindInFiles_AddToFindCombo( ByRef sText As Const String ) As Long
Declare Function frmFindInFiles_AddToFilesCombo( ByRef sText As Const String ) As Long
Declare Function frmFindInFiles_AddToFolderCombo( ByRef sText As Const String ) As Long
Declare Function ParseSearchResults( byref wszResultFile as wstring ) as long
Declare Function DoFindInFilesEx() as LONG
Declare Function frmFindInFiles_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmFindInFiles_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmFindInFiles_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmFindInFiles_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmFindInFiles_Show( ByVal hWndParent As HWnd ) As Long
Declare Function frmFindReplace_DoReplace( byval fReplaceAll as BOOLEAN = false ) as long
Declare Function frmFindReplace_NextSelection( byval StartPos as long, byval bGetNext as boolean ) as long
Declare Function frmFindReplace_HighlightSearches() as long
Declare Function frmFindReplace_SetButtonStates() as LONG
Declare Function frmFindReplace_ShowControls() as long
Declare Function frmFindReplace_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmFindReplace_OnPaint( ByVal HWnd As HWnd) As LRESULT
Declare Function frmFindReplace_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmFindReplace_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmFindReplace_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmFindReplace_Show( ByVal hWndParent As HWnd, byval fShowReplace as BOOLEAN ) As Long
Declare Function CreateResourceManifest( byval idx as long ) as Long
Declare Function SaveProjectOptions( ByVal HWnd As HWnd ) As BOOLEAN
Declare Function frmProjectOptions_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmProjectOptions_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmProjectOptions_OnClose(HWnd As HWnd) As LRESULT
Declare Function frmProjectOptions_OnDestroy(HWnd As HWnd) As LRESULT
Declare Function frmProjectOptions_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmProjectOptions_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0, byval IsNewProject as boolean = false ) As Long
Declare Function frmMain_GotoDefinition( ByVal pDoc As clsDocument Ptr ) As Long
Declare Function frmMain_GotoLastPosition() As Long
Declare Function frmMain_UpdateLineCol( ByVal HWnd As HWnd) As Long
Declare Function frmMain_ProcessCommandLine( ByVal HWnd As HWnd) As Long
Declare Function Scintilla_OnNotify( ByVal HWnd As HWnd, ByVal pNSC As SCNOTIFICATION Ptr ) As Long
Declare Function frmMain_SetFocusToCurrentCodeWindow() As Long
Declare Function frmMain_PositionWindows( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False) As LRESULT
Declare Function OnCommand_ProjectClose( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectNew( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_ProjectOpen( ByVal HWnd As HWnd ) As LRESULT
Declare Function frmMain_OpenFileSafely( ByVal HWnd As HWnd, ByVal bIsNewFile As BOOLEAN, ByVal bIsTemplate As BOOLEAN, _
                           ByVal bShowInTab As BOOLEAN, byval bIsInclude as BOOLEAN, ByVal pwszName As WString Ptr, _
                           ByVal pDocIn As clsDocument Ptr, byval bIsDesigner as Boolean = false, byval ProjectFileType as Long=-1,byval idx as Long=-1) As clsDocument Ptr
Declare Function OnCommand_FileNew( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileOpen( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_OpenIncludeFile( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileSave( ByVal HWnd As HWnd, ByVal bSaveAs As BOOLEAN = False) As LRESULT
Declare Function OnCommand_FileSaveDeclares( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileSaveAll( ByVal HWnd As HWnd ) As LRESULT
Declare Function OnCommand_FileClose( ByVal HWnd As HWnd, ByVal veFileClose As eFileClose = EFC_CLOSECURRENT ) As LRESULT
Declare Function frmMain_OnCreate(ByVal HWnd As HWnd, ByVal lpCreateStructPtr As LPCREATESTRUCT) As BOOLEAN
Declare Function frmMain_OnSize(ByVal HWnd As HWnd, ByVal state As UINT, ByVal cx As Long, ByVal cy As Long) As LRESULT
Declare Function frmMain_OnCommand(ByVal HWnd As HWnd, ByVal id As Long, ByVal hwndCtl As HWnd, ByVal codeNotify As UINT) As LRESULT
Declare Function frmMain_OnNotify(ByVal HWnd As HWnd, ByVal id As Long, ByVal pNMHDR As NMHDR Ptr) As LRESULT
Declare Function frmMain_OnActivateApp(ByVal HWnd As HWnd, ByVal fActivate As BOOLEAN, ByVal dwThreadId As DWORD) As LRESULT
Declare Function frmMain_OnContextMenu( ByVal HWnd As HWnd, ByVal hwndContext As HWnd, ByVal xPos As Long, ByVal yPos As Long ) As LRESULT
Declare Function frmMain_OnDropFiles( ByVal HWnd As HWnd, ByVal hDrop As HDROP ) As LRESULT
Declare Function frmMain_OnClose(ByVal HWnd As HWnd) As LRESULT
Declare Function frmMain_OnDestroy(byval HWnd As HWnd) As LRESULT
Declare Function frmMain_TabCtl_SubclassProc ( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM, ByVal uIdSubclass As UINT_PTR, ByVal dwRefData As DWORD_PTR ) As LRESULT
Declare Function frmMain_WndProc( ByVal HWnd As HWnd, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM ) As LRESULT
Declare Function frmMain_Show( ByVal hWndParent As HWnd, ByVal nCmdShow As Long = 0 ) As Long
Declare Function WinMain( ByVal hInstance As HINSTANCE, ByVal hPrevInstance As HINSTANCE, ByVal szCmdLine As ZString Ptr, ByVal nCmdShow As Long ) As Long
Declare Function ImageIndex4ExplorerTreeview( ByVal pDoc As clsDocument Ptr ) As long
Declare Function WStrIsEqual(Byref lpWord As wString,Byref lpList As wString,ByVal fMatchCase As BOOLEAN=FALSE) As BOOLEAN
Declare Function WStrNIsEqual(Byref lpWord As wString,Byref lpList As wString,Byref Size As long,ByVal fMatchCase As BOOLEAN=FALSE) As BOOLEAN
Declare Function StrIsEqual(Byref lpWord As String,Byref lpList As String,ByVal fMatchCase As BOOLEAN=FALSE) As BOOLEAN
Declare Function StrNIsEqual(Byref lpWord As String,Byref lpList As String,Byref Size As long,ByVal fMatchCase As BOOLEAN=FALSE) As BOOLEAN
Declare sub ChangeImage4ExplorerTreeview2doc( ByVal pDoc As clsDocument Ptr )
Declare sub ChangeImage4ExplorerTreeview2db(ByVal hWndControl As HWnd,  ByVal hNode As HTREEITEM,  ByVal DB2ID as LONG )
Declare function GetClsDocument4ExplorerTreeview( ByVal hWndControl As HWnd, byVal hNode As HTREEITEM ) as clsDocument ptr
Declare function CreateName4ExplorerTreeview(ByVal hWndControl As HWnd,byref ReturnValue as WSTRing, ByVal hCurrentNode As HTREEITEM =0) as LONG
Declare FUNCTION AfxStrRemoveWithMark (BYREF wszMainStr AS CONST WSTRING, BYREF wszDelim1 AS CONST WSTRING, BYREF wszDelim2 AS CONST WSTRING, BYREF MarkKeys AS CONST WSTRING="", BYVAL fRemoveAll AS BOOLEAN = FALSE, BYVAL IsInstrRev AS BOOLEAN = FALSE) AS CWSTR
Declare FUNCTION AfxStrMid (BYREF wszMainStr AS CONST WSTRING, BYREF wszDelim1 AS CONST WSTRING, BYREF wszDelim2 AS CONST WSTRING, BYVAL Index AS long = 1) AS CWSTR
Declare function CheckIsDoc4ExplorerTreeview( ByVal hWndControl As HWnd, byVal hNode As HTREEITEM ) as Boolean
Declare function GetProjectIndex4ExplorerTreeview( ByVal hWndControl As HWnd, byVal hNode As HTREEITEM ) as long
Declare function GetProjecthNode4ExplorerTreeview( ByVal hWndControl As HWnd, byVal hNode As HTREEITEM ) as HTREEITEM