' ########################################################################################
' Microsoft Windows
' File: AfxRichEdit.inc
' Contents: Rich Edit control wrapper functions
' Compiler: Free Basic 32 & 64 bit
' Copyright (c) 2016 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "win/richedit.bi"
#include once "win/richole.bi"
#include once "Afx/CWSTR.inc"

' ========================================================================================
' Enables or disables automatic detection of URLs by a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_AutoUrlDetect (BYVAL hRichEdit AS HWND, BYVAL fUrlDetect AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_AUTOURLDETECT, fUrlDetect, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether a rich edit control can paste a specified clipboard format.
' ========================================================================================
PRIVATE FUNCTION RichEdit_CanPaste (BYVAL hRichEdit AS HWND, BYVAL clipformat AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_CANPASTE, clipformat, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether there are any actions in the control redo queue.
' ========================================================================================
PRIVATE FUNCTION RichEdit_CanRedo (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_CANREDO, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines whether there are any actions in an edit control's undo queue.
' ========================================================================================
PRIVATE FUNCTION RichEdit_CanUndo (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_CANUNDO, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about the character closest to a specified point in the client area of
' an edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_CharFromPos (BYVAL hRichEdit AS HWND, BYVAL lppl AS POINTL PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_CHARFROMPOS, 0, cast(LPARAM, lppl))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Displays a portion of the contents of a rich edit control, as previously formatted for a
' device using the EM_FORMATRANGE message.
' ========================================================================================
PRIVATE FUNCTION RichEdit_DisplayBand (BYVAL hRichEdit AS HWND, BYVAL lprc AS RECT PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_DISPLAYBAND, 0, cast(LPARAM, lprc))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Resets the undo flag of a rich edit control. The undo flag is set whenever an operation
' within the rich edit control can be undone.
' ========================================================================================
PRIVATE SUB RichEdit_EmptyUndoBuffer (BYVAL hRichEdit AS HWND)
   SendMessageW hRichEdit, EM_EMPTYUNDOBUFFER, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the starting and ending character positions of the selection in a rich edit control.
' ========================================================================================
PRIVATE SUB RichEdit_ExGetSel (BYVAL hRichEdit AS HWND, BYVAL lpchr AS CHARRANGE PTR)
   SendMessageW hRichEdit, EM_EXGETSEL, 0, cast(LPARAM, lpchr)
END SUB
' ========================================================================================

' ========================================================================================
' Sets an upper limit to the amount of text the user can type or paste into a rich edit control.
' ========================================================================================
PRIVATE SUB RichEdit_ExLimitText (BYVAL hRichEdit AS HWND, BYVAL dwLimit AS DWORD)
   SendMessageW hRichEdit, EM_EXLIMITTEXT, 0, dwLimit
END SUB
' ========================================================================================

' ========================================================================================
' Determines which line contains the specified character in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_ExLineFromChar (BYVAL hRichEdit AS HWND, BYVAL iIndex AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_EXLINEFROMCHAR, 0, iIndex)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects a range of characters and/or Component Object Model (COM) objects in a Microsoft
' Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_ExSetSel (BYVAL hRichEdit AS HWND, BYVAL lpcr AS CHARRANGE PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_EXSETSEL, 0, cast(LPARAM, lpcr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds text within a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_FindText (BYVAL hRichEdit AS HWND, BYVAL fOptions AS DWORD, BYVAL lpft AS FINDTEXTW PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_FINDTEXTW, fOptions, cast(LPARAM, lpft))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds text within a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_FindTextEx (BYVAL hRichEdit AS HWND, BYVAL fOptions AS DWORD, BYVAL lpftexw AS FINDTEXTEXW PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_FINDTEXTEXW, fOptions, cast(LPARAM, lpftexw))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Finds the next word break before or after the specified character position or retrieves
' information about the character at that position.
' ========================================================================================
PRIVATE FUNCTION RichEdit_FindWordBreak (BYVAL hRichEdit AS HWND, BYVAL fOperation AS DWORD, BYVAL dwStartPos AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_FINDWORDBREAK, fOperation, dwStartPos)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Formats a range of text in a rich edit control for a specific device.
' ========================================================================================
PRIVATE FUNCTION RichEdit_FormatRange (BYVAL hRichEdit AS HWND, BYVAL fRender AS LONG, BYVAL lpfr AS FORMATRANGE PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_FORMATRANGE, fRender, cast(LPARAM, lpfr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether the auto URL detection is turned on in the rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetAutoUrlDetect (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETAUTOURLDETECT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates the current state of the bidirectional options in the rich edit control.
' ========================================================================================
PRIVATE SUB RichEdit_GetBidiOptions (BYVAL hRichEdit AS HWND, BYVAL lpbo AS BIDIOPTIONS PTR)
   SendMessageW hRichEdit, EM_GETBIDIOPTIONS, 0, cast(LPARAM, lpbo)
END SUB
' ========================================================================================

' ========================================================================================
' Determines the current character formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetCharFormat (BYVAL hRichEdit AS HWND, BYVAL fOption AS DWORD, BYVAL lpcf AS CHARFORMAT PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETCHARFORMAT, fOption, cast(LPARAM, lpcf))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the Text Services Framework mode bias values for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetCTFModeBias (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETCTFMODEBIAS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Determines if the Text Services Framework (TSF) keyboard is open or closed.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetCTFOpenStatus (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETCTFOPENSTATUS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current edit style flags.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetEditStyle (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETEDITSTYLE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the event mask for a rich edit control. The event mask specifies which
' notification messages the control sends to its parent window.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetEventMask (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETEVENTMASK, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the zero-based index of the uppermost visible line in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetFirstVisibleLine (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETFIRSTVISIBLELINE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets information about hyphenation for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE SUB RichEdit_GetHyphenateInfo (BYVAL hRichEdit AS HWND, BYVAL lphi AS HYPHENATEINFO PTR)
   SendMessageW hRichEdit, EM_GETHYPHENATEINFO, cast(WPARAM, lphi), 0
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the Input Method Editor (IME) composition color. This message is available
' only in Asian-language versions of the operating system.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetIMEColor (BYVAL hRichEdit AS HWND, BYVAL rgCmpclr AS COMPCOLOR PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMECOLOR, 0, cast(LPARAM, rgCmpclr))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current IME mode for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetIMECompMode (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMECOMPMODE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the Input Method Editor (IME) composition text.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetIMECompText (BYVAL hRichEdit AS HWND, BYVAL lpict AS IMECOMPTEXT PTR, BYVAL buffer AS ANY PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMECOMPTEXT, cast(WPARAM, lpict), cast(LPARAM, buffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetIMEModeBias (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMEMODEBIAS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the current Input Method Editor (IME) options. This message is available only
' in Asian-language versions of the operating system.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetIMEOptions (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMEOPTIONS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the property and capabilities of the Input Method Editor (IME) associated with the
' current input locale.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetIMEProperty (BYVAL hRichEdit AS HWND, BYVAL figp AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETIMEPROPERTY, figp, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets a rich edit control's option settings for Input Method Editor (IME) and Asian
' language support.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetLangOptions (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETLANGOPTIONS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current text limit for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetLimitText (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETLIMITTEXT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Copies a line of text from a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetLine (BYVAL hRichEdit AS HWND, BYVAL which AS DWORD) AS CWSTR
   DIM buffer AS CWSTR = MKI(32765) + STRING(32765, 0)
   DIM n AS LONG = SendMessageW(hRichEdit, EM_GETLINE, which, cast(LPARAM, *buffer))
   RETURN LEFT(**buffer, n)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the number of lines in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetLineCount (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETLINECOUNT, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the state of a rich edit control's modification flag. The flag indicates whether
' the contents of the rich edit control have been modified.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetModify (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETMODIFY, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the state of a rich edit control's modification flag. The flag indicates whether
' the contents of the rich edit control have been modified.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetOleInterface (BYVAL hRichEdit AS HWND, BYVAL ppObject AS IUnknown PTR PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETOLEINTERFACE, 0, cast(LPARAM, @ppObject))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves rich edit control options.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetOptions (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETOPTIONS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deprecated. Gets the text layout for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetPageRotate (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETPAGEROTATE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the paragraph formatting of the current selection in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetParaFormat (BYVAL hRichEdit AS HWND, BYVAL pParaFmt AS PARAFORMAT PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETPARAFORMAT, 0, cast(LPARAM, pParaFmt))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the password character that a rich edit control displays when the user enters text.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetPasswordChar (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETPASSWORDCHAR, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current punctuation characters for the rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetPunctuation (BYVAL hRichEdit AS HWND, BYVAL punctp AS DWORD, BYVAL lppunct AS PUNCTUATION PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETPUNCTUATION, punctp, cast(LPARAM, lppunct))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the formatting rectangle of a rich edit control.
' ========================================================================================
PRIVATE SUB RichEdit_GetRect (BYVAL hRichEdit AS HWND, BYVAL prc AS RECT PTR)
   SendMessageW hRichEdit, EM_GETRECT, 0, cast(LPARAM, prc)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the type of the next action, if any, in the control's redo queue.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetRedoName (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETREDONAME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Obtains the current scroll position of the edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetScrollPos (BYVAL hRichEdit AS HWND, BYVAL lppt AS POINT PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETSCROLLPOS, 0, cast(LPARAM, lppt))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the starting and ending character positions of the current selection in a rich
' edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetSel (BYVAL hRichEdit AS HWND, BYVAL pdwStartPos AS DWORD PTR, BYVAL pdwEndPos AS DWORD PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETSEL, cast(WPARAM, pdwStartPos), cast(LPARAM, pdwEndPos))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the currently selected text in a rich edit control.
' Note: From version 4.1 (RichEdit50W class in MSFTEDIT.DLL), Windows XP, EM_GETSELTEXT
' returns the text in unicode.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetSelText (BYVAL hRichEdit AS HWND) AS CWSTR
   DIM dwStartPos AS DWORD, dwEndPos AS DWORD, cr AS CHARRANGE
   SendMessageW(hRichEdit, EM_EXGETSEL, 0, cast(LPARAM, @cr))
   DIM cbLen AS DWORD = ABS(cr.cpMax - cr.cpMin)
   IF cbLen < 1 THEN RETURN ""
   DIM cwsText AS CWSTR = cbLen + 1
   cbLen = SendMessageW(hRichEdit, EM_GETSELTEXT, 0, cast(LPARAM, *cwsText))
   RETURN LEFT(**cwsText, cbLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the text from a rich edit control.
' Note: GetWindowText cannot retrieve the text of a control in another application.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetText (BYVAL hRichEdit AS HWND) AS CWSTR
   DIM cbLen AS DWORD = SendMessageW(hRichEdit, WM_GETTEXTLENGTH, 0, 0)
   IF cbLen < 1 THEN RETURN ""
   DIM cwsText AS CWSTR = cbLen + 1
   cbLen = SendMessageW(hRichEdit, WM_GETTEXT, cbLen + 1, cast(LPARAM, *cwsText))
   RETURN LEFT(**cwsText, cbLen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets all of the text from the rich edit control in any particular code base you want.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetTextEx (BYVAL hRichEdit AS HWND, BYVAL lpgtex AS GETTEXTEX PTR, BYVAL buffer AS ANY PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETTEXTEX, cast(WPARAM, lpgtex), cast(LPARAM, buffer))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length of all text in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetTextLength (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, WM_GETTEXTLENGTH, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Calculates text length in various ways. It is usually called before creating a buffer to
' receive the text from the control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetTextLengthEx (BYVAL hRichEdit AS HWND, BYVAL lpgtex AS GETTEXTLENGTHEX PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETTEXTLENGTHEX, cast(WPARAM, lpgtex), 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current text mode and undo level of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetTextMode (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETTEXTMODE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves a specified range of characters from a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetTextRange (BYVAL hRichEdit AS HWND, BYVAL lptrg AS TEXTRANGE PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETTEXTRANGE, 0, cast(LPARAM, lptrg))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the position of the scroll box (thumb) in the vertical scroll bar of a multiline
' rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetThumb (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETTHUMB, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the current state of the typography options of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetTypographyOptions (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETTYPOGRAPHYOPTIONS, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the type of the next undo action, if any.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetUndoName (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETUNDONAME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the address of the current Wordwrap function.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetWordBreakProc (BYVAL hRichEdit AS HWND) AS LONG_PTR
   FUNCTION = SendMessageW(hRichEdit, EM_GETWORDBREAKPROC, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the address of the currently registered extended word-break procedure.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetWordBreakProcEx (BYVAL hRichEdit AS HWND) AS LONG_PTR
   FUNCTION = SendMessageW(hRichEdit, EM_GETWORDBREAKPROCEX, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current word wrap and word-break options for the rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetWordWrapMode (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_GETWORDWRAPMODE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the current zoom ratio, which is always between 1/64 and 64.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetZoom (BYVAL hRichEdit AS HWND, BYVAL pzNum AS DWORD PTR, BYVAL pzDen AS DWORD PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_GETZOOM, cast(WPARAM, pzNum), cast(LPARAM, pzDen))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Hides or shows the selection in a rich edit control.
' ========================================================================================
PRIVATE SUB RichEdit_HideSelection (BYVAL hRichEdit AS HWND, BYVAL fHide AS DWORD)
   SendMessageW hRichEdit, EM_HIDESELECTION, fHide, 0
END SUB
' ========================================================================================

' ========================================================================================
' Determines if current input locale is an East Asian locale.
' ========================================================================================
PRIVATE FUNCTION RichEdit_IsIME (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_ISIME, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text limit of a rich edit control. The text limit is the maximum amount of text,
' in TCHARs, that the user can type into the edit control.
' ========================================================================================
PRIVATE SUB RichEdit_LimitText (BYVAL hRichEdit AS HWND, BYVAL chMax AS DWORD)
   SendMessageW hRichEdit, EM_LIMITTEXT, chMax, 0
END SUB
' ========================================================================================

' ========================================================================================
' Gets the index of the line that contains the specified character index in a multiline
' rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_LineFromChar (BYVAL hRichEdit AS HWND, BYVAL index AS DWORD) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_LINEFROMCHAR, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Gets the character index of the first character of a specified line in a multiline
' rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_LineIndex (BYVAL hRichEdit AS HWND, BYVAL nLine AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_LINEINDEX, nLine, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the length, in characters, of a line in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_LineLength (BYVAL hRichEdit AS HWND, BYVAL index AS DWORD) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_LINELENGTH, index, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the text in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_LineScroll (BYVAL hRichEdit AS HWND, BYVAL y AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_LINESCROLL, 0, y)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Pastes a specific clipboard format in a rich edit control.
' ========================================================================================
PRIVATE SUB RichEdit_PasteSpecial (BYVAL hRichEdit AS HWND, BYVAL clpfmt AS DWORD, BYVAL lprps AS REPASTESPECIAL PTR)
   SendMessageW hRichEdit, EM_PASTESPECIAL, clpfmt, cast(LPARAM, lprps)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the client area coordinates of a specified character in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_PosFromChar (BYVAL hRichEdit AS HWND, BYVAL pt AS POINTL PTR, BYVAL index as DWORD) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_POSFROMCHAR, cast(WPARAM, pt), index)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Invokes the Input Method Editor (IME) reconversion dialog box.
' ========================================================================================
PRIVATE FUNCTION RichEdit_Reconversion (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_RECONVERSION, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Redoes the next action in the control's redo queue.
' ========================================================================================
PRIVATE FUNCTION RichEdit_Redo (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_REDO, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces the current selection in a rich edit control with the specified text.
' ========================================================================================
PRIVATE SUB RichEdit_ReplaceSel (BYVAL hRichEdit AS HWND, BYVAL bCanBeUndone AS LONG, BYVAL pwszText AS WSTRING PTR)
   SendMessageW hRichEdit, EM_REPLACESEL, bCanBeUndone, cast(LPARAM, pwszText)
END SUB
' ========================================================================================

' ========================================================================================
' Forces a rich edit control to send an EN_REQUESTRESIZE notification message to its
' parent window.
' ========================================================================================
PRIVATE SUB RichEdit_RequestResize (BYVAL hRichEdit AS HWND)
   SendMessageW hRichEdit, EM_REQUESTRESIZE, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Scrolls the text vertically in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_Scroll (BYVAL hRichEdit AS HWND, BYVAL nAction AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SCROLL, nAction, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Scrolls the caret into view in a rich edit control.
' ========================================================================================
PRIVATE SUB RichEdit_ScrollCaret (BYVAL hRichEdit AS HWND)
   SendMessageW hRichEdit, EM_SCROLLCARET, 0, 0
END SUB
' ========================================================================================

' ========================================================================================
' Determines the selection type for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SelectionType (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SELECTIONTYPE, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current state of the bidirectional options in the rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetBidiOptions (BYVAL hRichEdit AS HWND, BYVAL pOptions AS BIDIOPTIONS PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETBIDIOPTIONS, 0, cast(LPARAM, pOptions))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the background color for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetBkgndColor (BYVAL hRichEdit AS HWND, BYVAL pSysColor AS DWORD, BYVAL pBkColor AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_SETBKGNDCOLOR, pSysColor, pBkColor)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets character formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetCharFormat (BYVAL hRichEdit AS HWND, BYVAL chfmt AS DWORD, BYVAL pchfmt AS DWORD) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETCHARFORMAT, chfmt, pchfmt)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the Text Services Framework (TSF) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetCTFModeBias (BYVAL hRichEdit AS HWND, BYVAL nModeBias AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETCTFMODEBIAS, nModeBias, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Opens or closes the Text Services Framework (TSF) keyboard.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetCTFOpenStatus (BYVAL hRichEdit AS HWND, BYVAL fTSFkbd AS LONG) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_SETCTFOPENSTATUS, fTSFkbd, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the current edit style flags.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetEditStyle (BYVAL hRichEdit AS HWND, BYVAL fStyle AS LONG, BYVAL fMask AS LONG) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_SETEDITSTYLE, fStyle, fMask)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the event mask for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetEventMask (BYVAL hRichEdit AS HWND, BYVAL fMask AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETEVENTMASK, 0, fMask)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font size for the selected text.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetFontSize (BYVAL hRichEdit AS HWND, BYVAL ptsize AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETFONTSIZE, ptsize, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the way a Microsoft Rich Edit control does hyphenation.
' ========================================================================================
PRIVATE SUB RichEdit_SetHyphenateInfo (BYVAL hRichEdit AS HWND, BYVAL phinfo AS HYPHENATEINFO PTR)
   SendMessageW hRichEdit, EM_SETHYPHENATEINFO, cast(WPARAM, phinfo), 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the Input Method Editor (IME) composition color.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetIMEColor (BYVAL hRichEdit AS HWND, BYVAL pcompcolor AS COMPCOLOR PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETIMECOLOR, 0, cast(LPARAM, pcompcolor))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the Input Method Editor (IME) mode bias for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetIMEModeBias (BYVAL hRichEdit AS HWND, BYVAL nModeBias AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit,EM_SETIMEMODEBIAS, nModeBias, nModeBias)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the Input Method Editor (IME) options.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetIMEOptions (BYVAL hRichEdit AS HWND, BYVAL fCoop AS LONG, BYVAL fOptions AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETIMEOPTIONS, fCoop, fOptions)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets options for Input Method Editor (IME) and Asian language support in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetLangOptions (BYVAL hRichEdit AS HWND, BYVAL lgoptions AS LONG) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_SETLANGOPTIONS, 0, lgoptions)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text limit of a rich edit control. The text limit is the maximum amount of
' text, in TCHARs, that the user can type into the edit control.
' ========================================================================================
PRIVATE SUB RichEdit_SetLimitText (BYVAL hRichEdit AS HWND, BYVAL chMax AS DWORD)
   SendMessageW hRichEdit, EM_LIMITTEXT, chMax, 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the widths of the left and right margins for a rich edit control. The message
' redraws the control to reflect the new margins.
' ========================================================================================
PRIVATE SUB RichEdit_SetMargins (BYVAL hRichEdit AS HWND, BYVAL nMargins AS LONG, BYVAL nWidth AS LONG)
   SendMessageW(hRichEdit, EM_SETMARGINS, nMargins, nWidth)
END SUB
' ========================================================================================

' ========================================================================================
' Sets or clears the modification flag for a rich edit control. The modification flag
' indicates whether the text within the rich edit control has been modified.
' ========================================================================================
PRIVATE SUB RichEdit_SetModify (BYVAL hRichEdit AS HWND, BYVAL fModify AS LONG)
   SendMessageW hRichEdit, EM_SETMODIFY, fModify, 0
END SUB
' ========================================================================================

' ========================================================================================
' Gives a rich edit control an IRichEditOleCallback object that the control uses to get
' OLE-related resources and information from the client.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetOleCallback (BYVAL hRichEdit AS HWND, BYVAL pCallback AS ANY PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETOLECALLBACK, 0, cast(LPARAM, pCallback))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the options for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetOptions (BYVAL hRichEdit AS HWND, BYVAL fCoop AS LONG, BYVAL fOptions AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETOPTIONS, fCoop, fOptions)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deprecated. Sets the text layout for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetPageRotate (BYVAL hRichEdit AS HWND, BYVAL txtlayout AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETPAGEROTATE, txtlayout, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Deprecated. Sets the text layout for a Microsoft Rich Edit control.
' ========================================================================================
PRIVATE SUB RichEdit_SetPalette (BYVAL hRichEdit AS HWND, BYVAL newPalette AS HPALETTE)
   SendMessageW hRichEdit, EM_SETPALETTE, cast(WPARAM, newPalette), 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the paragraph formatting for the current selection in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetParaFormat (BYVAL hRichEdit AS HWND, BYVAL pfmt AS PARAFORMAT PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETPARAFORMAT, 0, cast(LPARAM, pfmt))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets or removes the password character for a rich edit control. When a password
' character is set, that character is displayed in place of the characters typed by the user.
' ========================================================================================
PRIVATE SUB RichEdit_SetPasswordChar (BYVAL hRichEdit AS HWND, BYVAL dwchar AS DWORD)
   SendMessageW hRichEdit, EM_SETPASSWORDCHAR, dwchar, 0
END SUB
' ========================================================================================

' ========================================================================================
' Sets the punctuation characters for a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetPunctuation (BYVAL hRichEdit AS HWND, BYVAL ptype AS LONG, BYVAL ppunct AS PUNCTUATION PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETPUNCTUATION, ptype, cast(LPARAM, ppunct))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets or removes the read-only style (ES_READONLY) of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetReadOnly (BYVAL hRichEdit AS HWND, BYVAL fReadOnly AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETREADONLY, fReadOnly, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the formatting rectangle of a multiline rich edit control.
' ========================================================================================
PRIVATE SUB RichEdit_SetRect (BYVAL hRichEdit AS HWND, BYVAL fCoord AS LONG, BYVAL prect AS RECT PTR)
   SendMessageW hRichEdit, EM_SETRECT, fCoord, cast(LPARAM, prect)
END SUB
' ========================================================================================

' ========================================================================================
' Sets the formatting rectangle of a multiline rich edit control.
' ========================================================================================
PRIVATE SUB RichEdit_SetRectNP (BYVAL hRichEdit AS HWND, BYVAL fCoord AS LONG, BYVAL prect AS RECT PTR)
   SendMessageW hRichEdit, EM_SETRECTNP, fCoord, cast(LPARAM, prect)
END SUB
' ========================================================================================

' ========================================================================================
' Tells the rich edit control to scroll to a particular point.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetScrollPos (BYVAL hRichEdit AS HWND, BYVAL pt AS POINT PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_SETSCROLLPOS, 0, cast(LPARAM, pt))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Selects a range of characters in a rich edit control.
' ========================================================================================
PRIVATE SUB RichEdit_SetSel (BYVAL hRichEdit AS HWND, BYVAL nStart AS LONG, BYVAL nEnd AS LONG)
   SendMessageW hRichEdit, EM_SETSEL, nStart, nEnd
END SUB
' ========================================================================================

' ========================================================================================
' Sets the tab stops in a multiline rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetTabStops (BYVAL hRichEdit AS HWND, BYVAL nTabs AS LONG, BYVAL rgTabStops AS LONG_PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETTABSTOPS, nTabs, cast(LPARAM, rgTabStops))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the target device and line width used for WYSIWYG formatting in a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetTargetDevice (BYVAL hRichEdit AS HWND, BYVAL hDC AS HDC, BYVAL lnwidth AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETTARGETDEVICE, cast(WPARAM, hDC), lnwidth)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text of an edit control.
' Note: SetWindowText cannot change the text of a control in another application.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetText (BYVAL hRichEdit AS HWND, BYVAL pwszText AS WSTRING PTR) AS LONG
   FUNCTION = SendMessageW(hRichEdit, WM_SETTEXT, 0, cast(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Combines the functionality of WM_SETTEXT and EM_REPLACESEL and adds the ability to set
' text using a code page and to use either Rich Text Format (RTF) rich text or plain text.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetTextExW (BYVAL hRichEdit AS HWND, BYVAL pstex AS SETTEXTEX PTR, BYVAL pwszText AS WSTRING PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_SETTEXTEX, cast(WPARAM, pstex), cast(LPARAM, pwszText))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text mode or undo level of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetTextMode (BYVAL hRichEdit AS HWND, BYVAL pvalues AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETTEXTMODE, 0, pvalues)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the text mode or undo level of a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetTypographyOptions (BYVAL hRichEdit AS HWND, BYVAL pto AS LONG, BYVAL fMask AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETTYPOGRAPHYOPTIONS, pto, fMask)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the maximum number of actions that can stored in the undo queue.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetUndoLimit (BYVAL hRichEdit AS HWND, BYVAL maxactions AS DWORD) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_SETUNDOLIMIT, maxactions, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces a rich edit control's default Wordwrap function with an application-defined
' Wordwrap function.
' ========================================================================================
PRIVATE SUB RichEdit_SetWordBreakProc (BYVAL hRichEdit AS HWND, BYVAL pfn AS EDITWORDBREAKPROC)
   SendMessageW hRichEdit, EM_SETWORDBREAKPROC, 0, cast(LPARAM, cast(EDITWORDBREAKPROC, pfn))
END SUB
' ========================================================================================

' ========================================================================================
' Sets the extended word-break procedure.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetWordBreakProcEx (BYVAL hRichEdit AS HWND, BYVAL pfn AS EDITWORDBREAKPROC) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_SETWORDBREAKPROCEX, 0, cast(LPARAM, cast(EDITWORDBREAKPROC, pfn)))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the word-wrapping and word-breaking options for the rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetWordWrapMode (BYVAL hRichEdit AS HWND, BYVAL pvalues AS LONG) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETWORDWRAPMODE, pvalues, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the zoom ratio anywhere between 1/64 and 64.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetZoom (BYVAL hRichEdit AS HWND, BYVAL zNum AS DWORD, BYVAL zDen AS DWORD) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_SETZOOM, zNum, zDen)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Shows or hides one of the scroll bars in the Text Host window.
' ========================================================================================
PRIVATE SUB RichEdit_ShowScrollBar (BYVAL hRichEdit AS HWND, BYVAL nScrollBar AS DWORD, BYVAL fShow AS LONG)
   SendMessageW hRichEdit, EM_SHOWSCROLLBAR, nScrollBar, fShow
END SUB
' ========================================================================================

' ========================================================================================
' Stops the control from collecting additional typing actions into the current undo action.
' ========================================================================================
PRIVATE FUNCTION RichEdit_StopGroupTyping (BYVAL hRichEdit AS HWND) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_STOPGROUPTYPING, 0, 0)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Replaces the contents of a rich edit control with a stream of data provided by an
' application defined–EditStreamCallback callback function.
' ========================================================================================
PRIVATE FUNCTION RichEdit_StreamIn (BYVAL hRichEdit AS HWND, BYVAL psf AS LONG, BYVAL pedst AS EDITSTREAM PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_STREAMIN, psf, cast(LPARAM, pedst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Causes a rich edit control to pass its contents to an application–defined
' EditStreamCallback callback function.
' ========================================================================================
PRIVATE FUNCTION RichEdit_StreamOut (BYVAL hRichEdit AS HWND, BYVAL psf AS LONG, BYVAL pedst AS EDITSTREAM PTR) AS DWORD
   FUNCTION = SendMessageW(hRichEdit, EM_STREAMOUT, psf, cast(LPARAM, pedst))
END FUNCTION
' ========================================================================================

' ========================================================================================
' This message undoes the last edit control operation in the control's undo queue.
' ========================================================================================
PRIVATE FUNCTION RichEdit_Undo (BYVAL hRichEdit AS HWND) AS LONG
   FUNCTION = SendMessageW(hRichEdit, EM_UNDO, 0, 0)
END FUNCTION
' ========================================================================================


' ########################################################################################
'                          H E L P E R    F U N C T I O N S
' ########################################################################################

' ========================================================================================
' Enumerates font families. Used by the RichEdit_SetFontW function.
' ========================================================================================
PRIVATE FUNCTION RichEdit_EnumFontFamProcW ( _
   BYVAL lpelf    AS ENUMLOGFONTW PTR, _     ' // Address of ENUMLOGFONT structure
   BYVAL lpntm    AS NEWTEXTMETRICW PTR, _   ' // Address of NEWTEXTMETRIC structure
   BYVAL FontType AS LONG, _                 ' // Font type
   BYVAL lplf     AS LOGFONTW PTR _          ' // Address of LOGFONT struct
   ) AS LONG

   lplf->lfCharSet        = lpelf->elfLogFont.lfCharSet
   lplf->lfPitchAndFamily = lpelf->elfLogFont.lfPitchAndFamily
   lplf->lfFaceName       = lpelf->elfLogFont.lfFaceName

   FUNCTION = FALSE

END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the font used by a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_SetFontW ( _
   BYVAL hRichEdit AS HWND, _             ' // Handle to the RichEdit control
   BYREF wszFaceName AS WSTRING, _        ' // Font name
   BYVAL ptsize AS LONG _                 ' // Font size in points
   ) AS LRESULT

   DIM lResult AS LRESULT                 ' // Result code
   DIM hDC AS HDC                         ' // Handle of the device context
   DIM tlf AS LOGFONTW                    ' // LOGFONT structure
   DIM tcf AS CHARFORMATW                 ' // CHARFORMAT structure

   hDC = GetDC(NULL)
   EnumFontFamiliesW(hDC, wszFaceName, cast(FONTENUMPROCW, @RichEdit_EnumFontFamProcW), cast(LPARAM, @tlf))
   ReleaseDC NULL, hDC
   tcf.cbSize = SIZEOF(tcf)
   tcf.dwMask = CFM_BOLD OR CFM_ITALIC OR CFM_UNDERLINE OR CFM_STRIKEOUT OR _
                CFM_FACE OR CFM_CHARSET OR CFM_SIZE
   tcf.yHeight = ptsize * 20   ' // Expects it in 20ths of a point
   tcf.bCharSet = tlf.lfCharSet
   tcf.bPitchAndFamily = tlf.lfPitchAndFamily
   tcf.szFaceName = tlf.lfFaceName
   lResult = SendMessageW(hRichEdit, EM_SETCHARFORMAT, SCF_ALL, cast(LPARAM, @tcf))
   ' // Specify which notifications the control sends to its parent window
   IF lResult <> 0 THEN lResult = SendMessageW(hRichEdit, EM_SETEVENTMASK, 0, ENM_CHANGE)
   FUNCTION = lResult

END FUNCTION
' ========================================================================================

' ========================================================================================
' Callback function used by the RichEdit_LoadRtfFromFileW function.
' Transfers a stream of data into a rich edit control.
' ========================================================================================
FUNCTION RichEdit_LoadRtfFromFileCallback ( _
   BYVAL hFile AS HANDLE _                  ' // Value of the dwCookie member of the EDITSTREAM structure.
 , BYVAL lpBuff AS BYTE PTR _               ' // Pointer to a buffer to write to
 , BYVAL cb AS LONG _                       ' // Maximum number of bytes to read
 , BYVAL pcb AS LONG PTR _                  ' // Number of bytes actually read
 ) AS UINT                                  ' // 0 for success, or an error code

   IF ReadFile(hFile, lpBuff, cb, pcb, NULL) = 0 THEN FUNCTION = GetLastError

END FUNCTION
' ========================================================================================

' ========================================================================================
FUNCTION RichEdit_LoadRtfFromFileW ( _
   BYVAL hRichEdit AS HWND _                ' // Handle of the Rich Edit control
 , BYREF wszFileName AS WSTRING _           ' // Name of the file to load
 ) AS BOOLEAN                               ' // TRUE or FALSE

   DIM hFile AS HANDLE                      ' // File handle
   DIM eds AS EDITSTREAM                    ' // EDITSTREAM structure

   ' // Checks the validity of the parameters
   IF hRichEdit = 0 THEN EXIT FUNCTION
   IF LEN(wszFileName) = 0 THEN EXIT FUNCTION

   ' // Opens the file and sends the message
   hFile = CreateFileW(wszFileName, GENERIC_READ, FILE_SHARE_READ, _
                       BYVAL NULL, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, NULL)
   IF hFile = INVALID_HANDLE_VALUE THEN EXIT FUNCTION
   eds.dwCookie = cast(DWORD_PTR, hFile)
   eds.pfnCallback = cast(EDITSTREAMCALLBACK, @RichEdit_LoadRtfFromFileCallback)
   IF SendMessageW(hRichEdit, EM_STREAMIN, SF_RTF, cast(LPARAM, @eds)) > 0 AND eds.dwError = 0 THEN FUNCTION = TRUE
   CloseHandle hFile

END FUNCTION
' ========================================================================================

' ========================================================================================
' Custom structure used by the RichEdit_LoadRtfFromResourceW function.
' ========================================================================================
TYPE AFX_RICHEDIT_CUSTOMDATA
   pData  AS BYTE PTR
   nLen   AS LONG
   curPos AS LONG
END TYPE
' ========================================================================================

' ========================================================================================
' Callback function used by the RichEdit_LoadRtfFromResourceW function.
' Transfers a stream of data into a rich edit control.
' ========================================================================================
PRIVATE FUNCTION RichEdit_LoadRtfFromResourceCallback ( _
   BYVAL pCustData AS AFX_RICHEDIT_CUSTOMDATA PTR _   ' // Value of the dwCookie member of the EDITSTREAM structure.
 , BYVAL lpBuff AS BYTE PTR _                         ' // Pointer to a buffer to write to.
 , BYVAL cb AS LONG _                                 ' // Number of bytes to write.
 , BYVAL pcb AS LONG PTR _                            ' // Number of bytes actually written.
 ) AS DWORD                                           ' // 0 for success, or an error code

   DIM nBytes AS LONG
   IF pCustData->nLen - pCustData->curPos > cb THEN nBytes = cb ELSE nBytes = pCustData->nLen - pCustData->curPos
   IF nBytes THEN
      CopyMemory(lpBuff, pCustData->pData + pCustData->curPos, nBytes)
      pCustData->curPos = pCustData->curPos + nBytes
      FUNCTION = 0
   ELSE
      FUNCTION = 1
   END IF
   *pcb = nBytes

END FUNCTION
' ========================================================================================

' ========================================================================================
' Loads a RTF resource file into a Rich Edit control.
' The EM_STREAMIN message replaces the contents of a rich edit control with a stream of
' data provided by an application defined–EditStreamCallback callback function.
' ========================================================================================
PRIVATE FUNCTION RichEdit_LoadRtfFromResourceW ( _
   BYVAL hRichEdit AS HWND _                ' // Handle of the Rich Edit control
 , BYVAL hInstance AS HINSTANCE _           ' // Instance handle
 , BYREF wszResourceName AS WSTRING _       ' // Name of the resource to load
 ) AS BOOLEAN                               ' // TRUE or FALSE

   DIM hResInfo AS HRSRC                        ' // Resource handle
   DIM pResData AS LPVOID                       ' // Pointer to the resource data
   DIM eds AS EDITSTREAM                        ' // EDITSTREAM structure
   DIM rtfCustData AS AFX_RICHEDIT_CUSTOMDATA   ' // AFX_RICHEDIT_CUSTOMDATA structure

   ' // Checks the validity of the parameters
   IF hRichEdit = NULL OR hInstance = NULL THEN EXIT FUNCTION
   IF LEN(wszResourceName) = 0 THEN EXIT FUNCTION

   ' // Loads the resource
   hResInfo = FindResourceW(hInstance, wszResourceName, RT_RCDATA)
   IF hResInfo = NULL THEN EXIT FUNCTION

   ' // Loads and locks the resource
   ' // Note  LockResource does not actually lock memory; it is just used to obtain
   ' // a pointer to the memory containing the resource data.
   pResData = LockResource(LoadResource(hInstance, hResInfo))
   IF pResData = NULL THEN EXIT FUNCTION
   DIM cbSize AS LONG = SizeofResource(hInstance, hResInfo)
   DIM buffer AS STRING = SPACE(cbSize)
   CopyMemory(STRPTR(buffer), pResData, cbSize)

   ' // Sends the message
   rtfCustData.pData = STRPTR(buffer)
   rtfCustData.nLen = cbSize
   rtfCustData.curPos = 0
   eds.dwCookie = cast(DWORD_PTR, @rtfCustData)
   eds.pfnCallback = cast(EDITSTREAMCALLBACK, @RichEdit_LoadRtfFromResourceCallback)
   IF SendMessageW(hRichEdit, EM_STREAMIN, SF_RTF, cast(LPARAM, @eds)) > 0 AND eds.dwError = 0 THEN
      FUNCTION = TRUE
   END IF

END FUNCTION
' ========================================================================================

' ========================================================================================
' Callback used by the RichEdit_GetRtfText function.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetTextCallback ( _
   BYVAL dwCookie AS DWORD_PTR _                      ' // Value of the dwCookie member of the EDITSTREAM structure.
 , BYVAL pbBuff AS BYTE PTR _                         ' // Pointer to the buffer to read from.
 , BYVAL cb AS LONG _                                 ' // Number of bytes to read.
 , BYVAL pcb AS LONG PTR _                            ' // Number of bytes actually read.
 ) AS DWORD                                           ' // 0 for success, or an error code

   DIM pcws AS CWSTR PTR = cast(CWSTR PTR, dwCookie)
   pcws->AppendBuffer(pbBuff, cb)
   *pcb = cb
   FUNCTION = 0

END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves RTF formatted text from a Rich Edit control
' - hRichEdit = Handle of the Rich Edit control.
' Returns the retrieved text or a null string.
' ========================================================================================
PRIVATE FUNCTION RichEdit_GetRtfText (BYVAL hRichEdit AS HWND) AS STRING

   DIM eds AS EDITSTREAM, cws AS CWSTR
   eds.dwCookie = cast(DWORD_PTR, @cws)
   eds.pfnCallBack = cast(EDITSTREAMCALLBACK, @RichEdit_GetTextCallback)
   SendMessageW hRichEdit, EM_STREAMOUT, SF_RTF, cast(LPARAM, @eds)
   ' // Copy the ansi contents of the CWSTR to a STRING
   DIM s AS STRING = SPACE(cws.m_BufferLen - 2)   ' // -2 to remove the ending nulls
   IF LEN(s) THEN CopyMemory(STRPTR(s), cws.m_pBuffer, cws.m_BufferLen - 2)
   RETURN s

END FUNCTION
' ========================================================================================
