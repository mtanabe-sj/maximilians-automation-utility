/*
  Copyright (c) 2022 Makoto Tanabe <mtanabe.sj@outlook.com>
  Licensed under the MIT License.

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.
*/
#include "stdafx.h"
#include "InputBoxImpl.h"


/* put_Caption - [propput] replaces a default InputBox dialog caption with a custom caption.

Parameters:
NewValue - [in] a custom caption to use.

Remarks:
Set the Caption to a custome value before calling the Show method.
*/
STDMETHODIMP InputBoxImpl::put_Caption(/* [in] */ BSTR NewValue)
{
	_caption = NewValue;
	return S_OK;
}

/* get_Caption - [propget] retrieves a current caption set by a prior call to put_Caption. A NULL is returned if no custom caption exists.

Parameters:
Value - [retval][out] contains a custom caption or NULL.
*/
STDMETHODIMP InputBoxImpl::get_Caption(/* [retval][out] */ BSTR *Value)
{
	*Value = _caption.clone();
	return S_OK;
}

/* put_Note - [propput] assigns custom text as a note to be displayed under the input field.

Parameters:
NewValue - [in] text as a note.

Remarks:
Set the Note to a custome value before calling the Show method.
*/
STDMETHODIMP InputBoxImpl::put_Note(/* [in] */ BSTR NewValue)
{
	_note = NewValue;
	return S_OK;
}

/* get_Note - [propget] retrieves the text of a current note set in a prior call to put_Note.

Parameters:
Value - [retval][out] contains a current note if one is set, or is set to NULL if no note exists.
*/
STDMETHODIMP InputBoxImpl::get_Note(/* [retval][out] */ BSTR *Value)
{
	*Value = _note.clone();
	return S_OK;
}

/* put_CheckBox - [propput] assigns text as a label to the checkbox.

Parameters:
NewValue - [in] text as a note.

Remarks:
The checkbox is hidden if no text is assigned to it.
*/
STDMETHODIMP InputBoxImpl::put_CheckBox(/* [in] */ BSTR NewValue)
{
	_checkbox = NewValue;
	return S_OK;
}

/* get_CheckBox - [propget] retrieves text currently assigned to the checkbox.

Parameters:
Value - [retval][out] contains a current label assigned to the checkbox.
*/
STDMETHODIMP InputBoxImpl::get_CheckBox(/* [retval][out] */ BSTR *Value)
{
	*Value = _checkbox.clone();
	return S_OK;
}

/* put_Options - [propput] enables optional settings of InputBox. available options are as follows:
option 1) IB_MULTILINE (0x4) - stretches the input field to a multi-line box so that multiple lines of text can be entered.
option 2) IB_PASSWORD (0x20) - hides characters typed into the input field by changing them to a series of '*' in display.
option 3) IB_NUMBER (0x2000) - permits entry of numbers (0..9) only. note that a user can paste in non-numeric characters in spite of the option.

Parameters:
NewValue - [in] set to one of the above option specifying IB_* strings as VT_BSTR. or, it can be set to an ES_* style bit of Win32 as a VT_I4 or VT_I2, which corresponds to one of the IB_* options.

Remarks:
Set the Options to a custome value before calling the Show method.

The IB_* options supported by InputBox map to these ES_* style bits. The styles are used to control the appearance and behavior of the InputValue edit control.

'IB_MULTILINE' ==> ES_MULTILINE (ES_AUTOHSCROLL, ES_AUTOVSCROLL, and ES_WANTRETURN will be added automatically)
'IB_PASSWORD' ==> ES_PASSWORD
'IB_NUMBER' ==> ES_NUMBER

You could pass a Win32 ES_* style bit corresponding to an IB_* option as an integer instead of a BSTR. we don't make any validation. So make sure the value being passed is correct.
*/
STDMETHODIMP InputBoxImpl::put_Options(/* [in] */ VARIANT NewValue)
{
	if (NewValue.vt == VT_BSTR && NewValue.bstrVal)
	{
		if (0 == _wcsicmp(NewValue.bstrVal, L"IB_MULTILINE"))
		{
			_options |= ES_MULTILINE;
			return S_OK;
		}
		if (0 == _wcsicmp(NewValue.bstrVal, L"IB_PASSWORD"))
		{
			_options |= ES_PASSWORD;
			return S_OK;
		}
		if (0 == _wcsicmp(NewValue.bstrVal, L"IB_NUMBER"))
		{
			_options |= ES_NUMBER;
			return S_OK;
		}
	}
	if (NewValue.vt == VT_I2 || NewValue.vt == VT_UI2)
	{
		_options = MAKELONG(NewValue.uiVal,0);
		return S_OK;
	}
	if (NewValue.vt == VT_I4 || NewValue.vt == VT_UI4)
	{
		_options = NewValue.lVal;
		return S_OK;
	}
	return E_INVALIDARG;
}

/* get_Options - [propget] retrieves currently enabled options. the returned value comes in an integer (VT_UI4). use this mapping to identify which option the bit stands for.
0x4 (ES_MULTILINE) --> 'IB_MULTILINE'
0x20 (ES_PASSWORD) --> 'IB_PASSWORD'
0x2000 (ES_NUMBER) --> 'IB_NUMBER'

Parameters:
Value - [retval][out] contains a current optional setting assigned by a previous call to put_Options. 0 if no assignment has been made. the returned value is always in VT_UI4.

Remarks:
Reading the Options property only makes sense if you reuse a single InputBox instance.
*/
STDMETHODIMP InputBoxImpl::get_Options(/* [retval][out] */ VARIANT *Value)
{
	Value->vt = VT_UI4;
	Value->lVal = _options;
	return S_OK;
}

/* put_InitialValue - [propput] assigns custom text as an initial value to be set in the input field. If a user okays it in the dialog, the custom text is output as the value of the InputValue property.

Parameters:
NewValue - [in] initial text to show in the input field.

Remarks:
Set the InitialValue to a custome value before calling the Show method.
*/
STDMETHODIMP InputBoxImpl::put_InitialValue(/* [in] */ BSTR NewValue)
{
	_initVal = NewValue;
	return S_OK;
}

/* get_InitialValue - [propget] retrieves a current initial value, or a NULL if no initial value exists.

Parameters:
Value - [retval][out] contains the text of an initial value set in a prior call to put_InitialValue.
*/
STDMETHODIMP InputBoxImpl::get_InitialValue(/* [retval][out] */ BSTR *Value)
{
	*Value = _initVal.clone();
	return S_OK;
}

/* get_InputValue - [propget] retrieves the text that a user has entered in the input field.

Parameters:
Value - [retval][out] contains a BSTR pointer to a character string a user has entered.

Remarks:
Read the InputValue property after calling the Show method.
*/
STDMETHODIMP InputBoxImpl::get_InputValue(/* [retval][out] */ BSTR *Value)
{
	*Value = _inputVal.clone();
	return S_OK;
}

/* put_Checked - [propput] assigns a new state to the checkbox. Use this to place a check mark as soon as the dialog runs.

Parameters:
NewValue - [in] TRUE to put a check mark, or FALSE to remove it.
*/
STDMETHODIMP InputBoxImpl::put_Checked(/* [in] */ VARIANT_BOOL NewValue)
{
	_isChecked = NewValue;
	return S_OK;
}

/* get_Checked - [propget] retrieves the state of the checkbox.

Parameters:
Value - [retval][out] TRUE if the checkbox is check, or FALSE otherwise.
*/
STDMETHODIMP InputBoxImpl::get_Checked(/* [retval][out] */ VARIANT_BOOL *Value)
{
	*Value = _isChecked;
	return S_OK;
}

/* put_BrowserButtonType - [propput] assigns a new browser dialog type to the browser button.

Parameters:
NewValue - [in] a new value in IB_BROWSERBUTTONTYPE.
*/
STDMETHODIMP InputBoxImpl::put_BrowserButtonType(/* [in] */ IB_BROWSERBUTTONTYPE NewValue)
{
	_buttonType = NewValue;
	return S_OK;
}

/* put_BrowserButtonType - [propget] retrieves the current button type.

Parameters:
Value - [retval][out] an IB_BROWSERBUTTONTYPE set in a prior call to put_BrowserButtonType.
*/
STDMETHODIMP InputBoxImpl::get_BrowserButtonType(/* [retval][out] */ IB_BROWSERBUTTONTYPE *Value)
{
	*Value = _buttonType;
	return S_OK;
}

/* put_FileFilter - [propput] assigns a file filter.

Parameters:
NewValue - [in] a new file filter string.

Remarks:
A filter string is formed according to the format,
	<File type description> | <One or more extensions> [| <2nd description> | <2nd extensions> [| <3rd...> ...] ] ]
The example below combines three filters. The first one lists image files of .jpeg, etc. The second audio files. The third all files.
"Image files|*.jpeg;*.jpg;*.png;*.bmp|Audio files|*.mp3;*.wav;All files|*.*"
*/
STDMETHODIMP InputBoxImpl::put_FileFilter(/* [in] */ BSTR NewValue)
{
	_fileFilter = NewValue;
	return S_OK;
}

/* get_FileFilter - [propget] retrieves the current file filter.

Parameters:
Value - [retval][out] a current file filter set in a prior call to put_FileFilter.
*/
STDMETHODIMP InputBoxImpl::get_FileFilter(/* [retval][out] */ BSTR *Value)
{
	*Value = _fileFilter.clone();
	return S_OK;
}

/* Show - [method] prepares and runs the dialog of IDD_INPUTBOX and stores text the user enters in the input field edit control (IDC_EDIT_INPUTVALUE). The caller should follow up by reading the InputValue property to retrieve the user input text.

Parameters:
PromptMessage - [in] a prompting message to indicate to end users what should be entered in the input field.
ParentWindow - [optional][in] a handle to a parent window or NULL. use VT_I4 or VT_INT to pass the handle value.
Result - [retval][out] contains the number of characters a user has entered, or returns -1 if the dialog is canceled.

Return value:
S_OK - the dialog closes normally. cancelation is considered a normal operation, and so, generates an S_OK, too.
*/
STDMETHODIMP InputBoxImpl::Show(/* [in] */ BSTR PromptMessage, /* [optional][in] */  VARIANT *ParentWindow, /* [retval][out] */ long *Result)
{
	_prompt = PromptMessage;
	if (ParentWindow && (ParentWindow->vt == VT_I4 || ParentWindow->vt == VT_INT))
		_hwndParent = (HWND)(ULONG_PTR)ParentWindow->ulVal;
	if (IDOK == DoModal())
		*Result = _inputVal.length();
	else
		*Result = -1; // canceled
	return S_OK;
}

/* OnInitDialog - initializes the InputBox dialog by reconfiguring the input field by applying any enabled options, and by displaying the custom caption, message, and note.
*/
BOOL InputBoxImpl::OnInitDialog()
{
	SimpleDlg::OnInitDialog();

	_hctl = GetDlgItem(_hdlg, IDC_EDIT_INPUTVALUE);
	if (_options)
	{
		// to change the style, we need to recreate the edit control. calling SetWindowLong for the changed style is not enough.
		RECT rc;
		GetWindowRect(_hctl, &rc);
		ScreenToClient(_hdlg, (LPPOINT)&rc.left);
		ScreenToClient(_hdlg, (LPPOINT)&rc.right);
		HFONT hfont = (HFONT)SendMessage(_hctl, WM_GETFONT, 0, 0);
#ifdef _DEBUG
		DWORD style0 = GetWindowLong(_hctl, GWL_STYLE);
#endif//_DEBUG
		DWORD exstyle = GetWindowLong(_hctl, GWL_EXSTYLE);
		DWORD style = _options & (ES_PASSWORD | ES_MULTILINE | ES_NUMBER);
		if (style & ES_MULTILINE)
		{
			// multiline style increases the edit control's height.
			rc.bottom += 3 * 14;
			style |= ES_AUTOHSCROLL | ES_AUTOVSCROLL | ES_WANTRETURN;
		}
		style |= WS_BORDER | WS_VISIBLE | WS_CHILD | WS_TABSTOP;
		DestroyWindow(_hctl);
		_hctl = CreateWindowEx(exstyle, L"EDIT", NULL, style, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, _hdlg, (HMENU)IDC_EDIT_INPUTVALUE, LibInstanceHandle, NULL);
		SendMessage(_hctl, WM_SETFONT, (WPARAM)hfont, 1);
		if (style & ES_MULTILINE)
		{
			// lower the Note area so it stays clear of the stretched edit control.
			HWND hnote = GetDlgItem(_hdlg, IDC_STATIC_NOTE);
			GetWindowRect(hnote, &rc);
			ScreenToClient(_hdlg, (LPPOINT)&rc.left);
			ScreenToClient(_hdlg, (LPPOINT)&rc.right);
			rc.top += 30;
			SetWindowPos(hnote, NULL, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, SWP_NOZORDER);
		}
	}

	// show custom caption, message, and note if they are defined.
	if (_caption.length() > 0)
		SetWindowText(_hdlg, _caption);
	if (_prompt.length() > 0)
		SetDlgItemText(_hdlg, IDC_STATIC_MESSAGE, _prompt);
	if (_note.length() > 0)
		SetDlgItemText(_hdlg, IDC_STATIC_NOTE, _note);
	// show the initial value in the input field if it's defined.
	if (_initVal.length() > 0)
		SetDlgItemText(_hdlg, IDC_EDIT_INPUTVALUE, _initVal);
	// unhide the checkbox if a label is defined for it.
	if (_checkbox.length() > 0)
	{
		SetDlgItemText(_hdlg, IDC_CHECK1, _checkbox);
		if (_isChecked)
			CheckDlgButton(_hdlg, IDC_CHECK1, _isChecked ? BST_CHECKED : BST_CHECKED);
		ShowWindow(GetDlgItem(_hdlg, IDC_CHECK1), SW_SHOW);
	}

	if (_buttonType != IB_BROWSERBUTTONTYPE_NONE)
		ShowWindow(GetDlgItem(_hdlg, IDC_BUTTON1), SW_SHOW);

	return TRUE;
}

/* OnOK - responds to a user clicking the OK button by retrieving the input text from the edit control and storing it in _inputVal. Also, stored the state of the checkbox. The InputBox client can read the InputValue property to retrieve the user text after the Show method returns control.
*/
BOOL InputBoxImpl::OnOK()
{
	int len = GetWindowTextLength(_hctl);
	GetWindowText(_hctl, _inputVal.alloc(len), len + 1);
	_isChecked = IsDlgButtonChecked(_hdlg, IDC_CHECK1) ? VARIANT_TRUE : VARIANT_FALSE;
	return SimpleDlg::OnOK();
}

// converts an InputBox.FileFilter string of |-delimited file filters to string of 0-delimited filters complient with requirements of the Win32 OPENFILENAME.lpstrFilter member variable.
LPCWSTR _toOfnFilter(LPWSTR outFilter, int outSize, LPCWSTR inputboxFilter)
{
	// assume that the input string has one or more filters already correctly formatted and ordered. so, just copy it to the output buffer. will replace all separator '|'s with 0's.
	wcscpy_s(outFilter, outSize, inputboxFilter);
	int cc = (int)wcslen(outFilter);
	// zero out the trailing buffer space that's unused. this is important because a filter string must end in a double null char sequence.
	ZeroMemory(outFilter + cc, (outSize-cc)*sizeof(WCHAR));
	// convert all '|' occurences with zeros.
	for (int i = 0; i < cc; i++)
	{
		if (outFilter[i] == '|')
			outFilter[i] = 0;
	}
	return outFilter;
}

// finds the very first occurence of a file extension in the input filter string. copy it to the output buffer.
LPCWSTR _findOfnDefExt(LPWSTR outFilter, int outSize, LPCWSTR inputboxFilter)
{
	// the input string contains one or more filter definitions in this order.
	// <1st description>|<1st extension[;2nd extension;...] [| <2nd description>|<1st extension[...]] >
	// we want to extract the 1st extension which follows the very first | occurence.
	LPCWSTR p1 = wcschr(inputboxFilter, '|');
	if (p1++)
	{
		/* p1, p2 and p3 points at these separators.

		<1st description>|<1st extension[;2nd extension;...] [| <2nd description> ...
		                 ^               ^                    ^
						 |               |                    |
						 + p1            + p2                 + p3
						 <----- cc ------>
		*/
		LPCWSTR p3 = wcschr(p1, '|');
		if (p3)
		{
			p1 = wcschr(p1, '.');
			if (p1)
			{
				LPCWSTR p2 = wcschr(p1, ';');
				if (!p2)
					p2 = p3;
				int cc = (int)(p2 - p1);
				if (cc)
				{
					// found the 1st extension. copy it to the output buffer.
					wcsncpy_s(outFilter, outSize, p1, cc);
					ZeroMemory(outFilter + cc, (outSize - cc) * sizeof(WCHAR));
					return outFilter;
				}
			}
		}
	}
	// an extension could not be found. is the input string malformed or corrupt?
	return NULL;
}

// runs a common file browser dialog in open mode.
HRESULT InputBoxImpl::getOpenFileName()
{
	WCHAR filter[256];
	WCHAR buf[_MAX_PATH] = { 0 };
	OPENFILENAME ofn = { sizeof(OPENFILENAME) };
	ofn.hwndOwner = _hdlg;
	ofn.Flags |= OFN_FILEMUSTEXIST;
	ofn.nFilterIndex = 1;
	ofn.lpstrInitialDir = _initVal;
	ofn.lpstrFile = buf;
	ofn.nMaxFile = ARRAYSIZE(buf);
	if (_fileFilter.length() > 0)
		ofn.lpstrFilter = _toOfnFilter(filter, ARRAYSIZE(filter), _fileFilter);
	int res = GetOpenFileName(&ofn);
	if (res == IDOK)
		SetWindowText(_hctl, buf);
	DBGPRINTF((L"GetOpenFileName: RES=%d; PATH='%s'\n", res, buf));
	return res==IDOK? S_OK : E_ABORT;
}

// runs a common file browser dialog in save mode.
HRESULT InputBoxImpl::getSaveAsFileName()
{
	WCHAR defExt[80];
	WCHAR filter[256];
	WCHAR buf[_MAX_PATH] = { 0 };
	OPENFILENAME ofn = { sizeof(OPENFILENAME) };
	ofn.hwndOwner = _hdlg;
	ofn.Flags |= OFN_PATHMUSTEXIST;
	ofn.nFilterIndex = 1;
	ofn.lpstrInitialDir = _initVal;
	ofn.lpstrFile = buf;
	ofn.nMaxFile = ARRAYSIZE(buf);
	if (_fileFilter.length() > 0)
	{
		ofn.lpstrFilter = _toOfnFilter(filter, ARRAYSIZE(filter), _fileFilter);
		ofn.lpstrDefExt = _findOfnDefExt(defExt, ARRAYSIZE(defExt), _fileFilter);
	}
	int res = GetSaveFileName(&ofn);
	if (res == IDOK)
		SetWindowText(_hctl, buf);
	DBGPRINTF((L"GetSaveFileName: RES=%d; PATH='%s'\n", res, buf));
	return res == IDOK ? S_OK : E_ABORT;
}

// runs a folder browser dialog. uses a callback to go straight to it if the client app has specified an initial directory.
HRESULT InputBoxImpl::browseForFolder() {
	LPITEMIDLIST pidlRoot = NULL;
	LPITEMIDLIST pidlSelected = NULL;
	if (S_OK != SHGetFolderLocation(_hdlg, CSIDL_DRIVES, NULL, NULL, &pidlRoot))
		return E_UNEXPECTED;
	bstringv msg(LibResourceHandle, IDS_BROWSEFORFOLDER_MESSAGE);
	WCHAR buf[_MAX_PATH] = { 0 };
	wcscpy_s(buf, ARRAYSIZE(buf)-1, _initVal);
	BROWSEINFO bi = { 0 };
	bi.hwndOwner = _hdlg;
	bi.pidlRoot = pidlRoot;
	bi.pszDisplayName = buf;
	bi.lpszTitle = msg;
	bi.ulFlags = BIF_RETURNONLYFSDIRS;
	if (_initVal.length() > 0)
	{
		bi.lpfn = (BFFCALLBACK)bffCallback;
		bi.lParam = (LPARAM)this;
	}
	pidlSelected = SHBrowseForFolder(&bi);
	if (pidlSelected)
	{
		DBGPRINTF((L"SHBrowseForFolder: %s\n", buf));
		if (!SHGetPathFromIDList(pidlSelected, buf))
			*buf = 0;
	}
	if (*buf == 0)
		return E_ABORT; // canceled
	SetWindowText(_hctl, buf);
	CoTaskMemFree(pidlRoot);
	CoTaskMemFree(pidlSelected);
	return S_OK;
}

// this is a callback for SHBrowseForFolder. it's used for selecting a starting folder on browser initialization.
int CALLBACK InputBoxImpl::bffCallback(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData)
{
	if (uMsg == BFFM_INITIALIZED)
	{
		InputBoxImpl *pThis = (InputBoxImpl*)lpData;
		SendMessage(hwnd, BFFM_SETSELECTION, 1, (LPARAM)(LPCWSTR)pThis->_initVal);
		return 1;
	}
	return 0;
}

