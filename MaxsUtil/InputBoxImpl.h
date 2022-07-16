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
#pragma once
#include "IDispatchImpl.h"
#include "MaxsUtil_h.h"
#include "resource.h"
#include "SimpleDlg.h"



/* implements the IInputBox interface of the InputBox coclass. it delegates win32 message routing to the SimplDlg base class. it overrides the dialog initialization and the ok button click. */
class InputBoxImpl :
	public SimpleDlg,
	public IDispatchWithObjectSafetyImpl<IInputBox, &IID_IInputBox, &LIBID_MaxsUtilLib>
{
public:
	InputBoxImpl() : SimpleDlg(IDD_INPUTBOX), _options(0), _hctl(NULL), _isChecked(VARIANT_FALSE), _buttonType(IB_BROWSERBUTTONTYPE_NONE) {}
	~InputBoxImpl() {}

	// IUnknown methods
	DELEGATE_IUNKNOWN_TO_IDISPATCHWITHOBJECTSAFETYIMPL(IInputBox, &IID_IInputBox, &LIBID_MaxsUtilLib)

	// IInputBox methods
	STDMETHOD(put_Caption)(/* [in] */ BSTR NewValue);
	STDMETHOD(get_Caption)(/* [retval][out] */ BSTR *Value);
	STDMETHOD(put_Note)(/* [in] */ BSTR NewValue);
	STDMETHOD(get_Note)(/* [retval][out] */ BSTR *Value);
	STDMETHOD(put_CheckBox)(/* [in] */ BSTR NewValue);
	STDMETHOD(get_CheckBox)(/* [retval][out] */ BSTR *Value);
	STDMETHOD(put_Options)(/* [in] */ VARIANT NewValue);
	STDMETHOD(get_Options)(/* [retval][out] */ VARIANT *Value);
	STDMETHOD(put_InitialValue)(/* [in] */ BSTR NewValue);
	STDMETHOD(get_InitialValue)(/* [retval][out] */ BSTR *Value);
	STDMETHOD(get_InputValue)(/* [retval][out] */ BSTR *Value);
	STDMETHOD(put_Checked)(/* [in] */ VARIANT_BOOL NewValue);
	STDMETHOD(get_Checked)(/* [retval][out] */ VARIANT_BOOL *Value);
	STDMETHOD(put_BrowserButtonType)(/* [in] */ IB_BROWSERBUTTONTYPE NewValue);
	STDMETHOD(get_BrowserButtonType)(/* [retval][out] */ IB_BROWSERBUTTONTYPE *Value);
	STDMETHOD(put_FileFilter)(/* [in] */ BSTR NewValue);
	STDMETHOD(get_FileFilter)(/* [retval][out] */ BSTR *Value);
	STDMETHOD(Show)(/* [in] */ BSTR PromptMessage, /* [optional][in] */ VARIANT *ParentWindow, /* [retval][out] */ long *Result);

protected:
	bstring _caption; // keeps a custom caption set by put_Caption. OnInitDialog replaces the built-in caption with this.
	bstring _prompt; // keeps the prompting message of Show so OnInitDialog can see it and display it in the dialog box.
	bstring _note; // text from put_Note. it's displayed under the input field as an annotation note.
	bstring _checkbox; // text from put_CheckBox. IDC_CHECK1 is made visible if this is not empty.
	bstring _initVal; // keeps a default text set by put_InitialValue before method Show is called.
	bstring _inputVal; // contains text a user has entered into the input field edit control. it's set by our OnOK override.
	long _options; // edit control styling options, e.g., ES_MULTILINE.
	HWND _hctl; // window handle of the edit control.
	VARIANT_BOOL _isChecked; // indicates that the checkbox is checked or not.
	IB_BROWSERBUTTONTYPE _buttonType; // indicates which browser dialog should be run when the browser button of IDC_BUTTON1 is clicked.
	bstring _fileFilter; // a string of file filter(s) for restricting the list to group(s) of files in the open or save-as file browser dialog, started by the browser button.

	// SimpleDlg overrides
	virtual BOOL OnInitDialog();
	virtual BOOL OnOK();
	virtual BOOL OnCommand(WPARAM wp_, LPARAM lp_)
	{
		if (LOWORD(wp_) == IDC_BUTTON1)
		{
			if (_buttonType == IB_BROWSERBUTTONTYPE_FILE_OPEN)
				getOpenFileName();
			else if (_buttonType == IB_BROWSERBUTTONTYPE_FILE_SAVEAS)
				getSaveAsFileName();
			else if (_buttonType == IB_BROWSERBUTTONTYPE_FOLDER)
				browseForFolder();
			return TRUE;
		}
		return SimpleDlg::OnCommand(wp_, lp_);
	}

	// responders to the browser button click.
	HRESULT getOpenFileName();
	HRESULT getSaveAsFileName();
	HRESULT browseForFolder();

	static int CALLBACK bffCallback(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData);
};

