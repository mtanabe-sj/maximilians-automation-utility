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
#include "ProgressBoxImpl.h"


#define PROGRESSBOX_SUPPORTS_MARQUEE
#ifdef PROGRESSBOX_SUPPORTS_MARQUEE
// this embeds a manifest for using v6 common controls of win32.
// it is needed for a PBS_MARQUEE used in the progress bar.
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif//#ifdef PROGRESSBOX_SUPPORTS_MARQUEE


ProgressBoxImpl::ProgressBoxImpl() : _cplist(this, this)
{
	// create an instance of our modeless dialog.
	_dlg = new ProgressBoxDlg(this);
}

ProgressBoxImpl::~ProgressBoxImpl()
{
	// delete the dialog instance.
	auto dlg = (ProgressBoxDlg*)InterlockedExchangePointer((LPVOID*)&_dlg, NULL);
	delete dlg;
}

#ifdef PROGRESSBOX_SUPPORTS_EVENT
VARIANT_BOOL ProgressBoxImpl::fireCancel()
{
	DBGPRINTF((L"ProgressBoxImpl::fireCancel: %p; %d\n", this, _cplist.size()));
	// initialize a local boolean to true to indicate a 'canceled' state. We will pass it to all listening clients. Any client flipping the cancelation state stops the notification.
	VARIANT_BOOL bCanceled = VARIANT_TRUE;
	for (long i = 0; i < _cplist.size(); i++)
	{
		ConnectionPointImpl* pCP = _cplist[i];
		pCP->AddRef();
		VARIANT args[1];
		args[0].vt = VT_BOOL | VT_BYREF;
		args[0].pboolVal = &bCanceled;
		pCP->FireEvent(PROGRESSBOXEVENTS_DISPID_CANCEL, args, ARRAYSIZE(args));
		pCP->Release();
		if (!bCanceled)
			break; // someone does not want us to quit.
	}
	return bCanceled;
}

// called when a client subscribes to our notification service. IConnectionPoint::Advise calls us.
// Currently, no processing is performed. Just log the call in debug build.
void ProgressBoxImpl::OnAdviseConnectionPoint()
{
	DBGPRINTF((L"ProgressBoxImpl::OnAdviseConnectionPoint: %p\n", this));
}

// called by IConnectionPoint::Unadvise when a client unsubscribes from our notification. Just write to the log for now.
void ProgressBoxImpl::OnUnadviseConnectionPoint()
{
	DBGPRINTF((L"ProgressBoxImpl::OnUnadviseConnectionPoint: %p\n", this));
}
#endif//#ifdef PROGRESSBOX_SUPPORTS_EVENT

//////////////////////////////////////////////////////////////

/* caches option flags and associated parameters from ProgressBoxImpl. If a dialog is already running, the selected options applied immediately. The method also clears a variable that holds the current state of the cancel button. This means that everytime a client invokes IProgressBox.Start which calls this method, the cancel button is reactivated unless PROGRESSBOXSTARTOPTION_DISABLE_CANCEL is selected.
*/
void ProgressBoxDlg::setOptions(long options, VARIANT *params)
{
#ifdef PROGRESSBOX_SUPPORTS_MARQUEE
	if (_hdlg)
	{
		// already interactive. may need to switch to marquee or progressive mode depending on current mode.
		if (options & PROGRESSBOXSTARTOPTION_MARQUEE)
			PostMessage(_hdlg, WM_PBD_PROGRESS_SETMARQUEE, TRUE, parseOptionalIntArg(params));
		else if (_PI.options & PROGRESSBOXSTARTOPTION_MARQUEE)
			PostMessage(_hdlg, WM_PBD_PROGRESS_SETMARQUEE, FALSE, 0);
	}
	else if (options & PROGRESSBOXSTARTOPTION_MARQUEE)
	{
		// marque update time in milliseconds.
		// 0=default marquee update time (30 ms), non-zero=custom marquee update time.
		_PI.marquee = parseOptionalIntArg(params);
	}
#endif//#ifdef PROGRESSBOX_SUPPORTS_MARQUEE

	_PI.options = options;
	// if client has invoked ProgressBox.Start again which calls us, reset _PI.canceled and update the controls based on the selected options.
	if (_hdlg)
	{
		EnableWindow(GetDlgItem(_hdlg, IDCANCEL), (_PI.options & PROGRESSBOXSTARTOPTION_DISABLE_CANCEL)?FALSE:TRUE);
		ShowWindow(_hProgess, (_PI.options & PROGRESSBOXSTARTOPTION_SHOW_PROGRESSBAR)? SW_SHOW:SW_HIDE);
	}
	_PI.canceled = VARIANT_FALSE;
}

/* responds to WM_INITDIALOG. Initializes the controls of the dialog with values assigned by ProgressBoxImpl.
*/
BOOL ProgressBoxDlg::OnInitDialog()
{
	DBGPUTS((L"ProgressBoxDlg::OnInitDialog\n"));
	SimpleModelessDlg::OnInitDialog();

#ifdef PROGRESSBOX_SUPPORTS_MARQUEE
	INITCOMMONCONTROLSEX icx = {sizeof(INITCOMMONCONTROLSEX), ICC_PROGRESS_CLASS };
	InitCommonControlsEx(&icx);
#endif//#ifdef PROGRESSBOX_SUPPORTS_MARQUEE

	_hProgess = GetDlgItem(_hdlg, IDC_PROGRESSBAR);

	// set a range for the progress bar.
	if (_PI.boundHigh > _PI.boundLow)
	{
		::SendMessage(_hProgess, PBM_SETRANGE32, _PI.boundLow, _PI.boundHigh);
		_PI.options |= PROGRESSBOXSTARTOPTION_SHOW_PROGRESSBAR;
	}
	// reset the progress position.
	if (_PI.pos)
		::SendMessage(_hProgess, PBM_SETPOS, (WPARAM)_PI.pos, 0);

#ifdef PROGRESSBOX_SUPPORTS_MARQUEE
	// put it in marquee mode if that's requested.
	if (_PI.options & PROGRESSBOXSTARTOPTION_MARQUEE)
	{
		_setMarquee();
		_PI.options |= PROGRESSBOXSTARTOPTION_SHOW_PROGRESSBAR;
	}
#endif//#ifdef PROGRESSBOX_SUPPORTS_MARQUEE

	// use a custom color for the progress bar if requested.
	if (_PI.barColor != CLR_DEFAULT)
		::SendMessage(_hProgess, PBM_SETBARCOLOR, 0, _PI.barColor);

	// show the progress bar if the show option is enabled.
	if (_PI.options & PROGRESSBOXSTARTOPTION_SHOW_PROGRESSBAR)
		ShowWindow(_hProgess, SW_SHOW);

	// hide the cancel button if the disable option is enabled.
	if (_PI.options & PROGRESSBOXSTARTOPTION_DISABLE_CANCEL)
		EnableWindow(GetDlgItem(_hdlg, IDCANCEL), FALSE);

	// show the caption, message and note assigned by ProgressBarImpl.
	if (_caption.length() > 0)
		SetWindowText(_hdlg, _caption);
	if (_message.length() > 0)
		SetDlgItemText(_hdlg, IDC_STATIC_MESSAGE, _message);
	if (_note.length() > 0)
	{
		SetDlgItemText(_hdlg, IDC_EDIT_NOTE, _note);
		PostMessage(GetDlgItem(_hdlg, IDC_EDIT_NOTE), EM_SETSEL, 0, 0);
	}

	// relocate the dialog if ProgressBox.Move has been called and has set a destination.
	_moveDialog();

	return TRUE;
}

/* responds to the cancel button selection by changing the _PI.canceled state to true and by reporting the event to ProgressBoxImpl which will forward the event to subscribing clients.
*/
BOOL ProgressBoxDlg::OnCancel()
{
	DBGPUTS((L"ProgressBoxDlg::OnCancel\n"));
	if (!_owner->fireCancel())
		return TRUE; // the client does not want to quit at this time.
	// update the state flag.
	WaitForSingleObject(_mutex, INFINITE);
	_PI.canceled = VARIANT_TRUE;
	ReleaseMutex(_mutex);
	// gray out the cancel button. don't call the base class method SimpleModelessDlg::OnCancel(). we want to keep the dialog running until the client app explicitly closes it.
	EnableWindow(GetDlgItem(_hdlg, IDCANCEL), FALSE);
	return TRUE;
}

/* handles WM_PBD messages sent from the clinet thread to the dialog UI thread to update the message text and progress display. The method hooks into SimpleDlg::DlgProc to handle Win32 messages not handled by the base class. WM_PBD messages are posted by the main thread set* methods. This method applies the data from a message to an associated control in the dialog UI thread. The _mutex is used to avoid a clash in accessing the member variable shared by the two threads.

A return value of TRUE means the method has handled the message. Before it calls us, SimpleDlg::DlgProc initializes lres to 0. SimpleDlg::DlgProc passes lres to the system. So, if a message handled by _subclassProc requires a value other than 0, make sure that this method assigns it to lres.
*/
BOOL ProgressBoxDlg::_subclassProc(UINT m_, WPARAM wp_, LPARAM lp_, LRESULT &lres)
{
	switch (m_)
	{
	case WM_PBD_DESTROY:
		// this message is from beforeDestroy in the main (client) thread. The client has called ProgressBox.Stop. ProgressBoxImpl called SimpleModelessDlg::Stop which then called beforeDestory() which posted this message.
		Destroy();
		return TRUE;

	case WM_PBD_MOVE:
		// move the dialog to a requested location.
		_moveDialog();
		return TRUE;

	case WM_PBD_SET_VISIBLE:
		// the message is from setVisible in response to a visibility change request from ProgressBox.Visible.
		ShowWindow(_hdlg, wp_ ? SW_SHOW : SW_HIDE);
		// make sure the dialog is at the top.
		SetForegroundWindow(_hdlg);
		return TRUE;

	case WM_PBD_SET_TEXT:
		WaitForSingleObject(_mutex, INFINITE);
		if (wp_ == TEXT_FIELD_CAPTION)
			SetWindowText(_hdlg, _caption); // update the dialog's caption bar with the new text.
		else if (wp_ == TEXT_FIELD_MESSAGE)
			SetDlgItemText(_hdlg, IDC_STATIC_MESSAGE, _message); // update the main message field.
		else if (wp_ == TEXT_FIELD_NOTE)
		{
			// update the note field.
			SetDlgItemText(_hdlg, IDC_EDIT_NOTE, _note);
			// if note is in append mode, move the caret to the beginning of the most recent (last) line in the edit control.
			if (_PI.options & PROGRESSBOXSTARTOPTION_APPEND_TO_NOTE)
			{
				LPCWSTR p = wcsrchr((LPCWSTR)_note, '\n');
				if (p++)
				{
					int n = (int)(p - _note._b);
					HWND hedit = GetDlgItem(_hdlg, IDC_EDIT_NOTE);
					PostMessage(hedit, EM_SETSEL, n, n);
					PostMessage(hedit, EM_SCROLLCARET, 0, 0);
					DBGPRINTF((L"EM_SETSEL: %d\n", n));
				}
			}
			else
				SendDlgItemMessage(_hdlg, IDC_EDIT_NOTE, EM_SETSEL, 0, 0);
		}
		ReleaseMutex(_mutex);
		return TRUE;

	case WM_PBD_PROGRESS_SHOW:
		// show or hide the progress bar in response to ProgressBar.ShowProgressBar or .HideProgressBar, respectively.
		if (_hProgess)
			ShowWindow(_hProgess, wp_ ? SW_SHOW : SW_HIDE);
		if (wp_)
			_PI.options |= PROGRESSBOXSTARTOPTION_SHOW_PROGRESSBAR;
		else
			_PI.options &= ~PROGRESSBOXSTARTOPTION_SHOW_PROGRESSBAR;
		return TRUE;

	case WM_PBD_PROGRESS_SETRANGE:
		// update the progress range upper or lower bound.
		WaitForSingleObject(_mutex, INFINITE);
		if (wp_ == PROGRESS_RANGE_LOW_BOUND)
			_PI.boundLow = LOWORD(lp_);
		else if (wp_ == PROGRESS_RANGE_HIGH_BOUND)
			_PI.boundHigh = LOWORD(lp_);
		if (_hProgess)
			::SendMessage(_hProgess, PBM_SETRANGE32, _PI.boundLow, _PI.boundHigh);
		ReleaseMutex(_mutex);
		return TRUE;

	case WM_PBD_PROGRESS_SETPOS:
		// update the progress position.
		WaitForSingleObject(_mutex, INFINITE);
		_PI.pos = LOWORD(wp_);
		if (_hProgess)
			::SendMessage(_hProgess, PBM_SETPOS, wp_, 0);
		ReleaseMutex(_mutex);
		return TRUE;

	case WM_PBD_PROGRESS_SETMARQUEE:
#ifdef PROGRESSBOX_SUPPORTS_MARQUEE
		// enable or disable the marquee mode, or update the marquee update time.
		WaitForSingleObject(_mutex, INFINITE);
		if (wp_) // enable marquee
			_PI.options |= PROGRESSBOXSTARTOPTION_MARQUEE;
		else
			_PI.options &= ~PROGRESSBOXSTARTOPTION_MARQUEE;
		_PI.marquee = (LONG)lp_;
		if (_hProgess)
			_setMarquee();
		ReleaseMutex(_mutex);
#endif//#ifdef PROGRESSBOX_SUPPORTS_MARQUEE
		return TRUE;

	case WM_PBD_PROGRESS_SETBARCOLOR:
		// update the bar color.
		WaitForSingleObject(_mutex, INFINITE);
		_PI.barColor = (COLORREF)wp_;
		if (_hProgess)
			::SendMessage(_hProgess, PBM_SETBARCOLOR, 0, wp_);
		ReleaseMutex(_mutex);
		return TRUE;
	}
	return FALSE;
}

// enable or disable the marquee mode, or update the marquee update time.
// _PI.marquee: 0=default marquee update time (30 ms), non-zero=custom marquee update time in ms.
void ProgressBoxDlg::_setMarquee()
{
#ifdef PROGRESSBOX_SUPPORTS_MARQUEE
	LONG styles = GetWindowLong(_hProgess, GWL_STYLE);

	if (_PI.options & PROGRESSBOXSTARTOPTION_MARQUEE)
	{
		// turn on the marquee mode and set the update time in ms.
		if (!(styles & PBS_MARQUEE))
		{
			styles |= PBS_MARQUEE;
			SetWindowLong(_hProgess, GWL_STYLE, styles);
		}
		SendMessage(_hProgess, PBM_SETMARQUEE, TRUE, _PI.marquee);
	}
	else
	{
		// turn off the marquee mode.
		if (styles & PBS_MARQUEE)
		{
			styles &= ~PBS_MARQUEE;
			SetWindowLong(_hProgess, GWL_STYLE, styles);
		}
		SendMessage(_hProgess, PBM_SETMARQUEE, FALSE, 0);
	}
#endif//#ifdef PROGRESSBOX_SUPPORTS_MARQUEE
}

/* moves the dialog box to a location requested by the client. we are invoked in the dialog thread after ProgressBox.Move calls moveDialog in the client thread. that's if a dialog has started. if dialog has not yet started, OnInitDialog calls us.
*/
void ProgressBoxDlg::_moveDialog()
{
	WaitForSingleObject(_mutex, INFINITE);
	PROGRESSBOXMOVEFLAG flag = _moveInfo.Flag;
	long x = _moveInfo.X;
	long y = _moveInfo.Y;
	ReleaseMutex(_mutex);
	DBGPRINTF((L"_moveDialog: flag=%d; x=%d, y=%d\n", flag, x, y));
	if (flag == PROGRESSBOXMOVEFLAG_NONE)
		return;
	// move the dialog box to screen center, left or right positions. the input x and y are ignored.
	RECT rcScreen, rcDlg;
	GetWindowRect(GetDesktopWindow(), &rcScreen);
	GetWindowRect(_hdlg, &rcDlg);
	if (flag == PROGRESSBOXMOVEFLAG_CENTER)
	{
		x = rcScreen.left + ((rcScreen.right - rcScreen.left) - (rcDlg.right - rcDlg.left)) / 2;
		y = rcScreen.top + ((rcScreen.bottom - rcScreen.top) - (rcDlg.bottom - rcDlg.top)) / 2;
	}
	else if (flag == PROGRESSBOXMOVEFLAG_LEFTTOP)
	{
		x = rcScreen.left;
		y = rcScreen.top;
	}
	else if (flag == PROGRESSBOXMOVEFLAG_RIGHTTOP)
	{
		x = rcScreen.right - (rcDlg.right - rcDlg.left);
		y = rcScreen.top;
	}
	else if (flag == PROGRESSBOXMOVEFLAG_LEFTBOTTOM)
	{
		x = rcScreen.left;
		y = rcScreen.bottom - (rcDlg.bottom - rcDlg.top);
	}
	else if (flag == PROGRESSBOXMOVEFLAG_RIGHTBOTTOM)
	{
		x = rcScreen.right - (rcDlg.right - rcDlg.left);
		y = rcScreen.bottom - (rcDlg.bottom - rcDlg.top);
	}
	// move the dialog.
	SetWindowPos(_hdlg, NULL, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
	// clear the destination flag so we won't repeat the relocation.
	WaitForSingleObject(_mutex, INFINITE);
	_moveInfo.Flag = PROGRESSBOXMOVEFLAG_NONE;
	ReleaseMutex(_mutex);
}

// returns a copy of the dialog caption in the main thread. A client reads ProgressBox.Caption. IProgressBox::get_Caption calls this method to pass the caption text to the client.
BSTR ProgressBoxDlg::getCaption()
{
	WaitForSingleObject(_mutex, INFINITE);
	BSTR val = _caption.clone();
	ReleaseMutex(_mutex);
	return val;
}

// returns a copy of the message in the main thread. A client reads ProgressBox.Message. IProgressBox::get_Message calls this method to pass the message text to the client.
BSTR ProgressBoxDlg::getMessage()
{
	WaitForSingleObject(_mutex, INFINITE);
	BSTR val = _message.clone();
	ReleaseMutex(_mutex);
	return val;
}

// returns a copy of the note in the main thread. A client reads ProgressBox.Note. IProgressBox::get_Note calls this method to pass the note text to the client.
BSTR ProgressBoxDlg::getNote()
{
	WaitForSingleObject(_mutex, INFINITE);
	BSTR val = _note.clone();
	ReleaseMutex(_mutex);
	return val;
}

/* assigns new text to the read-only Note edit control.

If the APPEND_TO_NOTE option is enabled, newVal plus a CRLF pair are appended to the existing text. The old text is scrolled up to ensure that the new text is visible.
*/
void ProgressBoxDlg::setNote(LPCWSTR newVal)
{
	WaitForSingleObject(_mutex, INFINITE);
	if (_PI.options & PROGRESSBOXSTARTOPTION_APPEND_TO_NOTE)
		_note.append(newVal);
	else
		_note.assignW(newVal);
	ReleaseMutex(_mutex);
	if (_hdlg)
		PostMessage(_hdlg, WM_PBD_SET_TEXT, TEXT_FIELD_NOTE, 0);
}

// returns the lower bound of the progress range.
long ProgressBoxDlg::getLowerBound()
{
	WaitForSingleObject(_mutex, INFINITE);
	long val = _PI.boundLow;
	ReleaseMutex(_mutex);
	return val;
}

// returns the upper bound of the progress range.
long ProgressBoxDlg::getUpperBound()
{
	WaitForSingleObject(_mutex, INFINITE);
	long val = _PI.boundHigh;
	ReleaseMutex(_mutex);
	return val;
}

// returns the progress position.
long ProgressBoxDlg::getProgressPos()
{
	WaitForSingleObject(_mutex, INFINITE);
	long val = _PI.pos;
	ReleaseMutex(_mutex);
	return val;
}

// returns the current bar color in RGB. the default color is used when this prop is set to CLR_DEFAULT (0xFF000000).
COLORREF ProgressBoxDlg::getBarColor()
{
	WaitForSingleObject(_mutex, INFINITE);
	COLORREF val = _PI.barColor;
	ReleaseMutex(_mutex);
	return val;
}

// returns the boolean state of the cancel button.
VARIANT_BOOL ProgressBoxDlg::getCanceled()
{
	WaitForSingleObject(_mutex, INFINITE);
	VARIANT_BOOL val = _PI.canceled;
	ReleaseMutex(_mutex);
	return val;
}

// returns the boolean state of the dialog's visibility.
VARIANT_BOOL ProgressBoxDlg::getVisible()
{
	if (IsWindowVisible(_hdlg))
		return VARIANT_TRUE;
	return VARIANT_FALSE;
}

///////////////////////////////////////////////////////////////////
// this helper returns it if the input arg has an integer value, or returns a default value provided by the caller if arg is not integer-typed.
long parseOptionalIntArg(VARIANT *arg, long defaultValue)
{
	if (arg && arg->vt != VT_ERROR)
	{
		if (arg->vt == VT_I4 || arg->vt == VT_UI4)
			return arg->ulVal;
		else if (arg->vt == VT_I2 || arg->vt == VT_UI2)
			return MAKELONG(arg->uiVal, 0);
		else if (arg->vt == VT_R8)
			return (long)arg->dblVal;
		else if (arg->vt == VT_R4)
			return (long)arg->fltVal;
	}
	return defaultValue;
}

