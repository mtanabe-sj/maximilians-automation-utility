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
#include <Windows.h>
#include "SimpleDlg.h"
#include "MaxsUtil_h.h"
#include "IDispatchImpl.h"
#include "resource.h"

#define PROGRESSBOX_SUPPORTS_EVENT
#ifdef PROGRESSBOX_SUPPORTS_EVENT
#include "ConnectionPointImpl.h"
#endif//#ifdef PROGRESSBOX_SUPPORTS_EVENT


long parseOptionalIntArg(VARIANT *arg, long defaultValue = 0);


// a forward declaration needed by ProgressBoxDlg.
class ProgressBoxImpl;

/* This is a handler class for a modeless dialog that reports job progress and displays related message text. ProgressBoxImpl uses this handler to run a progress dialog in a UI thread different from the client's thread. The handler offers set* and get* methods for accessing the parameters of progress range and position and related status text settings in the client thread. When a new value for a progress parameter (e.g., the progress position) is passed in from the client, ProgressBoxImpl uses a set method (e.g., setProgressPos) to forward the new value to the dialog. The dialog handler then caches the value in a class member variable (e.g., _PI.pos) and posts a special Win32 message (WM_PBD_*) to the dialog thread. On receiving the message, the dialog thread (method _subclassProc) reads the saved value and assigns it to the UI control. To arbitrate contentious variable access between the threads, the handler uses a mutex.
*/
class ProgressBoxDlg : public SimpleModelessDlg
{
public:
	ProgressBoxDlg(ProgressBoxImpl *owner) :
		SimpleModelessDlg(IDD_PROGRESSBOX),
		_owner(owner),
		_hProgess(NULL),
		_PI{ 0 }
	{
		// use a mutex to synchronize internal and external access to variables.
		_mutex = CreateMutex(NULL, FALSE, NULL);
		_moveInfo.Flag = PROGRESSBOXMOVEFLAG_NONE;
		_PI.barColor = CLR_DEFAULT;
	}
	~ProgressBoxDlg()
	{
		HANDLE h = InterlockedExchangePointer((LPVOID*)&_mutex, NULL);
		if (h)
			CloseHandle(h);
	}

	struct PROGRESSINFO
	{
		VARIANT_BOOL canceled; // set to VARIANT_TRUE when the Cancel button is selected.
		long boundLow, boundHigh, pos; // range bounds and current position of the progress bar.
		long options; // PROGRESSBOXSTARTOPTION bits
		long marquee;
		COLORREF barColor;
	} _PI;

	enum WM_PBD {
		WM_PBD_DESTROY = WM_USER + 100,
		WM_PBD_MOVE,
		WM_PBD_SET_VISIBLE,
		WM_PBD_SET_TEXT,
		WM_PBD_PROGRESS_SHOW,
		WM_PBD_PROGRESS_SETRANGE,
		WM_PBD_PROGRESS_SETPOS,
		WM_PBD_PROGRESS_SETMARQUEE,
		WM_PBD_PROGRESS_SETBARCOLOR,
	};

	enum TEXT_FIELD {
		TEXT_FIELD_CAPTION=0,
		TEXT_FIELD_MESSAGE,
		TEXT_FIELD_NOTE,
	};
	enum PROGRESS_RANGE {
		PROGRESS_RANGE_LOW_BOUND=0,
		PROGRESS_RANGE_HIGH_BOUND
	};

	virtual void beforeDestroy()
	{
		DBGPUTS((L"ProgressBoxDlg::beforeDestroy\n"));
		// tell the UI thread to stop.
		PostMessage(_hdlg, WM_PBD_DESTROY, 0, 0);
	}

	// delegated from IProgressBox property put calls
	void setCaption(LPCWSTR newVal)
	{
		WaitForSingleObject(_mutex, INFINITE);
		_caption = newVal;
		ReleaseMutex(_mutex);
		if (_hdlg)
			PostMessage(_hdlg, WM_PBD_SET_TEXT, TEXT_FIELD_CAPTION, 0);
	}
	void setMessage(LPCWSTR newVal)
	{
		WaitForSingleObject(_mutex, INFINITE);
		_message = newVal;
		ReleaseMutex(_mutex);
		if (_hdlg)
			PostMessage(_hdlg, WM_PBD_SET_TEXT, TEXT_FIELD_MESSAGE, 0);
	}
	void setNote(LPCWSTR newVal);
	void setLowerBound(long newVal)
	{
		if (_hdlg)
			PostMessage(_hdlg, WM_PBD_PROGRESS_SETRANGE, PROGRESS_RANGE_LOW_BOUND, newVal);
		else
			_PI.boundLow = newVal;
	}
	void setUpperBound(long newVal)
	{
		if (_hdlg)
			PostMessage(_hdlg, WM_PBD_PROGRESS_SETRANGE, PROGRESS_RANGE_HIGH_BOUND, newVal);
		else
			_PI.boundHigh = newVal;
	}
	void setBarColor(COLORREF newVal)
	{
		if (_hdlg)
			PostMessage(_hdlg, WM_PBD_PROGRESS_SETBARCOLOR, newVal, 0);
		else
			_PI.barColor = newVal;
	}
	void setProgressPos(long newVal)
	{
		if (_hdlg)
			PostMessage(_hdlg, WM_PBD_PROGRESS_SETPOS, newVal, 0);
		else
			_PI.pos = newVal;
	}
	void setVisible(VARIANT_BOOL newVal)
	{
		if (_hdlg)
			PostMessage(_hdlg, WM_PBD_SET_VISIBLE, MAKEWPARAM(MAKEWORD(newVal,0), 0), 0);
	}

	// delegated from IProgressBox property get calls
	BSTR getCaption();
	BSTR getMessage();
	BSTR getNote();
	long getLowerBound();
	long getUpperBound();
	long getProgressPos();
	COLORREF getBarColor();
	VARIANT_BOOL getCanceled();
	VARIANT_BOOL getVisible();

	void showProgressBar(bool newState) { PostMessage(_hdlg, WM_PBD_PROGRESS_SHOW, MAKEWPARAM(newState, 0), 0); }
	// set PBMOVE_CENTER in flags to center the dialog. x and y will be ignored.
	void moveDialog(long flags, long x, long y)
	{
		DBGPRINTF((L"moveDialog: flag=%d; x=%d, y=%d\n", flags, x, y));
		WaitForSingleObject(_mutex, INFINITE);
		_moveInfo.Flag = (PROGRESSBOXMOVEFLAG)flags;
		_moveInfo.X = x;
		_moveInfo.Y = y;
		ReleaseMutex(_mutex);
		if (_hdlg)
			PostMessage(_hdlg, WM_PBD_MOVE, 0, 0);
	}

	void setOptions(long options, VARIANT *params);

protected:
	ProgressBoxImpl *_owner; // a ProgressBoxImpl instance who owns and calls us.
	HANDLE _mutex; // a mutex to arbitrate variable access between the owner's main and dialog's UI threads.
	HWND _hProgess; // window handle of the progress bar.
	bstring _caption; // caption text of the progress dialog.
	bstring _message; // single-line text assigned to IDC_STATIC_MESSAGE.
	bstringCRLF _note; // CRLF-formatted multi-line text assigned to IDC_EDIT_NOTE.
	struct MOVEINFO {
		PROGRESSBOXMOVEFLAG Flag;
		long X, Y;
	} _moveInfo;

	void _moveDialog();
	void _setMarquee();

	virtual BOOL OnInitDialog();
	virtual BOOL OnCancel();
	virtual BOOL _subclassProc(UINT m_, WPARAM wp_, LPARAM lp_, LRESULT &lres);
};


/* Implements coclass ProgressBox exposing IProgressBox. It also exposes IConnectionPointContainer to provide event notification.

This is a subclass of a threaded dialog class. It runs a modeless dialog and displays job progress. COM clients use IProgressBox to control aspects of the progress display and manage the life of the dialog. Clients can subscribe to the IConnectionPoint notification and automatically receive a call whenever the dialog is canceled.
*/
class ProgressBoxImpl :
	public SimpleModelessDlgThread,
#ifdef PROGRESSBOX_SUPPORTS_EVENT
	public ConnectionPointCallback,
	public IConnectionPointContainer,
	public IEnumConnectionPoints,
#endif//#ifdef PROGRESSBOX_SUPPORTS_EVENT
	public IDispatchWithObjectSafetyImpl<IProgressBox, &IID_IProgressBox, &LIBID_MaxsUtilLib>
{
public:
	ProgressBoxImpl();
	~ProgressBoxImpl();

	// IUnknown methods
	DELEGATE_IUNKNOWN_REF_TO_IDISPATCHWITHOBJECTSAFETYIMPL(IProgressBox, &IID_IProgressBox, &LIBID_MaxsUtilLib)
	STDMETHOD(QueryInterface)(REFIID riid, LPVOID* ppvObj)
	{
		AddRef();
		if (riid == IID_IConnectionPointContainer)
		{
			*ppvObj = dynamic_cast<IConnectionPointContainer*>(this);
			return S_OK;
		}
		else if (riid == IID_IEnumConnectionPoints)
		{
			*ppvObj = dynamic_cast<IEnumConnectionPoints*>(this);
			return S_OK;
		}
		Release();
		return IDispatchWithObjectSafetyImpl<IProgressBox, &IID_IProgressBox, &LIBID_MaxsUtilLib>::QueryInterface(riid, ppvObj);
	}

#ifdef PROGRESSBOX_SUPPORTS_EVENT
	// *** IConnectionPointContainer ***
	STDMETHODIMP EnumConnectionPoints(IEnumConnectionPoints **ppEnum) {
		return _cplist.EnumConnectionPoints(ppEnum);
	}
	STDMETHODIMP FindConnectionPoint(REFIID riid, IConnectionPoint **ppCP) {
		return _cplist.FindConnectionPoint(riid, ppCP);
	}
	// *** IEnumConnectionPoints methods ***
	STDMETHODIMP Next(ULONG cConnections, IConnectionPoint **rgpcn, ULONG *pcFetched) {
		return _cplist.Next(cConnections, rgpcn, pcFetched);
	}
	STDMETHODIMP Skip(ULONG cConnections) {
		return _cplist.Skip(cConnections);
	}
	STDMETHODIMP Reset() {
		return _cplist.Reset();
	}
	STDMETHODIMP Clone(IEnumConnectionPoints **ppEnum) {
		return _cplist.Clone(ppEnum);
	}
	// fires a cancel event. ProgressBoxDlg::OnCancel calls us from the dialog thread.
	VARIANT_BOOL fireCancel();
#endif//#ifdef PROGRESSBOX_SUPPORTS_EVENT

	// IProgressBox property Caption
	STDMETHOD(get_Caption)(/* [retval][out] */ BSTR *pbsData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		*pbsData = static_cast<ProgressBoxDlg*>(_dlg)->getCaption();
		return S_OK;
	}
	STDMETHOD(put_Caption)(/* [in] */ BSTR bsData)
	{
		if (!_dlg)
			_dlg = new ProgressBoxDlg(this);
		static_cast<ProgressBoxDlg*>(_dlg)->setCaption(bsData);
		return S_OK;
	}
	// IProgressBox property Message
	STDMETHOD(get_Message)(/* [retval][out] */ BSTR *pbsData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		*pbsData = static_cast<ProgressBoxDlg*>(_dlg)->getMessage();
		return S_OK;
	}
	STDMETHOD(put_Message)(/* [in] */ BSTR bsData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		static_cast<ProgressBoxDlg*>(_dlg)->setMessage(bsData);
		return S_OK;
	}
	// IProgressBox property Note
	STDMETHOD(get_Note)(/* [retval][out] */ BSTR *pbsData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		*pbsData = static_cast<ProgressBoxDlg*>(_dlg)->getNote();
		return S_OK;
	}
	STDMETHOD(put_Note)(/* [in] */ BSTR bsData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		static_cast<ProgressBoxDlg*>(_dlg)->setNote(bsData);
		return S_OK;
	}
	// IProgressBox property LowerBound
	STDMETHOD(get_LowerBound)(/* [retval][out] */ long *plData)
	{
		if (!_dlg)
			*plData = _stoppedProgInfo.boundLow; // read it from the cache since _dlg has already been stopped.
		else
			*plData = static_cast<ProgressBoxDlg*>(_dlg)->getLowerBound();
		return S_OK;
	}
	STDMETHOD(put_LowerBound)(/* [in] */ long lData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		static_cast<ProgressBoxDlg*>(_dlg)->setLowerBound(lData);
		return S_OK;
	}
	// IProgressBox property UpperBound
	STDMETHOD(get_UpperBound)(/* [retval][out] */ long *plData)
	{
		if (!_dlg)
			*plData = _stoppedProgInfo.boundHigh; // read it from the cache since _dlg has already been stopped.
		else
			*plData = static_cast<ProgressBoxDlg*>(_dlg)->getUpperBound();
		return S_OK;
	}
	STDMETHOD(put_UpperBound)(/* [in] */ long lData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		static_cast<ProgressBoxDlg*>(_dlg)->setUpperBound(lData);
		return S_OK;
	}
	// IProgressBox property ProgressPos
	STDMETHOD(get_ProgressPos)(/* [retval][out] */ long *plData)
	{
		if (!_dlg)
			*plData=_stoppedProgInfo.pos; // read it from the cache since _dlg has already been stopped.
		else
			*plData = static_cast<ProgressBoxDlg*>(_dlg)->getProgressPos();
		return S_OK;
	}
	STDMETHOD(put_ProgressPos)(/* [in] */ long lData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		static_cast<ProgressBoxDlg*>(_dlg)->setProgressPos(lData);
		return S_OK;
	}
	// IProgressBox property Canceled
	STDMETHOD(get_Canceled)(/* [retval][out] */ VARIANT_BOOL *pbData)
	{
		if (!_dlg)
			*pbData = _stoppedProgInfo.canceled; // read it from the cache since _dlg has already been stopped.
		else
			*pbData = static_cast<ProgressBoxDlg*>(_dlg)->getCanceled();
		return S_OK;
	}
	// IProgressBox property WindowHandle
	STDMETHOD(get_WindowHandle)(/* [retval][out] */ OLE_HANDLE *pohData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		*pohData = (OLE_HANDLE)(ULONG_PTR)(HWND)(*_dlg);
		return S_OK;
	}
	// IProgressBox property VIsible
	STDMETHOD(get_Visible)(/* [retval][out] */ VARIANT_BOOL *pbData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		*pbData = static_cast<ProgressBoxDlg*>(_dlg)->getVisible();
		return S_OK;
	}
	STDMETHOD(put_Visible)(/* [in] */ VARIANT_BOOL bData)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		static_cast<ProgressBoxDlg*>(_dlg)->setVisible(bData);
		return S_OK;
	}
	STDMETHOD(get_BarColor)(/* [retval][out] */ long *Value)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		*Value = static_cast<ProgressBoxDlg*>(_dlg)->getBarColor();
		return S_OK;
	}
	STDMETHOD(put_BarColor)(/* [in] */ long NewValue)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		static_cast<ProgressBoxDlg*>(_dlg)->setBarColor((COLORREF)NewValue);
		return S_OK;
	}

	// ShowProgressBar - [method] shows (or unhides) the progress bar.
	STDMETHOD(ShowProgressBar)(void)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		static_cast<ProgressBoxDlg*>(_dlg)->showProgressBar(true);
		return S_OK;
	}
	STDMETHOD(HideProgressBar)(void)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		static_cast<ProgressBoxDlg*>(_dlg)->showProgressBar(false);
		return S_OK;
	}
	/* Start - [method] starts a modeless dialog in a separate thread and returns. The dialog does not block the calling client's thread. If a dialog is already running, only the supplied option settings are applied. Another dialog is not started.

	Parameters:
	Options - [optional][in] one or more PROGRESSBOXSTARTOPTION bit flags.
	Params - [optional][in] a parameter associated with the option(s) in Options. The data type Params takes on depends on the associated option.

	Return value:
	S_OK - a dialog has started and is running.
	S_FALSE - a dialog is already running. If any option has been enabled, the operation was successful.
	E_FAIL - a dialog could not be started due to an error.

	Remarks:
	The Start method may be called multiple times before Stop is called. The first time Start is called, a dialog is created and started. The second time it is called, the new Options and Params settings are applied to the existing dialog. For example, a client initially does not enable any option and so starts a dialog with the cancel button enabled. Later, the client does not wish to keep the button enabled, and so, calls Start again with PROGRESSBOXSTARTOPTION_DISABLE_CANCEL. The cancel button is grayed out and so is disabled.
	*/
	STDMETHOD(Start)(/* [optional][in] */ VARIANT *Options, /* [optional][in] */ VARIANT *Params)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		long options = parseOptionalIntArg(Options);
		if (options)
			static_cast<ProgressBoxDlg*>(_dlg)->setOptions(options, Params);
		if (*_dlg) // check the window handle (HWND)
			return S_FALSE; // already started.
		if (SimpleModelessDlgThread::Start())
			return S_OK; // just started.
		return E_FAIL; // failed to start.
	}
	/* Stop - [method] stops and closes the modeless dialog.
	*/
	STDMETHOD(Stop)(void)
	{
		if (!_dlg)
			return S_FALSE; // already gone.
		CopyMemory(&_stoppedProgInfo, &static_cast<ProgressBoxDlg*>(_dlg)->_PI, sizeof(ProgressBoxDlg::PROGRESSINFO));
		SimpleModelessDlgThread::Stop();
		return S_OK;
	}
	// Increment - [method] increments current progress position by the amount in Step or by one if Step is not specified.
	STDMETHOD(Increment)(/* [optional][in] */ VARIANT *Step)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		long delta = parseOptionalIntArg(Step, 1);
		long val = static_cast<ProgressBoxDlg*>(_dlg)->getProgressPos();
		static_cast<ProgressBoxDlg*>(_dlg)->setProgressPos(val+delta);
		return S_OK;
	}
	// Move - [method] moves the dialog window to the center of the screen or to a point given by X and Y.
	STDMETHOD(Move)(/* [in] */ VARIANT Flags, /* [optional][in] */ VARIANT *X, /* [optional][in] */ VARIANT *Y)
	{
		if (!_dlg)
			return E_UNEXPECTED; // already stopped.
		static_cast<ProgressBoxDlg*>(_dlg)->moveDialog(parseOptionalIntArg(&Flags), parseOptionalIntArg(X), parseOptionalIntArg(Y));
		if (!_dlg)
			return S_FALSE;
		return S_OK;
	}

protected:
	ProgressBoxDlg::PROGRESSINFO _stoppedProgInfo;

#ifdef PROGRESSBOX_SUPPORTS_EVENT
	ConnectionPointListImpl _cplist;

	// ConnectionPointCallback methods
	void GetIID(IID* pIID) { *pIID = DIID_ProgressBoxEvents; };
	void OnAdviseConnectionPoint();
	void OnUnadviseConnectionPoint();
#endif//#ifdef PROGRESSBOX_SUPPORTS_EVENT

	/* SimpleModelessDlgThread overrides are used to maintain a reference on the ProgressBox instance. It's to prevent premature destruction of the instance, which happens if the client releases the interface unexpectedly before the dialog closes normally (e.g., the Cancel button is clicked by a user).
	*/
	virtual void beforeDialogRun()
	{
		DBGPUTS((L"ProgressBoxImpl::beforeDialogRun\n"));
		// the dialog is ready to run. so, get a reference to make sure this ProgressBox instance is kept alive until the dialog closes.
		AddRef();
	}
	virtual void afterDialogRun()
	{
		DBGPUTS((L"ProgressBoxImpl::afterDialogRun\n"));
		// the dialog has just closed. release the reference so the ProgressBox can be freed shortly.
		Release();
	}
};

