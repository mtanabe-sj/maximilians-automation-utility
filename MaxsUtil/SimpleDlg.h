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


// Implements a basic win32 dialog handler providing essential plumbing work. Override virtuals to respond to win32 messages you are interested in in your subclass.
class SimpleDlg
{
public:
	SimpleDlg(UINT idd, HWND hparent = NULL) : _idd(idd), _hwndParent(hparent), _hdlg(NULL) {}
	virtual ~SimpleDlg() {}

	INT_PTR DoModal()
	{
		return DialogBoxParam(
			LibInstanceHandle,
			MAKEINTRESOURCE(_idd),
			_hwndParent,
			DlgProc,
			(LPARAM)(LPVOID)this);
	}

protected:
	UINT _idd;
	HWND _hdlg;
	HWND _hwndParent;

	virtual BOOL OnInitDialog() { return TRUE; }
	virtual void OnDestroy() { _hdlg = NULL; }
	virtual BOOL OnOK() { EndDialog(_hdlg, IDOK); return TRUE; }
	virtual BOOL OnCancel() { EndDialog(_hdlg, IDCANCEL); return TRUE; }
	virtual BOOL OnCommand(WPARAM wp_, LPARAM lp_)
	{
		if (LOWORD(wp_) == IDOK)
			return OnOK();
		if (LOWORD(wp_) == IDCANCEL)
			return OnCancel();
		return FALSE;
	}

protected:
	static INT_PTR CALLBACK DlgProc(HWND h_, UINT m_, WPARAM wp_, LPARAM lp_)
	{
		SimpleDlg* pThis = (SimpleDlg*)GetWindowLongPtr(h_, DWLP_USER);
		if (m_ == WM_INITDIALOG)
		{
			pThis = (SimpleDlg*)lp_;
			pThis->_hdlg = h_;
			::SetWindowLongPtr(h_, DWLP_USER, (ULONG_PTR)(LPVOID)pThis);
			return pThis->OnInitDialog();
		}
		else if (m_ == WM_DESTROY)
		{
			if (pThis)
				pThis->OnDestroy();
			SetWindowLongPtr(h_, DWLP_USER, NULL);
			return TRUE;
		}
		else if (m_ == WM_COMMAND)
		{
			if (pThis && pThis->OnCommand(wp_, lp_))
			{
				::SetWindowLongPtr(h_, DWLP_MSGRESULT, TRUE);
				return TRUE;
			}
		}
		else if (pThis)
		{
			// delegate handling of other messages (e.g., WM_USER) to the subclass.
			LRESULT lres=0;
			if (pThis->_subclassProc(m_, wp_, lp_, lres))
			{
				::SetWindowLongPtr(h_, DWLP_MSGRESULT, lres);
				return TRUE;
			}
		}
		return FALSE;
	}
	virtual BOOL _subclassProc(UINT m_, WPARAM wp_, LPARAM lp_, LRESULT &lres)
	{
		return FALSE;
	}
};


/* This is a subclass of SimpleDlg implementing a modeless dialog handler. Use method Create to start running the dialog. Call method Destroy to close it.
*/
class SimpleModelessDlg : public SimpleDlg
{
public:
	SimpleModelessDlg(UINT idd, HWND hparent = NULL) : SimpleDlg(idd, hparent), _result(0) {}
	virtual ~SimpleModelessDlg() {}
	UINT _result; // will be set to IDOK or IDCANCEL when the dialog closes.

	operator HWND() const { return _hdlg; }
	bool Create()
	{
		DBGPUTS((L"SimpleModelessDlg::Create\n"));
		if (::CreateDialogParam(
			LibInstanceHandle,
			MAKEINTRESOURCE(_idd),
			_hwndParent,
			(DLGPROC)DlgProc,
			(LPARAM)this))
			return true;
		return false;
	}
	void Destroy()
	{
		if (_hdlg)
		{
			DBGPUTS((L"SimpleModelessDlg::Destroy\n"));
			::PostQuitMessage(0);
			::DestroyWindow(_hdlg);
			_hdlg = NULL;
		}
	}

	virtual void beforeDestroy() {}

protected:
	// overrides SimpleDlg::OnOK (and OnCancel) to prevent calling EndDialog which is for a modal dialog and does not work for us.
	virtual BOOL OnOK()
	{
		DBGPUTS((L"SimpleModelessDlg::OnOK\n"));
		Destroy();
		_result = IDOK;
		return TRUE;
	}
	virtual BOOL OnCancel()
	{
		DBGPUTS((L"SimpleModelessDlg::OnCancel\n"));
		Destroy();
		_result = IDCANCEL;
		// don't call the base class method since it calls EndDialog which is for a modal dialog.
		return TRUE;
	}
};

/* Runs a modeless dialog in a separate UI thread.
Subclass this base class. Create and assign a SimpleModelessDlg to the subclass instance. Call method Start to start a thread and run the dialog. Call method Stop to close the dialog and exit the thread.
*/
class SimpleModelessDlgThread
{
public:
	SimpleModelessDlgThread() : _dlg(NULL) {}
	~SimpleModelessDlgThread() { Stop(); }

	bool Start()
	{
		DBGPUTS((L"SimpleModelessDlgThread::Start\n"));
		ASSERT(_thread == NULL);
		// the dialog thread can signal start and done events.
		_dlgReady = CreateEvent(NULL, TRUE, FALSE, NULL);
		_dlgClosed = CreateEvent(NULL, TRUE, FALSE, NULL);
		// start a dialog thread. the manager will start it.
		_thread = CreateThread(NULL, 0, _threadMan, (LPVOID)this, 0, NULL);
		if (_thread)
		{
			// wait until the dialog starts and is ready. closed means it failed to start.
			HANDLE h[2] = { _dlgReady, _dlgClosed };
			DWORD res = WaitForMultipleObjects(ARRAYSIZE(h), h, FALSE, INFINITE);
			if (res == WAIT_OBJECT_0)
				return true;
			// dialog creation failed.
		}
		return false;
	}
	void Stop()
	{
		DBGPUTS((L"SimpleModelessDlgThread::Stop\n"));
		SimpleModelessDlg *dlg = (SimpleModelessDlg*)InterlockedExchangePointer((LPVOID*)&_dlg, NULL);
		if (!dlg)
			return;

		// stop the dialog thread.
		HANDLE hThread = InterlockedExchangePointer(&_thread, NULL);
		if (hThread)
		{
			dlg->beforeDestroy();
			// wait until it exits.
			if (WAIT_OBJECT_0 != ::WaitForSingleObject(hThread, INFINITE))
			{
				ASSERT(FALSE);
			}
			CloseHandle(hThread);
		}
		HANDLE hDataavail = InterlockedExchangePointer(&_dlgClosed, NULL);
		HANDLE hWorkerready = InterlockedExchangePointer(&_dlgReady, NULL);
		if (hDataavail)
			CloseHandle(_dlgClosed);
		if (hWorkerready)
			CloseHandle(_dlgReady);

		delete dlg;
	}

protected:
	SimpleModelessDlg *_dlg; // assign your modeless handler to this before calling Start.
	
	// overridable
	virtual void beforeDialogRun() = 0; // called by Run before a dialog ready is signalled and before a message pump is started.
	virtual void afterDialogRun() = 0; // called by Run after the dialog closes and before a dialog closed is signalled.

private:
	HANDLE _thread; // handle of the dialog thread.
	HANDLE _dlgReady, _dlgClosed; // events for signalling starting and clsoing of the thread.

	static ULONG WINAPI _threadMan(LPVOID lpParam) {
		return ((SimpleModelessDlgThread*)lpParam)->Run();
	}
	// starts a modeless dialog and runs a message pump to keep it running. waits until the dialog closes either by a user action or by a programmatic call to our Destroy() method.
	ULONG Run()
	{
		ASSERT(_dlg);
		// start a modeless dialog.
		if (!_dlg->Create())
		{
			HANDLE h = InterlockedExchangePointer(&_dlgClosed, NULL);
			SetEvent(h); // tell primary thread that dialog creation failed.
			return ERROR_INTERNAL_ERROR;
		}
		beforeDialogRun();
		// keep the window handle on the stack. _dlg can be invalidated before we are done with the message pump below.
		HWND hdlg = *_dlg;
		// make sure our dialog box is pushed forward, not hidden behind the client's window.
		SetForegroundWindow(hdlg);
		// the dialog is up and running.
		SetEvent(_dlgReady);

		try {
			// pump messages for the dialog. SimpleModelessDlg::OnDestroy() posts WM_QUIT to stop the message pump.
			MSG msg;
			while (GetMessage(&msg, (HWND)NULL, 0, 0))
			{
				if (IsDialogMessage(hdlg, &msg))
					continue; // already sent.
				TranslateMessage(&msg);
				DispatchMessage(&msg);
			}
		}
		catch (...) {
		}

		afterDialogRun();
		// done. tell the primary thread we're terminating.
		SetEvent(_dlgClosed);
		return ERROR_SUCCESS;
	}
};