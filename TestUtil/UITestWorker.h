#pragma once
#include "IDispatchImpl.h"


class UITestWorker
{
public:
	UITestWorker(HWND h=NULL) : _hwnd(h), _hdlg(NULL), _dlgCaption(NULL), _thread(NULL), _killWorker(NULL), _workerDone(NULL) {}
	virtual ~UITestWorker() { Stop(); }

protected:
	HWND _hwnd; // who owns us.
	HWND _hdlg; // target dialog.
	BSTR _dlgCaption; // dialog's caption text

	virtual HRESULT Start(BSTR dlgCaption)
	{
		_dlgCaption = dlgCaption;
		// the thread can signal start and done events.
		_killWorker = CreateEvent(NULL, TRUE, FALSE, NULL);
		_workerDone = CreateEvent(NULL, TRUE, FALSE, NULL);
		// start a worker thread. the manager will start it.
		_thread = CreateThread(NULL, 0, _threadMan, (LPVOID)this, 0, NULL);
		if (_thread)
			return S_OK;
		// worker creation failed.
		return E_FAIL;
	}
	virtual void Stop()
	{
		// stop the worker thread.
		HANDLE hThread = InterlockedExchangePointer(&_thread, NULL);
		if (hThread)
		{
			// kill the worker.
			SetEvent(_killWorker);
			// wait until it exits.
			::WaitForSingleObject(hThread, INFINITE);
			CloseHandle(hThread);
		}
		HANDLE hDone = InterlockedExchangePointer(&_workerDone, NULL);
		HANDLE hKill = InterlockedExchangePointer(&_killWorker, NULL);
		if (hDone)
			CloseHandle(hDone);
		if (hKill)
			CloseHandle(hKill);
		_dlgCaption = NULL;
		_hdlg = NULL;
	}

	bool Killed()
	{
		if (WaitForSingleObject(_killWorker, 0) == WAIT_OBJECT_0)
			return true;
		return false;
	}

	void printHwnd(LPCSTR label, HWND hwnd)
	{
		WCHAR bufClass[256];
		GetClassName(hwnd, bufClass, ARRAYSIZE(bufClass));
		WCHAR bufTitle[256];
		GetWindowText(hwnd, bufTitle, ARRAYSIZE(bufTitle));
		printf("%s: %p; '%S'; '%S'\n", label, hwnd, bufClass, bufTitle);
	}

	virtual bool runTest();

private:
	HANDLE _thread; // handle of the dialog thread.
	HANDLE _killWorker, _workerDone; // events for signalling starting and clsoing of the thread.

	static ULONG WINAPI _threadMan(LPVOID lpParam) {
		return ((UITestWorker*)lpParam)->Run();
	}

	ULONG Run()
	{
		runTest();

		// done. tell the primary thread we're terminating.
		SetEvent(_workerDone);
		return ERROR_SUCCESS;
	}
};

class UITestWorker_InputBox : public UITestWorker
{
public:
	UITestWorker_InputBox(HWND h) : UITestWorker(h), _params{ 0 } {}

	HRESULT StartTest(LPCSTR caseLabel, BSTR dlgCaption, UINT buttonId, UINT inputId = 0, BSTR inputText = NULL, UINT checkboxId = 0, BOOL checked = 0)
	{
		_params.caseLabel = caseLabel;
		_params.inputText = inputText;
		_params.inputId = inputId;
		_params.buttonId = buttonId;
		_params.checkboxId = checkboxId;
		_params.checked = checked;
		return UITestWorker::Start(dlgCaption);
	}
	void StopTest()
	{
		_params = { 0 };
		UITestWorker::Stop();
	}

protected:
	struct Params
	{
		LPCSTR caseLabel;
		BSTR inputText;
		UINT inputId, checkboxId, buttonId;
		BOOL checked;
	} _params;

	virtual bool runTest();
};

class UITestWorker_ProgressBox : public UITestWorker, public IDispatchEventSinkAdviseImpl<IDispatch, &DIID_ProgressBoxEvents>
{
public:
	UITestWorker_ProgressBox(LPUNKNOWN punk) : _punk(punk), _params{ 0 } {}
	~UITestWorker_ProgressBox()
	{
		ASSERT(_ref == 1);
		_ref--;
	}

	HRESULT StartTest(LPCSTR caseLabel, BSTR dlgCaption, UINT buttonId, long delayInMilliseconds)
	{
		if (_punk)
		{
			HRESULT hr = IDispatchEventSinkAdviseImpl<IDispatch, &DIID_ProgressBoxEvents>::Advise(_punk);
			if (hr != S_OK)
				return hr;
		}
		_params.caseLabel = caseLabel;
		_params.buttonId = buttonId;
		_params.delayInMilliseconds = delayInMilliseconds;
		return UITestWorker::Start(dlgCaption);
	}
	void StopTest()
	{
		if (_punk)
			IDispatchEventSinkAdviseImpl<IDispatch, &DIID_ProgressBoxEvents>::Unadvise();
		_params = { 0 };
		UITestWorker::Stop();
	}

	bool IsProgressBoxCanceled()
	{
		return _params.cancelFired;
	}
	int ElapsedTimeSinceCancel()
	{
		long dt = -1;
		if (_params.cancelFired)
			dt = (long)(_params.tckCanceled - _params.tckClicked);
		return dt;
	}

protected:
	LPUNKNOWN _punk;
	struct Params
	{
		LPCSTR caseLabel;
		UINT buttonId;
		long delayInMilliseconds;
		int cancelFired;
		DWORD tckClicked;
		DWORD tckCanceled;
	} _params;

	virtual bool runTest();

	STDMETHOD(Invoke) (DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS FAR* pdispparams, VARIANT FAR* pvarResult, EXCEPINFO FAR* pexcepinfo, UINT FAR* puArgErr);
};

