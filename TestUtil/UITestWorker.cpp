#include "pch.h"
#include "UITestWorker.h"

using namespace std;


HMODULE LibInstanceHandle = NULL;
ULONG LibRefCount = 0;

bool UITestWorker::runTest()
{
	int i;
	for (i = 0; i < 10; i++)
	{
		if (Killed())
			return false;
		Sleep(1000);
		// note that we cannot pass _hwnd as the first param to the find api and expect to find the target dialog. a dialog class does not use the WS_params.CHILD style bit, and so, won't be turned up in a children search. so, we pass a NULL instead and run a global search for the dialog.
		if ((_hdlg = FindWindowEx(NULL, NULL, _T("#32770"), _dlgCaption)))
			break;
	}
	if (!_hdlg)
	{
		cout << "\trunInputBoxTestCase timed out." << endl;
		return false;
	}
	cout << "\tDialog located." << endl;

	// make sure the dialog's parent is our owner window.
	printHwnd("\tDialog", _hdlg);
	HWND hParent = GetParent(_hdlg);
	printHwnd("\tDialog's parent", hParent);
	ASSERT(hParent == _hwnd);
	return true;
}

bool UITestWorker_InputBox::runTest()
{
	cout << "\t====== INPUTBOX TEST CASE (" << _params.caseLabel << ") ======" << endl;
	if (!UITestWorker::runTest())
		return false;

	// if we're asked for it, find the edit control and programmatically enter preselected text.
	if (_params.inputId)
	{
		HWND hInput = GetDlgItem(_hdlg, _params.inputId);
		if (!hInput)
		{
			cout << "\tInput field not found." << endl;
			return false;
		}
		cout << "\tEntering preset text in input field." << endl;
		SetWindowText(hInput, _params.inputText);
		Sleep(1000);
	}
	// programmatically set a check mark on the check box.
	if (_params.checkboxId)
	{
		HWND hCheckbox = GetDlgItem(_hdlg, _params.checkboxId);
		if (!hCheckbox)
		{
			cout << "\tCheck box not found." << endl;
			return false;
		}
		cout << "\t" << (_params.checked? "Checking":"Unchecking") << " the check box" << endl;
		CheckDlgButton(_hdlg, _params.checkboxId, _params.checked);
		Sleep(1000);
	}
	// locate a button our client wants us to press.
	HWND hButton = GetDlgItem(_hdlg, _params.buttonId);
	if (!hButton)
	{
		cout << "\tDialog button not found." << endl;
		return false;
	}
	// click it programmatically.
	cout << "\tPressing the button to close the dialog." << endl;
	PostMessage(hButton, BM_CLICK, 0, 0);
	Sleep(1000);

	cout << "\tInputBox test case completed." << endl;
	return true;
}


bool UITestWorker_ProgressBox::runTest()
{
	cout << "\t====== PROGRESSBOX TEST CASE (" << _params.caseLabel << ") ======" << endl;
	if (!UITestWorker::runTest())
		return false;

	Sleep(1000);

	// locate a button our client wants us to press.
	HWND hButton = GetDlgItem(_hdlg, _params.buttonId);
	if (!hButton)
	{
		cout << "\tDialog button not found." << endl;
		return false;
	}
	Sleep(_params.delayInMilliseconds);
	// click it programmatically.
	cout << "\tPressing the button to close the dialog." << endl;
	PostMessage(hButton, BM_CLICK, 0, 0);
	_params.tckClicked = GetTickCount();
	Sleep(100);

	cout << "\tProgressBox test case completed." << endl;
	return true;
}

STDMETHODIMP UITestWorker_ProgressBox::Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS FAR* pdispparams, VARIANT FAR* pvarResult, EXCEPINFO FAR* pexcepinfo, UINT FAR* puArgErr)
{
	if (dispidMember == PROGRESSBOXEVENTS_DISPID_CANCEL)
	{
		// void Cancel([in, out] VARIANT_BOOL* CanCancel);
		ASSERT(pdispparams->cArgs == 1 && pdispparams->rgvarg[0].vt == (VT_BOOL|VT_BYREF));
		/* if we want to disallow cancelation, do this.
		pdispparams->rgvarg[0].pboolVal = VARIANT_FALSE;
		*/
		_params.cancelFired++;
		_params.tckCanceled = GetTickCount();
		return NOERROR;
	}
	return DISP_E_MEMBERNOTFOUND;
}

