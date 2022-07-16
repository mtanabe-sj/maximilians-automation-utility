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
/*
TestUtil is a console program for programmatically testing the automation interfaces of MaxsUtil.dll. 
A test is a condition argument passed to ASSERTX(). ASSERTX() stops the particular interface test if the test fails.

I. Testing VersionInfo
The program uses its own built-in version resource to do the job. While MaxsUtil loads TestUtil's resource, TestUtil sees if attribute values it asks for equal the values it independently gets from Win32.
TestUtil conducts these operations.
1) get our own exe path by calling Win32 GetModuleFileName with a null module handle. The test uses the version resource in the exe.
2) create a VersionInfo instance, and assign the exe to VersionInfo by passing the pathname.
4) test property VersionInfo.VersionString which corresponds to the FileVersion member of the file's Version resource. compare VersionString with the correct value we know.
5) test querying of a FixedFileInfo using method VersionInfo.QueryAttribute. call the method for "ProductVersion" which is a fixed-length attribute from the FixedFileInfo block of Win32 Version Info resource. compare the value with the correct value we know.
6) Test properties VersionInfo.Language and .CodePage. read and compare them to the right values we know. language is 1033 (0x409) meaning english. codepage is 1200 (0x40b), or little-endian unicode.
7) test VersionInfo.QueryAttribute for a variable-length attribute from the StringFileInfo block of Win32 Version Info. so, ask for "FileDescription", and compare it to the right english value we know.
7) test the multi-language version query by iterating through available languages and verifying variable-length version attributes for each language. to walk the languages, call QueryTranslation repeatedly, each time incrementing an index into the translation table. the translation code from the call is a combination of language id and code page. use it to access a StringFileInfo block that belongs to the translation language. use QueryAttribute to read the language-dependent product name and company name. compare them to the right values we know.
8) finally, test the IObjectSafety interface that VersionInfo inherits. QI VersionInfo for an IObjectSafety. use the latter to retrieve security settings. they must match the known correct values.

II. Testing InputBox
1) Create an InputBox instance Test for persistence of the caption text by assigning a value to the Caption property and reading it back and comparing the assigned and read text. Note that uniqueness in the caption text is necessary because a subsequent UI test tries to locate the InputBox dialog by searching for a window of the unique caption in the entire pool of windows currently open on the desktop. Note that UITestWorker will start a worker thread to do the caption search. Once it finds the dialog, the worker will programmatically enter preselected text and click the OK button. Class UITestWorker performs the automated UI test.
2) Test the Note, InitialValue, and CheckBox properties for persistence of the assigned value, too.
3) Start a UITestWorker on the InputBox. Start the InputBox dialog by calling method Show(). UITestWorker programmatically places predefined test text in the input field of the InputBox. It closes the InputBox by programmatically clicking the OK button. Read the InputValue and compare the reade value with the value planted by UITestWorker. If they match exactly, the test succeeds.
4) Prepare for another automated UI test. This time, the Cancel button is tested. Generate new unique caption text for a new test case with UITestWorker.
5) Restart the UITestWorker for the second time. Assign the new caption to the InputBox, and start another InputBox dialog. In a short while, UITestWorker programmatically clicks the Cancel button.
6) Check the return value from the Show() call. If the InputBox dialog is canceled, the return value is -1.

III. Testing ProgressBox
1) Create a ProgressBox instance. Assign and read back the Caption, Message and Note properties to test value persistence.
2) Test value persistence on LowerBound, UpperBound and ProgressPos for a range of values.
3) Next, test the progress bar's functionality. Define a progress range with LowerBound and UpperBound. Start the ProgressBox dialog, and enter a loop. In each iteration, increment the progress position. Exit the loop on reaching the upper bound. Also, at each step, generate a note indicating the current step position within the range. Check for an unexpected Cancel event.
4) After the iteration completes, stop the ProgressBox and read the ProgressPos. The test is a success if it has not been canceled, and if the read progress position equals the last assigned position value.
5) Next, test resuse of a stopped ProgressBox. The current ProgressBox instance will be reused. It's just been stopped. To restart the progress display after it's stopped, make a new assignment to the Caption property. That forces ProgressBox to start a new progress window. The test succeeds if the return value is a success code (S_OK). If the test fails, any subsequent property assignment raises an interface error.
6) Next, test the Move method and the Append-to-Note mode. Tell ProgressBox to move the progress window to the lower right corner of the screen. Then, start the progress dialog requesting that the Note control is put in Append mode for continuous feeding of text into the Note edit control.
7) Read the new position of the progress window. The coordinates should map to the lower right corner.
8) Enter a loop, incrementing index i from 0 to 100 at step 1. Generate an interation-specific text string and append it to Note. Update ProgressPos to i. Check the Canceled state.
9) After the loop is exited, read the latest cummulative text from Note. The text is the entire stack of lines of text appended to Note. Count the number of linefeeds in it. That should equal the progress range, if the Append-to-Note test was successful. Call the Stop method to kill the progress display.
10) Create a new instance of ProgressBox for the next test.
11) Test the Cancel button using UITestWorker. When progress reaches halfway, UITestWorker programmatically click the Cancel button. Configure the ProgressBox with a range of 0 to 100. Start the UITestWorker. Loop through the range stepping the progress position. Watch out for a Cancel event.
12) The test is successful if the loop terminated because of a cancelation, and if the two events, the programmatic clicking of the Cancel button and the detection of a true Canceled property, occurred simultaneously or closely together (0 to 500ms).
*/

#include "pch.h"
#include "MaxsUtil_i.c"
#include "bstring.h"
#include "VariantAutoRel.h"
#include "appver.h"
#include "UITestWorker.h"
#include "..\MaxsUtil\resource.h"


using namespace std;


#define ASSERTX(c) {ASSERT(c); if(!(c)) {goto _assertionFailed;} else {} }

wstring _toLangTag(DWORD langCode)
{
	WCHAR buf[40];
	wsprintf(buf, L"%04X%04X", LOWORD(langCode), HIWORD(langCode));
	return buf;
}

string hresultToString(HRESULT hr)
{
	if (hr == HRESULT_FROM_WIN32(ERROR_RESOURCE_DATA_NOT_FOUND))
		return "ERROR_RESOURCE_DATA_NOT_FOUND";
	if (hr == HRESULT_FROM_WIN32(ERROR_RESOURCE_TYPE_NOT_FOUND))
		return "ERROR_RESOURCE_TYPE_NOT_FOUND";
	if (hr == HRESULT_FROM_WIN32(ERROR_RESOURCE_NAME_NOT_FOUND))
		return "ERROR_RESOURCE_NAME_NOT_FOUND";
	if (hr == HRESULT_FROM_WIN32(ERROR_RESOURCE_LANG_NOT_FOUND))
		return "ERROR_RESOURCE_LANG_NOT_FOUND";
	//
	if (hr == NOERROR)
		return "CANCELED_OR_ABORTED";
	return "error " + to_string(hr);
}

HRESULT testVersionInfo()
{
	cout << "********** VERSIONINFO TESTS **********" << endl;

	// get a pathname of this module. we will use our own version resource to run the test.
	WCHAR fpath[MAX_PATH];
	GetModuleFileName(NULL, fpath, ARRAYSIZE(fpath));
	wcout << L"TEST FILE: " << fpath << endl;

	HRESULT hr;
	bstring fpathVerify;
	bstring fversion;
	VariantAutoRel fdesc;
	VariantAutoRel nextVer;
	short langId, codepage;
	short j;
	IObjectSafety *os;

	cout << "Creating VersionInfo" << endl;
	IVersionInfo *vi;
	// create a VersionInfo instance and get the automation interface.
	hr = CoCreateInstance(CLSID_VersionInfo, NULL, CLSCTX_INPROC_SERVER, IID_IVersionInfo, (LPVOID*)&vi);
	ASSERTX(hr == S_OK);
	// VersionInfo creation succeeded.
	cout << " RESULT --> PASS" << endl;

	// assign our path to VersionInfo.
	cout << "Testing File assignment" << endl;
	hr = vi->put_File(bstring(fpath));
	ASSERTX(hr == S_OK);
	// verify the assigned pathname.
	hr = vi->get_File(&fpathVerify);
	ASSERTX(_wcsicmp(fpath, fpathVerify) == 0);
	cout << " RESULT --> PASS" << endl;

	cout << "Testing VersionString get" << endl;
	hr = vi->get_VersionString(&fversion);
	ASSERT(hr == S_OK);
	wcout << "[VersionString=" << fversion._b << L"]" << endl;
	// the queried version must match the correct value we know.
	ASSERTX(wcscmp(fversion, TESTAPP_FILEVERSION) == 0);
	cout << " RESULT --> PASS" << endl;

	cout << "Testing QueryAttribute for ProductVersion" << endl;
	// test QueryAttribute() for a FixedFileInfo attribute.
	hr = vi->QueryAttribute(bstring(L"ProductVersion"), &nextVer._v);
	ASSERTX(hr == S_OK);
	wcout << "[ProductVersion=" << nextVer._v.bstrVal << L"]" << endl;
	ASSERTX(wcscmp(nextVer._v.bstrVal, TESTAPP_PRODUCTVERSION) == 0);
	cout << " RESULT --> PASS" << endl;

	cout << "Testing QueryAttribute for FileDescription for default language" << endl;
	hr = vi->QueryAttribute(bstring(L"FileDescription"), fdesc);
	ASSERTX(hr == S_OK);
	wcout << L" [FileDescription=" << fdesc._v.bstrVal << L"]" << endl;
	ASSERTX(wcsncmp(fdesc._v.bstrVal, TESTAPP_DESC_PREFIX, TESTAPP_DESC_PREFIX_LEN) == 0);
	hr = vi->get_Language(&langId);
	hr = vi->get_CodePage(&codepage);
	cout << " [Language=" << langId << ", CodePage=" << codepage << "]" << endl;
	ASSERTX(langId == 0x409 && codepage == 0x4b0);
	cout << " RESULT --> PASS" << endl;

	// enumerate the languages in the version info's translation table.
	// usually there is just one, e.g., 040904b0 for us english.
	cout << "Testing Translation-to-StringFileInfo mapping" << endl;
	for (j = 1; j <= 8; j++)
	{
		VariantAutoRel langCode;
		hr = vi->QueryTranslation(j, langCode);
		ASSERTX(hr == S_OK || hr == E_BOUNDS);
		if (hr != S_OK)
			break;
		cout << " [Translation " << j << ": Language=" << LOWORD(langCode._v.ulVal) << ", CodePage=" << HIWORD(langCode._v.ulVal) << "]" << endl;
		ASSERTX(langCode._v.vt == VT_UI4);
		vi->put_Language(LOWORD(langCode._v.ulVal));
		vi->put_CodePage(HIWORD(langCode._v.ulVal));
		// test QueryAttribute() for a StringFileInfo attribute.
		VariantAutoRel product;
		hr = vi->QueryAttribute(bstring(L"ProductName"), product);
		wcout << L" [ProductName=" << product._v.bstrVal << L"]" << endl;
		ASSERTX(wcsncmp(product._v.bstrVal, TESTAPP_PRODUCTNAME_PREFIX, TESTAPP_PRODUCTNAME_PREFIX_LEN) == 0);
		VariantAutoRel company;
		hr = vi->QueryAttribute(bstring(L"CompanyName"), company);
		wcout << L" [CompanyName=" << company._v.bstrVal << L"]" << endl;
		ASSERTX(wcsncmp(company._v.bstrVal, TESTAPP_COMPANY_PREFIX, TESTAPP_COMPANY_PREFIX_LEN) == 0);
	}
	ASSERTX(j == 4); // there are 3 translations.
	cout << " RESULT --> PASS" << endl;

	// test the IObjectSafety interface.
	cout << "Testing IObjectSafety" << endl;
	hr = vi->QueryInterface(IID_IObjectSafety, (LPVOID*)&os);
	ASSERTX(hr == S_OK);
	if (hr == S_OK)
	{
		os->AddRef();
		DWORD supported = 0, enabled = 0;
		hr = os->GetInterfaceSafetyOptions(IID_IDispatch, &supported, &enabled);
		cout << " [GetInterfaceSafetyOptions: SupportedOptions=" << supported << ", EnabledOptions=" << enabled << "]" << endl;
		ASSERTX(hr == S_OK);
		IUnknown *punk;
		hr = os->QueryInterface(IID_IUnknown, (LPVOID*)&punk);
		ASSERTX(hr == S_OK);
		if (punk)
			punk->Release();
		os->Release();
	}
	cout << " RESULT --> PASS" << endl;

	vi->Release();

	cout << "PASSED ALL VERSIONINFO TESTS" << endl;
	return S_OK;
_assertionFailed:
	cout << " TEST FAILED: (" << hresultToString(hr) << ")" << endl;
	return E_FAIL;
}


HRESULT testInputBox()
{
	cout << "********** INPUTBOX TESTS **********" << endl;
	cout << "Creating InputBox" << endl;

	HRESULT hr;
	bstring caption1;
	bstring caption2;
	bstring note1(L"Note is shown under input field.");
	bstring note2;
	bstring chkbox1(L"Test checkbox option.");
	bstring chkbox2;
	bstring initval1(L"This is a test value InitialValue has set.");
	bstring initval2;
	bstring testVal;
	bstring inputval;
	long len;
	VARIANT_BOOL checked = 0;

	HWND hwnd = GetConsoleWindow();
	UITestWorker_InputBox tester(hwnd);

	IInputBox *inbox;
	// create a ProgressBox instance and get the automation interface.
	hr = CoCreateInstance(CLSID_InputBox, NULL, CLSCTX_INPROC_SERVER, IID_IInputBox, (LPVOID*)&inbox);
	ASSERTX(hr == S_OK);
	// InputBox creation succeeded.
	cout << " RESULT --> PASS" << endl;

	cout << "Testing Caption put/get" << endl;
	caption1.format(L"$$TestUtil:%d:%d:InputBox:OK$$", GetCurrentProcessId(), GetCurrentThreadId());
	hr = inbox->put_Caption(caption1);
	ASSERTX(hr == S_OK);
	hr = inbox->get_Caption(&caption2);
	ASSERTX(hr == S_OK);
	ASSERTX(wcscmp(caption2, caption1) == 0);
	cout << " RESULT --> PASS" << endl;

	cout << "Testing Note put/get" << endl;
	hr = inbox->put_Note(note1);
	ASSERTX(hr == S_OK);
	hr = inbox->get_Note(&note2);
	ASSERTX(hr == S_OK);
	ASSERTX(wcscmp(note2, note1) == 0);
	cout << " RESULT --> PASS" << endl;

	cout << "Testing CheckBox put/get" << endl;
	hr = inbox->put_CheckBox(chkbox1);
	ASSERTX(hr == S_OK);
	hr = inbox->get_CheckBox(&chkbox2);
	ASSERTX(hr == S_OK);
	ASSERTX(wcscmp(chkbox2, chkbox1) == 0);
	cout << " RESULT --> PASS" << endl;

	cout << "Testing InitialValue put/get" << endl;
	hr = inbox->put_InitialValue(initval1);
	ASSERTX(hr == S_OK);
	hr = inbox->get_InitialValue(&initval2);
	ASSERTX(hr == S_OK);
	ASSERTX(wcscmp(initval2, initval2) == 0);
	cout << " RESULT --> PASS" << endl;

	testVal.attach(initval1);
	_wcsupr(testVal._b); // uppercase the text
	hr = tester.StartTest("OK UNCHECKED", caption1, IDOK, IDC_EDIT_INPUTVALUE, testVal);
	ASSERTX(hr == S_OK);
	cout << "Testing Show --> OK | UNCHECKED" << endl;
	hr = inbox->Show(bstring(L"$$TestUtil~Message$$"), VariantAutoRel(hwnd), &len);
	ASSERTX(hr == S_OK);
	hr = inbox->get_InputValue(&inputval);
	ASSERTX(hr == S_OK);
	if (len >= 0)
	{
		ASSERTX(wcscmp(inputval, testVal) == 0);
		cout << " RESULT --> PASS" << endl;
	}
	tester.StopTest();

	_wcslwr(testVal._b); // put the text in lowercase this time
	hr = tester.StartTest("OK CHECKED", caption1, IDOK, IDC_EDIT_INPUTVALUE, testVal, IDC_CHECK1, TRUE);
	ASSERTX(hr == S_OK);
	cout << "Testing Show --> OK | CHECKED" << endl;
	hr = inbox->Show(bstring(L"$$TestUtil~Message$$"), VariantAutoRel(hwnd), &len);
	ASSERTX(hr == S_OK);
	inputval.free();
	hr = inbox->get_InputValue(&inputval);
	ASSERTX(hr == S_OK);
	hr = inbox->get_Checked(&checked);
	ASSERTX(hr == S_OK && checked == VARIANT_TRUE);
	if (len >= 0 && checked)
	{
		ASSERTX(wcscmp(inputval, testVal) == 0);
		cout << " RESULT --> PASS" << endl;
	}
	tester.StopTest();

	cout << "Testing Show --> Cancel" << endl;
	caption2.format(L"$$TestUtil:%d:%d:InputBox:Cancel$$", GetCurrentProcessId(), GetCurrentThreadId());
	hr = tester.StartTest("CANCEL", caption2, IDCANCEL);
	ASSERTX(hr == S_OK);
	hr = inbox->put_Caption(caption2);
	ASSERTX(hr == S_OK);
	hr = inbox->Show(bstring(L"$$TestUtil~Message$$"), VariantAutoRel(hwnd), &len);
	ASSERTX(hr == S_OK);
	if (len < 0)
	{
		cout << " RESULT --> PASS" << endl;
	}
	tester.StopTest();

	inbox->Release();

	cout << "PASSED ALL INPUTBOX TESTS" << endl;
	return S_OK;
_assertionFailed:
	cout << " TEST FAILED: (" << hresultToString(hr) << ")" << endl;
	return E_FAIL;
}

HRESULT testProgressBox()
{
	HRESULT hr;
	long i, dt, val1, val2, pos;
	VARIANT_BOOL canceled = VARIANT_FALSE;
	IProgressBox *progbox;

	cout << "********** PROGRESSBOX TESTS **********" << endl;
	cout << "Creating ProgressBox" << endl;
	// create a ProgressBox instance and get the automation interface.
	hr = CoCreateInstance(CLSID_ProgressBox, NULL, CLSCTX_INPROC_SERVER, IID_IProgressBox, (LPVOID*)&progbox);
	ASSERTX(hr == S_OK);
	// ProgressBox creation succeeded.
	cout << " RESULT --> PASS" << endl;

	{
		bstring caption1;
		bstring caption2;
		bstring msg1(L"Test message.");
		bstring msg2;
		bstring note1(L"Notes are shown below Message.");
		bstring note2;

		caption1.format(L"$$TestUtil:%d:%d:ProgressBox$$", GetCurrentProcessId(), GetCurrentThreadId());

		// test text assignment with Caption, Message and Note.
		cout << "Testing Caption put/get" << endl;
		hr = progbox->put_Caption(caption1);
		ASSERTX(hr == S_OK);
		hr = progbox->get_Caption(&caption2);
		ASSERTX(hr == S_OK);
		ASSERTX(wcscmp(caption2, caption1) == 0);
		cout << " RESULT --> PASS" << endl;

		cout << "Testing Message put/get" << endl;
		hr = progbox->put_Message(msg1);
		ASSERTX(hr == S_OK);
		hr = progbox->get_Message(&msg2);
		ASSERTX(hr == S_OK);
		ASSERTX(wcscmp(msg2, msg1) == 0);
		cout << " RESULT --> PASS" << endl;

		cout << "Testing Note put/get" << endl;
		hr = progbox->put_Note(note1);
		ASSERTX(hr == S_OK);
		hr = progbox->get_Note(&note2);
		ASSERTX(hr == S_OK);
		ASSERTX(wcscmp(note2, note1) == 0);
		cout << " RESULT --> PASS" << endl;

		cout << "Testing ProgressPos, LowerBound and UpperBound put/get" << endl;
		for (val1 = 0; val1 < 0x1000; val1++)
		{
			hr = progbox->put_LowerBound(val1);
			ASSERTX(hr == S_OK);
			hr = progbox->get_LowerBound(&val2);
			ASSERTX(hr == S_OK);
			ASSERTX(val2 == val1);
		}
		for (val1 = 0; val1 < 0x1000; val1++)
		{
			hr = progbox->put_UpperBound(val1);
			ASSERTX(hr == S_OK);
			hr = progbox->get_UpperBound(&val2);
			ASSERTX(hr == S_OK);
			ASSERTX(val2 == val1);
		}
		for (val1 = 0; val1 < 0x1000; val1++)
		{
			hr = progbox->put_ProgressPos(val1);
			ASSERTX(hr == S_OK);
			hr = progbox->get_ProgressPos(&val2);
			ASSERTX(hr == S_OK);
			ASSERTX(val2 == val1);
		}
		cout << " RESULT --> PASS" << endl;
	}

	{
		bstring msg(L"Testing the progress bar. Do not touch the Cancel button. ProgressBox will close automatically when progress reaches completion.");

		cout << "Preparing ProgressBox for Completion Test" << endl;

		hr = progbox->put_Message(msg);
		ASSERTX(hr == S_OK);

		dt = 100;
		val1 = 0;
		val2 = 100;
		hr = progbox->put_LowerBound(val1);
		ASSERTX(hr == S_OK);
		hr = progbox->put_UpperBound(val2);
		ASSERTX(hr == S_OK);

		cout << "Testing ProgressBar Completion" << endl;
		hr = progbox->Start(NULL, NULL);
		ASSERTX(hr == S_OK);

		for (i = val1; i <= val2; i++)
		{
			Sleep(dt);
			bstring note;
			note.format(L"Test Note - Step %d of %d", i, val2);
			hr = progbox->put_Note(note);
			ASSERTX(hr == S_OK);
			hr = progbox->put_ProgressPos(i);
			ASSERTX(hr == S_OK);
			hr = progbox->get_Canceled(&canceled);
			ASSERTX(hr == S_OK);
			if (canceled)
			{
				cout << " Cancel detected at step " << i << "." << endl;
				break;
			}
		}
		ASSERTX(!canceled);

		Sleep(1000);

		hr = progbox->Stop();
		ASSERTX(hr == S_OK);
		cout << " RESULT --> PASS" << endl;

		cout << "Testing Access to ProgressPos after Stop" << endl;
		hr = progbox->get_ProgressPos(&pos);
		ASSERTX(hr == S_OK && pos == val2);
		cout << " RESULT --> PASS" << endl;

		cout << "Testing Unexpected Restart" << endl;
		// you may not restart a stopped dialog.
		hr = progbox->Start(NULL, NULL);
		ASSERTX(hr == E_UNEXPECTED);
		cout << " RESULT --> PASS" << endl;
	}

	{
		int n;
		HWND hprog = NULL;
		RECT rcBox;
		bstring caption;
		bstring note;
		bstring noteCummulative;
		bstring msg(L"Testing the Append-to-Note option. Do not touch the Cancel button. ProgressBox will be moved to lower right hand corner. It will close automatically when progress reaches completion.");

		caption.format(L"$$TestUtil:%d:%d:ProgressBox:Move$$", GetCurrentProcessId(), GetCurrentThreadId());

		cout << "Testing Reuse of Stopped ProgressBox" << endl;

		// by assigning a caption, one can force a stopped IProgressBox to re-establish its UI.
		hr = progbox->put_Caption(caption);
		ASSERTX(hr == S_OK);
		cout << " RESULT --> PASS" << endl;

		cout << "Preparing ProgressBox for Append-to-Note Test" << endl;
		hr = progbox->put_Message(msg);
		ASSERTX(hr == S_OK);

		canceled = VARIANT_FALSE;
		dt = 100;
		val1 = 0;
		val2 = 100;
		hr = progbox->put_LowerBound(val1);
		ASSERTX(hr == S_OK);
		hr = progbox->put_UpperBound(val2);
		ASSERTX(hr == S_OK);

		cout << "Testing Append-to-Note Option and Move" << endl;
		hr = progbox->put_Note(NULL); // clear the Note field.
		ASSERTX(hr == S_OK);
		hr = progbox->Move(VariantAutoRel((long)PROGRESSBOXMOVEFLAG_RIGHTBOTTOM), NULL, NULL);
		hr = progbox->Start(VariantAutoRel((long)PROGRESSBOXSTARTOPTION_APPEND_TO_NOTE), NULL);
		ASSERTX(hr == S_OK);

		hr = progbox->get_WindowHandle((OLE_HANDLE*)&hprog);
		ASSERTX(hr == S_OK);
		ASSERTX(IsWindowVisible(hprog));
		GetWindowRect(hprog, &rcBox);
		cout << " ProgressBox moved to: Right=" << rcBox.right << ", Bottom=" << rcBox.bottom << endl;
		RECT rcDT;
		GetWindowRect(GetDesktopWindow(), &rcDT);
		ASSERTX(rcDT.right == rcBox.right && rcDT.bottom == rcBox.bottom);

		for (i = val1; i <= val2; i++)
		{
			Sleep(dt);
			// append a new line of text to Note.
			note.format(L"Test Note - Step %d of %d", i, val2);
			hr = progbox->put_Note(note);
			ASSERTX(hr == S_OK);
			hr = progbox->put_ProgressPos(i);
			ASSERTX(hr == S_OK);
			hr = progbox->get_Canceled(&canceled);
			ASSERTX(hr == S_OK);
			if (canceled)
			{
				cout << " Cancel detected at step " << i << "." << endl;
				break;
			}
		}
		ASSERTX(!canceled);

		hr = progbox->get_Note(&noteCummulative);
		ASSERTX(hr == S_OK);
		n = noteCummulative.countChar('\n');
		cout << " Lines appended: " << n << endl;
		ASSERTX(n == val2 + 1);

		Sleep(1000);

		hr = progbox->Stop();
		ASSERTX(hr == S_OK);
		cout << " RESULT --> PASS" << endl;
	}

	progbox->Release();

	// again create a ProgressBox instance.
	hr = CoCreateInstance(CLSID_ProgressBox, NULL, CLSCTX_INPROC_SERVER, IID_IProgressBox, (LPVOID*)&progbox);
	ASSERTX(hr == S_OK);
	cout << "Testing ProgressBox Cancel." << endl;

	{
		bstring caption;
		bstring note;
		bstring msg(L"Testing the Cancel button. Do not touch the button. The progress box should cancel automatically when progress indication passes the midpoint.");

		caption.format(L"$$TestUtil:%d:%d:ProgressBox:Cancel$$", GetCurrentProcessId(), GetCurrentThreadId());

		UITestWorker_ProgressBox tester(progbox);

		// programmatically press Cancel when progress reaches midway.
		hr = tester.StartTest("CANCEL", caption, IDCANCEL, MulDiv(val2,dt,2));
		ASSERTX(hr == S_OK);
		hr = progbox->put_Caption(caption);
		ASSERTX(hr == S_OK);
		hr = progbox->put_Message(msg);
		ASSERTX(hr == S_OK);

		canceled = VARIANT_FALSE;
		dt = 100;
		val1 = 0;
		val2 = 1000;
		hr = progbox->put_LowerBound(val1);
		ASSERTX(hr == S_OK);
		hr = progbox->put_UpperBound(val2);
		ASSERTX(hr == S_OK);

		hr = progbox->Start(NULL, NULL);
		ASSERTX(hr == S_OK);

		for (i = val1; i <= val2; i++)
		{
			Sleep(dt);
			note.format(L"Test Note - Step %d of %d", i, val2);
			hr = progbox->put_Note(note);
			ASSERTX(hr == S_OK);
			hr = progbox->put_ProgressPos(i);
			ASSERTX(hr == S_OK);
			hr = progbox->get_Canceled(&canceled);
			ASSERTX(hr == S_OK);
			if (canceled)
			{
				ASSERTX(tester.IsProgressBoxCanceled());
				cout << " Cancel detected at step " << i << "." << endl;
				break;
			}
		}
		ASSERTX(canceled);
		dt = tester.ElapsedTimeSinceCancel();
		ASSERTX(0 <= dt && dt < 500);
		cout << " Cancel notification came in with a delay of " << dt << " ms." << endl;

		hr = progbox->Stop();
		ASSERTX(hr == S_OK);
		cout << " RESULT --> PASS" << endl;

		tester.StopTest();
	}

	progbox->Release();

	cout << "PASSED ALL PROGRESSBOX TESTS" << endl;
	return S_OK;
_assertionFailed:
	cout << " TEST FAILED: (" << hresultToString(hr) << ")" << endl;
	return E_FAIL;
}

int main(int argc, char **argv)
{
	HRESULT hr, hr1, hr2, hr3;
	if (SUCCEEDED(hr = CoInitialize(NULL)))
	{
		cout << "This program tests COM automation interfaces of MaxsUtil." << endl;
		cout << "The test takes a few minutes to complete. It runs without" << endl;
		cout << "need for input from user. However, it starts and closes" << endl;
		cout << "dialogs in various locations on desktop. So, close all apps," << endl;
		cout << "and do not interfere until the test completes." << endl;
		cout << endl;
		cout << "Press ENTER to start the test." << endl;

		if (cin.get() == '\n')
		{
			hr1 = testVersionInfo();
			hr2 = testInputBox();
			hr3 = testProgressBox();

			if (hr1 == S_OK && hr2 == S_OK && hr3 == S_OK)
				cout << ">>> ALL COCLASSES PASSED TESTS SUCCESSFULLY <<<";
			else
				cout << ">>> ONE OR MORE COCLASSES FAILED TO PASS A TEST <<<";
		}

		CoUninitialize();
	}
	return hr;
}
