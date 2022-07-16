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


// implements the IVersionInfo interface of the VersionInfo coclass.
class VersionInfoImpl :
	public IDispatchWithObjectSafetyImpl<IVersionInfo, &IID_IVersionInfo, &LIBID_MaxsUtilLib>
{
public:
	VersionInfoImpl() : _langId(0), _codepage(0) {}
	~VersionInfoImpl() {}

	// IUnknown methods
	DELEGATE_IUNKNOWN_TO_IDISPATCHWITHOBJECTSAFETYIMPL(IVersionInfo, &IID_IVersionInfo, &LIBID_MaxsUtilLib)

	// IVersionInfo methods
	STDMETHOD(get_File)(/* [retval][out] */ BSTR *Value);
	STDMETHOD(put_File)(/* [in] */ BSTR NewValue);
	STDMETHOD(get_MajorVersion)(/* [retval][out] */ long *Value);
	STDMETHOD(get_MinorVersion)(/* [retval][out] */ long *Value);
	STDMETHOD(get_VersionString)(/* [retval][out] */ BSTR *Value);
	STDMETHOD(QueryAttribute)(/* [in] */ BSTR Name, /* [retval][out] */ VARIANT *Value);
	STDMETHOD(get_Language)(/* [retval][out] */ short *Value);
	STDMETHOD(put_Language)(/* [in] */ short NewValue);
	STDMETHOD(get_CodePage)(/* [retval][out] */ short *Value);
	STDMETHOD(put_CodePage)(/* [in] */ short NewValue);
	STDMETHOD(QueryTranslation)(/* [in] */ short TranslationIndex, /* [retval][out] */ VARIANT *LangCode);

protected:
	bstring _file; // pathname of a file with a version resource.
	bstring _vi; // a version resource structure from win32 GetFileVersionInfo().
	short _langId; // langauge (e.g., 1033 for english)
	short _codepage; // codepage (e.g., 1200 for unicode)

	DWORD queryLangCp(VARIANT *langIndex);
	HRESULT queryVersionNumber(long *Value);
	HRESULT queryAttribData(LPCWSTR subblockPath, LPVOID* attribData, UINT *attribLen);
	HRESULT queryVersionInfo(bstring& strVerInfo);
	HRESULT ensureLangCp();
};

