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
#include "VersionInfoImpl.h"


/* fixed-length integer attributes available from the FixedFileInfo block of a version info resource */
enum VI_FIXEDFILEATTRIBUTE {
	VIFFA_Unknown,
	VIFFA_FileVersion,
	VIFFA_ProductVersion,
	VIFFA_Signature,
	VIFFA_StrucVersion,
	VIFFA_FileVersionMS,
	VIFFA_FileVersionLS,
	VIFFA_ProductVersionMS,
	VIFFA_ProductVersionLS,
	VIFFA_FileFlagsMask,
	VIFFA_FileFlags,
	VIFFA_FileOS,
	VIFFA_FileType,
	VIFFA_FileSubtype,
	VIFFA_FileDate,
	VIFFA_Max
};

LPCWSTR VIFixedFileAttrbs[] = {
	L"FileVersion",
	L"ProductVersion",
	L"Signature",
	L"StrucVersion",
	L"FileVersionMS",
	L"FileVersionLS",
	L"ProductVersionMS",
	L"ProductVersionLS",
	L"FileFlagsMask",
	L"FileFlags",
	L"FileOS",
	L"FileType",
	L"FileSubtype",
	L"FileDate",
};

/* variable-length string attributes stored in a StringFileInfo block */
enum VI_STRINGFILEATTRIBUTE {
	VISFA_Unknown,
	VISFA_Comments,
	VISFA_CompanyName,
	VISFA_FileDescription,
	VISFA_FileVersion,
	VISFA_InternalName,
	VISFA_LegalCopyright,
	VISFA_LegalTrademarks,
	VISFA_OriginalFilename,
	VISFA_ProductName,
	VISFA_ProductVersion,
	VISFA_PrivateBuild,
	VISFA_SpecialBuild,
};

LPCWSTR VIStringFileAttribs[] = {
	L"Comments",
	L"CompanyName",
	L"FileDescription",
	L"FileVersion",
	L"InternalName",
	L"LegalCopyright",
	L"LegalTrademarks",
	L"OriginalFilename",
	L"ProductName",
	L"ProductVersion",
	L"PrivateBuild",
	L"SpecialBuild",
};


/* get_File - [propget] returns a pathname identifying a file for which version info is queried.

Parameters:
Value - [retval][out] contains a pathname of a current file. version queries are made with the version resource from this file. The pathname is the one previously set by a call to put_File. an interface error of E_UNEXPECTED is returned if no pathname is available.
*/
STDMETHODIMP VersionInfoImpl::get_File(/* [retval][out] */ BSTR *Value)
{
	if (_file.length() == 0)
		return E_UNEXPECTED;
	bstring v(_file);
	*Value = v.detach();
	return S_OK;
}

/* put_File - [propput] accepts a pathname to a file for which vesion info is queried.

Parameters:
NewValue - [in] a pathname of a file. Version queries will be made against a version resource from this file.
*/
STDMETHODIMP VersionInfoImpl::put_File(/* [in] */ BSTR NewValue)
{
	_file.assignW(NewValue);
	/* a new path is assigned. it's time to clear cached version info structure and language settings associated with the previous file. the resetting is necessary because it prevents the obsolete version data from charading as the new file's. it's important because one can use a VersionInfo instance on one file now and re-assign it to another file later. */
	_vi.free();
	_langId = _codepage = 0;
	return S_OK;
}

/* get_Language - [propget] returns a current language id. if no language has been explicitly set, the method reads the first language entry from the version info's translation block, and reports it as the current language.

Parameters:
Value - [retval][out] contains a current language id. if no language has been set in a prior call to put_Language, returns a default language id which is the first entry in the translation block.
*/
STDMETHODIMP VersionInfoImpl::get_Language(/* [retval][out] */ short *Value)
{
	if (_file.length() == 0)
		return E_UNEXPECTED;
	// if _langId is undefined, revert to the first translation entry.
	ensureLangCp();
	*Value = _langId;
	return S_OK;
}

/* put_Language - [propput] sets a current language id. a language id is used together with a codepage setting in forming a fully qualified path to a string attribute. See QueryAttribute() on how the language and codepage settings are used.

Parameters:
NewValue - [in] a new language id to use in making string attribute queries. some files contain more than one StringFileInfo block differentiated by language. So, if you want a string attribute in a language different from the default language, you use the Language property and assign the id of a language you want to use.

Remarks:
  Examples of language id: 1033 for English. 1031 for German. 1036 for French.
*/
STDMETHODIMP VersionInfoImpl::put_Language(/* [in] */ short NewValue)
{
	_langId = NewValue;
	return S_OK;
}

/* get_CodePage - [propget] returns a current codepage. If no codepage has been explicitly set, return the one from the first entry in the translation block.

Parameters:
Value - [retval][out] contains a current codepage. if no codepage has been set in a prior call to put_CdoePage, returns a default codepage from the first entry in the translation block. Expect to get a value of 1200 which is a codepage for unicode, the only char set used in Win32 resources.
*/
STDMETHODIMP VersionInfoImpl::get_CodePage(/* [retval][out] */ short *Value)
{
	if (_file.length() == 0)
		return E_UNEXPECTED;
	ensureLangCp();
	*Value = _codepage;
	return S_OK;
}

/* put_CodePage - [propput] sets a current codepage. A valid codepage is 1200 (or 0x4b0). That's for unicode, and the only valid codepage, i would think. That's because Win32 resources always use unicode for character encoding. So, passing in anything other than 1200 doesn't make sense...

Parameters:
NewValue - [in] a codepage to use when making a query for a string version attribute. I don't think anything other than 1200 (for unicode) works...
*/
STDMETHODIMP VersionInfoImpl::put_CodePage(/* [in] */ short NewValue)
{
	_codepage = NewValue;
	return S_OK;
}

/* QueryTranslation - [method] tries to find a translation entry at a specified one-based index. If an entry exists, returns it which is a language id and codepage pair as a DWORD value. If no index is provided, returns the first translation entry. If a translation entry found consists of a language of english and a codepage for unicode, the return value is an integer value of 0x04090x0b0. The high word part of 0x409 specifies US english, while the low word part of 0x4b0 specifies the little-endian unicode codepage.

Parameters:
TranslationIndex - [in] 0 or one-based index into the version info translation block. If no entry exists at a specified index, an interface error of E_BOUNDS is returned.
LangCode - [retval][out] contains a translation entry consisting of a language id and codepage pair, positioned at the specified translation index.
*/
STDMETHODIMP VersionInfoImpl::QueryTranslation(/* [in] */ short TranslationIndex, /* [retval][out] */ VARIANT *LangCode)
{
	DWORD langCp;
	if (TranslationIndex == 0)
		langCp = queryLangCp(NULL);
	else
		langCp = queryLangCp(VariantAutoRel(TranslationIndex));
	if (langCp == 0)
		return E_BOUNDS;
	LangCode->vt = VT_UI4;
	LangCode->ulVal = langCp;
	return S_OK;
}

/* get_VersionString - [propget] returns a file version number string. Generates an interface error if the file does not have a version resource. The VersionString property does not depend on Language and CodePage. The version property is backed by the fixed-length attributes of the dwFileVersionMS and dwFileVersionLS members of the FixedFileInfo structure.

Parameters:
Value - [retval][out] contains a version string of form <major>.<minor>.<revision>.<build>, If the file has a version resource. If not, an interface error of ERROR_RESOURCE_DATA_NOT_FOUND is returned.
*/
STDMETHODIMP VersionInfoImpl::get_VersionString(/* [retval][out] */ BSTR *Value)
{
	VariantAutoRel var;
	HRESULT hr = QueryAttribute(bstring(L"FileVersion"), var);
	if (hr == S_OK)
	{
		*Value = var._v.bstrVal;
		var.init();
	}
	return hr;
}

/* get_MajorVersion - [propget] returns the major version number of the file, or an interface error if the file has no version resource.

Parameters:
Value - [retval][out] contains the major version number which is the high word part of the dwFileVersionMS member of the FixedFileInfo structure. If the file has no version resource, an interface error of ERROR_RESOURCE_DATA_NOT_FOUND is returned.
*/
STDMETHODIMP VersionInfoImpl::get_MajorVersion(/* [retval][out] */ long *Value)
{
	long v;
	HRESULT hr = queryVersionNumber(&v);
	if (hr == S_OK)
		*Value = HIWORD(v);
	return hr;
}

/* get_MajorVersion - [propget] returns the minor version number of the file, or an interface error if the file has no version resource.

Parameters:
Value - [retval][out] contains the minor version number which is the low word part of the dwFileVersionMS member of the FixedFileInfo structure. If the file has no version resource, an interface error of ERROR_RESOURCE_DATA_NOT_FOUND is returned.
*/
STDMETHODIMP VersionInfoImpl::get_MinorVersion(/* [retval][out] */ long *Value)
{
	long v;
	HRESULT hr = queryVersionNumber(&v);
	if (hr == S_OK)
		*Value = LOWORD(v);
	return hr;
}

/* _getFixedFileAttributeId - a helper to find an enumeration ordinal of VI_FIXEDFILEATTRIBUTE corresponding to an attribute of the FixedFileInfo block. For example, the function returns VIFFA_ProductVersion for an attribute name of 'ProductVersion'.

Parameters:
attribName - [in] contains an attribute name for which an VI_FIXEDFILEATTRIBUTE ordinal is searched.
*/
VI_FIXEDFILEATTRIBUTE _getFixedFileAttributeId(LPCWSTR attribName)
{
	for (int i = 0; i < ARRAYSIZE(VIFixedFileAttrbs); i++)
	{
		if (0 == _wcsicmp(attribName, VIFixedFileAttrbs[i]))
			return (VI_FIXEDFILEATTRIBUTE)(i+1);
	}
	// attribName is not a fixed-length attribute. perhaps, it's a string attribute.
	return VIFFA_Unknown;
}

/* _getStringFileAttributeId - a helper to find an ordinal corresponding to a string attribute of StringFileInfo.

Parameters:
attribName - [in] contains an attribute name for which an VI_STRINGFILEATTRIBUTE ordinal is searched.
*/
VI_STRINGFILEATTRIBUTE _getStringFileAttributeId(LPCWSTR attribName)
{
	for (int i = 0; i < ARRAYSIZE(VIStringFileAttribs); i++)
	{
		if (0 == _wcsicmp(attribName, VIStringFileAttribs[i]))
			return (VI_STRINGFILEATTRIBUTE)(i + 1);
	}
	return VISFA_Unknown;
}

/* QueryAttribute - [method] tries to find a version attribute and returns its value. If the file has no version resource, an interface error of ERROR_RESOURCE_DATA_NOT_FOUND is generated. If the file defines a version resource, but if the particular attribute does not exist, an interface error of ERROR_RESOURCE_TYPE_NOT_FOUND is generated.

Parameters:
Name - [in] contains the name of an attribute to search the version info for.
Value - [retval][out] contains the value of the queried attribute. The data type of the value returned depends on the queried attribute. A query for a fixed-length attribute returns the value as an integer (i.e., Value->vt==VT_I4). A query for a string attribute returns the value as a character string (i.e., Value->vt==VT_BSTR).
*/
STDMETHODIMP VersionInfoImpl::QueryAttribute(/* [in] */ BSTR Name, /* [retval][out] */ VARIANT *Value)
{
	if (!Name || *Name == 0)
		return E_INVALIDARG;

	HRESULT hr;
	bstring value;
	UINT dataLen = 0;
	VI_FIXEDFILEATTRIBUTE ffa = _getFixedFileAttributeId(Name);
	if (ffa != VIFFA_Unknown)
	{
		// a fixed-length attribute from the FixedFileInfo block is being requested for.
		// retrieve the block.
		VS_FIXEDFILEINFO* pVSFFI = NULL;
		hr = queryAttribData(L"\\", (LPVOID*)&pVSFFI, &dataLen);
		if (hr == S_OK)
		{
			if (ffa == VIFFA_FileVersion)
			{
				// return the file version as a BSTR in a <major>.<minor>.<revision>.<build> format.
				value.format(L"%d.%d.%d.%d", HIWORD(pVSFFI->dwFileVersionMS), LOWORD(pVSFFI->dwFileVersionMS), HIWORD(pVSFFI->dwFileVersionLS), LOWORD(pVSFFI->dwFileVersionLS)
				);
				Value->vt = VT_BSTR;
				Value->bstrVal = value.detach();
			}
			else if (ffa == VIFFA_ProductVersion)
			{
				// return the product version as a BSTR in a <major>.<minor>.<revision>.<build> format.
				value.format(L"%d.%d.%d.%d", HIWORD(pVSFFI->dwProductVersionMS), LOWORD(pVSFFI->dwProductVersionMS), HIWORD(pVSFFI->dwProductVersionLS), LOWORD(pVSFFI->dwProductVersionLS)
				);
				Value->vt = VT_BSTR;
				Value->bstrVal = value.detach();
			}
			else if (ffa == VIFFA_Signature)
			{
				hr = queryAttribData(L"\\", (LPVOID*)&pVSFFI, &dataLen);
				if (hr == S_OK)
				{
					Value->vt = VT_I4;
					Value->lVal = pVSFFI->dwSignature;
				}
			}
			else if (ffa == VIFFA_StrucVersion)
			{
				Value->vt = VT_I4;
				Value->lVal = pVSFFI->dwStrucVersion;
			}
			else if (ffa == VIFFA_FileVersionMS)
			{
				Value->vt = VT_I4;
				Value->lVal = pVSFFI->dwFileVersionMS;
			}
			else if (ffa == VIFFA_FileVersionLS)
			{
				Value->vt = VT_I4;
				Value->lVal = pVSFFI->dwFileVersionLS;
			}
			else if (ffa == VIFFA_ProductVersionMS)
			{
				Value->vt = VT_I4;
				Value->lVal = pVSFFI->dwProductVersionMS;
			}
			else if (ffa == VIFFA_ProductVersionLS)
			{
				Value->vt = VT_I4;
				Value->lVal = pVSFFI->dwProductVersionLS;
			}
			else if (ffa == VIFFA_FileFlagsMask)
			{
				Value->vt = VT_I4;
				Value->lVal = pVSFFI->dwFileFlagsMask;
			}
			else if (ffa == VIFFA_FileFlags)
			{
				Value->vt = VT_I4;
				Value->lVal = pVSFFI->dwFileFlags;
			}
			else if (ffa == VIFFA_FileOS)
			{
				Value->vt = VT_I4;
				Value->lVal = pVSFFI->dwFileOS;
			}
			else if (ffa == VIFFA_FileType)
			{
				Value->vt = VT_I4;
				Value->lVal = pVSFFI->dwFileType;
			}
			else if (ffa == VIFFA_FileSubtype)
			{
				Value->vt = VT_I4;
				Value->lVal = pVSFFI->dwFileSubtype;
			}
			else if (ffa == VIFFA_FileDate)
			{
				// convert the file date in FILETIME to OLE date.
				if (pVSFFI->dwFileDateMS && pVSFFI->dwFileDateLS)
				{
					FILETIME ft;
					ft.dwHighDateTime = pVSFFI->dwFileDateMS;
					ft.dwLowDateTime = pVSFFI->dwFileDateLS;
					SYSTEMTIME st;
					FileTimeToSystemTime(&ft, &st);
					if (SystemTimeToVariantTime(&st, &Value->date))
						Value->vt = VT_DATE;
					else
						hr = S_FALSE;
				}
				else
					hr = S_FALSE;
			}
			else
				hr = E_INVALIDARG;
		}
	}
	else
	{
		// the caller wants a variable-length text attribute, e.g., ProductName.
		// before we can make the query, we need a valid language-codepage translation. so, make sure we have one.
		ensureLangCp();
		// build a path for the attribute. it's in form "\StringFileInfo\<LangId><CodePage>\<Attribute>".
		bstring sfi;
		sfi.format(L"\\StringFileInfo\\%04X%04X\\%s", _langId, _codepage, Name);
		// retrieve the value of the attribute.
		LPCWSTR attribData = NULL;
		hr = queryAttribData(sfi, (LPVOID*)&attribData, &dataLen);
		if (hr == S_OK)
		{
			// dataLen has the number of characters of the string in attribData including a termination null.
			value.assignW(attribData, dataLen - 1);
			Value->vt = VT_BSTR;
			Value->bstrVal = value.detach();
		}
	}
	return hr;
}

/* ensureLangCp - makes sure we have a valid language and codepage. If the automation user has not set them, adopt the first entry in the translation block. */
HRESULT VersionInfoImpl::ensureLangCp()
{
	if (_langId != 0 && _codepage != 0)
		return S_FALSE;
	DWORD defLangCp = queryLangCp(NULL);
	_langId = LOWORD(defLangCp);
	_codepage = HIWORD(defLangCp);
	return S_OK;
}

/* queryAttribData - retrieves the value of an attribute from the file's version resource.

Parameters:
subblockPath - [in] contains a path into the version resource specifying the location of an attribute. Basically, there are three kinds of paths.

1) '\<Attribute>' - use this form of path to get a fixed-length attribute in the FixedFileInfo block.
2) '\VarFileInfo\Translation' - use this path to retrieve the translation block containing entries of langId and codepage pairs.
3) '\StringFileInfo\<LangId+CodePage>\<Attribute> - us this form of path to find a variable-length string attribute from a language-dependent StringFileInfo block.

Remarks:
A version info structure retrieved from the file's resource section is cached in class member _vi. It's to avoid repeating allocating space and reading the data from the resource on the file.
The caller should not try to free the data returned in attribData. It belongs to the data in cache and cannot be de-allocated.
*/
HRESULT VersionInfoImpl::queryAttribData(LPCWSTR subblockPath, LPVOID* attribData, UINT *attribLen)
{
	HRESULT hr = S_OK;
	// retrieve the entire version info structure by calling win32 GetFileVersionInfo.
	if (_vi.length() == 0)
		hr = queryVersionInfo(_vi);
	if (hr == S_OK)
	{
		// get to the requested subblock in the version structure.
		if (!VerQueryValue((LPVOID)_vi._b, subblockPath, attribData, attribLen))
			hr = HRESULT_FROM_WIN32(ERROR_RESOURCE_DATA_NOT_FOUND); // No version info is available for this library file.
	}
	return hr;
}

/* queryLangCp - retreives a language-codepage pair at a given index from the translation subblock. If no index is supplied, the method returns the first pair it finds.

Parameters:
langIndex - [optional, in] an index to a langId+codepage entry in the translation block of the version resource. If the parameter is not defined (i.e., set to NULL), the first translation entry is returned.
*/
DWORD VersionInfoImpl::queryLangCp(VARIANT *langIndex)
{
	UINT dataLen = 0;
	DWORD *langList = NULL;
	HRESULT hr = queryAttribData(L"\\VarFileInfo\\Translation", (LPVOID*)&langList, &dataLen);
	if (hr != S_OK)
		return { 0 }; // a corrupted version info?
	// langIndex contains a one-based index to a lang-code element in the Translation table.
	int j = 0;
	if (langIndex)
	{
		if (langIndex->vt == VT_I4)
			j = langIndex->lVal - 1;
		else if (langIndex->vt == VT_I2)
			j = langIndex->iVal - 1;
	}
	// get the language and code page entry at j.
	if (0 <= j && j < (int)(dataLen / sizeof(DWORD)))
	{
		return langList[j];
	}
	// Subscript out of range
	return { 0 };
}

/* queryVersionInfo - retrieves the entire version info structure from the file's resource section. Note that the method uses a BSTR to pass the binary structure. So, the allocation passed back in the output parameter, strVerInfo, is binary and may not be 2-byte aligned. To read the content, typecast it to LPVOID and pass it to Win32 VerQueryValue().

Parameters:
strVerInfo - [out] contains an entire version info structure retrieved from the file's resource section when the method returns control. The version info is read with a call to Win32 GetFileVersionInfo. make sure that strVerInfo has not allocated a string before the call is made. This method does not free an existing allocation. It would simply overwrite with a new allocation, causing the older allocation to leak.
*/
HRESULT VersionInfoImpl::queryVersionInfo(bstring& strVerInfo)
{
	if (_file.length() == 0)
		return E_UNEXPECTED;
	DWORD dwHandle;
	DWORD cbVerInfo = GetFileVersionInfoSize(_file, &dwHandle);
	if (cbVerInfo == 0)
	{
		DWORD errorCode = GetLastError();
		/* resource-related error codes: 
		0x714 - ERROR_RESOURCE_DATA_NOT_FOUND
		0x715 - ERROR_RESOURCE_TYPE_NOT_FOUND
		0x716 - ERROR_RESOURCE_NAME_NOT_FOUND
		0x717 - ERROR_RESOURCE_LANG_NOT_FOUND
		*/
		return HRESULT_FROM_WIN32(errorCode);
	}
	LPBYTE pbVerInfo = (LPBYTE)strVerInfo.byteAlloc(cbVerInfo);
	if (!GetFileVersionInfo(_file, dwHandle, cbVerInfo, pbVerInfo))
		return HRESULT_FROM_WIN32(GetLastError());
	return S_OK;
}

/* queryVersionNumber - retrieves the major+minor version of the file from the FixedFileInfo block of the version info resource. HIWORD(*Value) is the major version number, while LOWORD(*Value) the minor version number. If the file has no version resource, the method returns an ERROR_RESOURCE_* error code.

Parameters:
Value - [out] contains the value of VS_FIXEDFILEINFO::dwFileVersionMS, a fixed-length file version attribute, if the file has a version resource. If the file has no version resource, an appropriate error is returned. See queryVersionInfo for likely error codes.
*/
HRESULT VersionInfoImpl::queryVersionNumber(long *Value)
{
	HRESULT hr = S_OK;
	if (_vi.length() == 0)
		hr = queryVersionInfo(_vi);
	if (hr == S_OK)
	{
		UINT cbVSFFI = 0;
		VS_FIXEDFILEINFO *pVSFFI = NULL;
		if (!VerQueryValue((LPVOID)_vi._b, L"\\", (LPVOID*)&pVSFFI, &cbVSFFI))
			return HRESULT_FROM_WIN32(ERROR_RESOURCE_TYPE_NOT_FOUND); // FixedFileInfo is not available for this file. the resource section must be corrupt.
		*Value = pVSFFI->dwFileVersionMS;
	}
	return hr;
}

