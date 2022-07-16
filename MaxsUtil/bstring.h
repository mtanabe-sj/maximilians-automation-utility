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
#include <windows.h>
#include <stdio.h>

/* bstring wraps BSTR, a wide character string type often used in passing text data between COM automation server and client components. It automatically frees an allocated string when the destructor is invoked. It implements essential string assignment and string formatting functions. It also implements assignment and concatenation operators.
*/
class bstring
{
public:
	bstring() : _b(NULL) {}
	bstring(const bstring&src) : _b(NULL) { assignW(src); }
	bstring(LPCWSTR s, int slen = -1) :_b(NULL) { assignW(s, slen); }
	~bstring() { free(); }
	BSTR _b;

	BSTR detach()
	{
		return (BSTR)InterlockedExchangePointer((LPVOID*)&_b, NULL);
	}
	void attach(bstring &src)
	{
		free();
		_b = src.detach();
	}

	operator BSTR() const { return _b; }
	operator LPCWSTR() const { return _b ? _b : L""; }
	BSTR* operator&() { return &_b; }
	int length() const { return _b ? (int)SysStringLen(_b) : 0; }
	const bstring& operator=(LPCWSTR s)
	{
		assignW(s);
		return *this;
	}
	const bstring& operator+=(LPCWSTR s)
	{
		appendW(s);
		return *this;
	}
	int countChar(WCHAR sep) const
	{
		int i, n = 0, len = length();
		for (i = 0; i < len; i++)
		{
			if (_b[i] == sep)
				n++;
		}
		return n;
	}
#ifdef NEED_BSTRING_SPLIT
	std::vector<bstring> split(WCHAR sep) const
	{
		std::vector<bstring> v;
		int i, i0, cc, len = length();
		for (i = 0, i0 = 0; i < len; i++)
		{
			if (_b[i] != sep)
				continue;
			cc = i - i0;
			v.push_back(bstring(_b + i0, cc));
			i0 = i + 1;
		}
		if ((cc = i - i0) > 0)
			v.push_back(bstring(_b + i0, cc));
		return v;
	}
#endif//#ifdef NEED_BSTRING_SPLIT
	bool assignW(LPCWSTR s, int slen = -1)
	{
		free();
		if (!s)
			return true;
		if (slen != -1)
			CopyMemory(alloc(slen), s, slen * sizeof(WCHAR));
		else
			_b = SysAllocString(s);
		if (!_b)
			return false;
		return true;
	}
	bool appendW(LPCWSTR s, int slen = -1)
	{
		if (!s)
			return true;
		int n1 = 0;
		int n2 = slen == -1 ? (int)wcslen(s) : slen;
		if (n2 == 0)
			return true;
		if (_b)
			n1 = SysStringLen(_b);
		int n = n1 + n2;
		BSTR b2 = SysAllocStringLen(NULL, n);
		if (!b2)
			return false;
		CopyMemory(b2, _b, n1 * sizeof(WCHAR));
		CopyMemory(b2 + n1, s, n2 * sizeof(WCHAR));
		free();
		_b = b2;
		return true;
	}
	bool assignA(LPCSTR s, int slen = -1, int cp = CP_UTF8)
	{
		free();
		if (!s)
			return true;
		int n = MultiByteToWideChar(cp, 0, s, slen, NULL, 0);
		if (!alloc(n)) // this actually allocates n+1 wchar space.
			return false;
		MultiByteToWideChar(cp, 0, s, slen, _b, n + 1);
		return true;
	}
	bool appendA(LPCSTR s, int slen = -1, int cp = CP_UTF8)
	{
		if (!s)
			return true;
		int n2 = MultiByteToWideChar(cp, 0, s, slen, NULL, 0);
		int n1 = 0;
		if (_b)
			n1 = SysStringLen(_b);
		int n = n1 + n2;
		BSTR b2 = SysAllocStringLen(NULL, n);
		if (!b2)
			return false;
		wcscpy_s(b2, n, _b);
		MultiByteToWideChar(cp, 0, s, slen, b2 + n1, n2 + 1);
		free();
		_b = b2;
		return true;
	}
	bool format(LPCWSTR fmt, ...)
	{
		va_list args;
		va_start(args, fmt);
		bool ret = vformat(fmt, args);
		va_end(args);
		return ret;
	}
	bool vformat(LPCWSTR fmt, va_list &args)
	{
		free();
		// find out how many character spaces the formatting requires.
		int n = _vscwprintf(fmt, args);
		// allocate necessary space in buffer _b. note that alloc() allocates space for n unicode characters plus a termination null char.
		if (!alloc(n))
			return false;
		// format the arguments and write the string in the allocated buffer.
		vswprintf_s(_b, n + 1, fmt, args);
		return true;
	}
	bool loadString(HINSTANCE hinst, ULONG ids) {
		WCHAR buf[256] = { 0 };
		int cc = ::LoadString(hinst, ids, buf, ARRAYSIZE(buf));
		return assignW(buf, cc);
	}
	BSTR clone() const
	{
		if (!_b)
			return NULL;
		return SysAllocString(_b);
	}

	// Allocate a BSTR of character length len+1.
	BSTR alloc(ULONG len)
	{
		// Notice that free() is not called before the allocation. Call free() before calling alloc() if the instance has made an allocation.
		// This allocates space for len+1 characters. The extra space holds a null char.
		_b = SysAllocStringLen(NULL, len);
		return _b;
	}
	// Allocate a BSTR of len+2 bytes.
	LPBYTE byteAlloc(ULONG len)
	{
		// Note that a null wide char pads the len bytes, making the total bytes allocated len+2.
		_b = SysAllocStringByteLen(NULL, len);
		return (LPBYTE)_b;
	}
	void free()
	{
		BSTR b = (BSTR)InterlockedExchangePointer((LPVOID*)&_b, NULL);
		if (b)
			SysFreeString(b);
	}
};

// this bstring subclass features a constructor that formats arguments to a string which is then assigned to bstring::_b.
class bstringv : public bstring
{
public:
	bstringv(LPCWSTR fmt, ...)
	{
		va_list	args;
		va_start(args, fmt);
		bool success = vformat(fmt, args);
		va_end(args);
	}
	bstringv(HINSTANCE hinst, UINT fmtId, ...)
	{
		bstring fmt;
		if (!fmt.loadString(hinst, fmtId))
			return;
		va_list	args;
		va_start(args, fmt);
		bool success = vformat(fmt, args);
		va_end(args);
	}
};

/* use this bstring subclass to append text which is ensured to use CRLF pairs rather than just LFs or just CRs.
*/
class bstringCRLF : public bstring
{
public:
	bstringCRLF() {}
	bstringCRLF(LPCWSTR s, int slen = -1) : bstring(s, slen) {}

	// auto-append a CRLF after the new text of newVal.
	bool append(LPCWSTR newVal, bool addCRLF=true)
	{
		int cr = 0;
		int lf = 0;
		int n;
		WCHAR c, c0 = 0;
		LPCWSTR p, p0;
		p0 = p = newVal;
		while ((c = *p) != 0)
		{
			p++;
			if (c == '\n')
			{
				n = (int)(INT_PTR)(p - p0);
				if (c0 == '\r')
				{
					if (!appendW(p0, n))
						return false;
				}
				else
				{
					n--;
					bstring line;
					line.alloc(n + 2);
					CopyMemory(line._b, p0, n * sizeof(WCHAR));
					line._b[n++] = '\r';
					line._b[n] = '\n';
					if (!appendW(line))
						return false;
				}
				p0 = p;
				lf++;
			}
			else if (c0 == '\r')
			{
				n = (int)(INT_PTR)(p - p0);
				bstring line;
				line.alloc(n + 1);
				CopyMemory(line._b, p0, n * sizeof(WCHAR));
				line._b[n - 1] = '\n';
				line._b[n] = c;
				if (!appendW(line))
					return false;
				p0 = p;
				cr++;
			}
			c0 = c;
		}
		if (p0 < p)
			appendW(p0);
		if (addCRLF)
			return appendW(L"\r\n");
		return true;
	}
};