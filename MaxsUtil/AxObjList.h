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


#define OBJLIST_GROW_SIZE 8

template <class T>
class AxObjList
{
public:
	AxObjList() : _p(NULL), _n(0), _max(0) {}
	virtual ~AxObjList() { clear(); }

	T* operator[](const int index) const
	{
		ASSERT(0 <= index && index < _n);
		return _p[index];
	}
	int size() const
	{
		return _n;
	}

	void clear(bool releaseObjs = true)
	{
		if (!_p)
			return;
		if (releaseObjs)
		{
			for (int i = 0; i < _max; i++)
			{
				T *pi = (T*)InterlockedExchangePointer((LPVOID*)(_p + i), NULL);
				if (pi)
				{
					pi->Release();
					_n--;
				}
			}
		}
		ASSERT(_n == 0);
		T **p = (T**)InterlockedExchangePointer((LPVOID*)&_p, NULL);
		free(p);
		_max = 0;
	}

	int add(T *obj)
	{
		obj->AddRef();
		int i;
		for (i = 0; i < _max; i++)
		{
			T *pi = (T*)InterlockedCompareExchangePointer((LPVOID*)(_p + i), obj, NULL);
			if (!pi)
			{
				_n++;
				return i;
			}
		}
		T **p2 = (T**)realloc(_p, (_max + OBJLIST_GROW_SIZE) * sizeof(T*));
		if (!p2)
		{
			obj->Release();
			return -1;
		}
		T **p3 = p2 + _max;
		ZeroMemory(p3, OBJLIST_GROW_SIZE * sizeof(T*));
		*p3 = obj;
		_max += OBJLIST_GROW_SIZE;
		_p = p2;
		i = _n++;
		return i;
	}
	bool remove(T *obj)
	{
		for (int i = 0; i < _max; i++)
		{
			T *pi = InterlockedCompareExchangePointer((LPVOID*)(_p + i), NULL, obj);
			if (pi)
			{
				pi->Release();
				_n--;
				ASSERT(_n >= 0);
				_repack(i);
				return true;
			}
		}
		return false;
	}
	bool removeAt(int index)
	{
		if (index < 0 || index >= _max)
			return false;
		T *pi = InterlockedExchangePointer((LPVOID*)(_p + index), NULL);
		if (!pi)
			return false;
		pi->Release();
		_n--;
		ASSERT(_n >= 0);
		_repack(index);
		return true;
	}

protected:
	int _max, _n;
	T **_p;

	void _repack(int i0)
	{
		int j = i0;
		for (int i = i0; i < _max; i++)
		{
			if (_p[i])
			{
				if (j < i)
					_p[j] = InterlockedExchangePointer((LPVOID*)(_p + i), NULL);
				j++;
			}
		}
	}
};
