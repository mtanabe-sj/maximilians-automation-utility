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

#define stringify(s) #s
#define xstringify(s) stringify(s)


#define FILEVER_MAJOR 1
#define FILEVER_MINOR 0
#define FILEVER_REVISION 0
#define FILEVER_BUILD 1

#define PRODVER_MAJOR 1
#define PRODVER_MINOR 0
#define PRODVER_REVISION 0
#define PRODVER_BUILD 1

#define FILEVER			FILEVER_MAJOR,FILEVER_MINOR,FILEVER_REVISION,FILEVER_BUILD
#define PRODUCTVER		PRODVER_MAJOR,PRODVER_MINOR,PRODVER_REVISION,PRODVER_BUILD
#define STRFILEVER		xstringify(FILEVER)
#define STRPRODUCTVER	xstringify(PRODUCTVER)

