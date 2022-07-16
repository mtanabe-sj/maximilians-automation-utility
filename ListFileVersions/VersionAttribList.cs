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

using System;
using System.IO;

namespace ListFileVersions
{
	public struct VersionAttribList
	{
		String _fpath;
		String _fname;
		String _fversion;
		String _pname;
		String _pversion;
		String _fdesc;
		public VersionAttribList(String filePath)
		{
			var vi = new MaxsUtilLib.VersionInfo();
			vi.File = filePath;
			_fpath = filePath;
			_fname = Path.GetFileName(filePath);
			try
			{
				_fversion = vi.VersionString;
				_pname = vi.QueryAttribute("ProductName");
				_pversion = vi.QueryAttribute("ProductVersion");
				_fdesc = vi.QueryAttribute("FileDescription");
			}
			catch (Exception e)
			{
				_fdesc = "[Error 0x"+Convert.ToString(e.HResult, 16)+"; "+e.Message+"]";
				_fversion = "(na)";
				_pname = "(na)";
				_pversion = "(na)";
			}
		}
		public String FileName { get { return _fname; } set { _fname = value; } }
		public String FileVersion { get { return _fversion; } set { _fversion = value; } }
		public String ProductName { get { return _pname; } set { _pname = value; } }
		public String ProductVersion { get { return _pversion; } set { _pversion = value; } }
		public String FileDescription { get { return _fdesc; } set { _fdesc = value; } }

	}
}
