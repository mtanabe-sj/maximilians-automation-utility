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

/* ListFileVersions - uses COM server MaxsUtil to list files in a folder, each with version info retrieved from the file's version resource.

The main purpose of this program is to test the automation interface of MaxsUtil.VersionInfo.

These are the operations the program performs.
1) A user clicks the Folder button. Button_Click() responds by running a folder browsing dialog.
2) The user selects a folder. The folder will be scanned for files.
3) A ProgressBox is set up for tracking the scan progress.
4) An ObservableCollection of struct VersionAttribList is created. VersionAttribList collects essential version attributes of each file from the VersionInfo. 
5) Iterate through the files in the folder. Increment the progress position. a VersionAttribList is created per file. It's added to the ObservableCollection.
6) Lastly, the collection is assigned to FileVersionList as its data source. This results in display of the version info of the files in a list view in the main window.

Make sure that the correct x64 or x86 edition of MaxsUtil.dll is registered with your system. Check Project Properties. Remember the 'Prefer 32-bit' option is enabled by default.
*/

using System;
using System.Windows;
using System.IO;
using System.Windows.Forms; // need Forms for FolderBrowserDialog(). make sure this assembly is added to References.
using System.Collections.ObjectModel;
using System.Windows.Controls;


namespace ListFileVersions
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		bool scanCanceled;

		public MainWindow()
		{
			InitializeComponent();
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			String dirPath=null;
			using (var dlg = new FolderBrowserDialog())
			{
				var res = dlg.ShowDialog();
				if (res == System.Windows.Forms.DialogResult.OK)
					dirPath = dlg.SelectedPath;
			}
			if (dirPath == null)
				return;

			FolderPath.Text = dirPath;

			int fileCount = Directory.GetFiles(dirPath).Length;

			var progress = new MaxsUtilLib.ProgressBox();
			// this invokes a QI call for IConnectionPointContainer to ProgressBox.
			progress.Cancel += Progress_Cancel;
			scanCanceled = false;

			// configure the progress box with a status message and progress range, and move it to the upper left corner of our main window.
			progress.Message = "Scanning "+fileCount+"files...";
			progress.UpperBound = fileCount;
			Point pt = this.PointToScreen(new Point(0, 0));
			progress.Move(MaxsUtilLib.PROGRESSBOXMOVEFLAG.PROGRESSBOXMOVEFLAG_X_Y, pt.X, pt.Y);
			// start the progress display.
			progress.Start();

			// this will collect version info queried from each file in the source directory.
			var c = new ObservableCollection<VersionAttribList>();

			// run a directory scan and make a version query on each file found.
			int i = 0;
			var files = Directory.EnumerateFiles(dirPath);
			foreach(var f in files)
			{
				progress.Note = "File " + (++i) + ": " + Path.GetFileName(f);
				var val = new VersionAttribList(f);
				c.Add(val);
				progress.Increment();
				if (scanCanceled)
				{
					// quit. the cancel button has been hit.
					progress.Message = "Scan canceled.";
					break;
				}
			}

			FileVersionList.ItemsSource = c;

			progress.Stop();
		}

		// responds to a cancelation event from ProgressBox.
		private void Progress_Cancel(ref bool CanCancel)
		{
			// just accept the cancel state. we could ask the user for confirmation and reverse CanCancel to force ProgressBox to continue.
			scanCanceled = CanCancel;
		}

		double folderpathLeft = 0;
		double fileversionlistTop = 0;
		private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
		{
			Grid grid = (Grid)this.Content;
			if(folderpathLeft==0)
			{
				Point pt1 = FolderPath.TranslatePoint(new Point(0, 0), grid);
				folderpathLeft = pt1.X;
			}
			if(fileversionlistTop==0)
			{
				Point pt2 = FileVersionList.TranslatePoint(new Point(0, 0), grid);
				fileversionlistTop = pt2.Y;
			}
			FolderPath.Width = grid.ActualWidth - folderpathLeft - FolderPath.Margin.Right;
			FileVersionList.Height = grid.ActualHeight - fileversionlistTop - FileVersionList.Margin.Bottom;
		}
	}
}
