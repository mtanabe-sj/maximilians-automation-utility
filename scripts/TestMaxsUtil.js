/*
  Copyright (c) 2022 Makoto Tanabe <mtanabe.sj@outlook.com>
  Licensed under the MIT License.
*/
/* TestMaxsUtil.js

This is a jscript program to test the three automation interfaces of MaxsUtil, InputBox, ProgressBox, and VersionInfo. The program performs these operations.

1) Run InputBox, and obtain a pathname to a directory. The Windows directory, i.e., c:\windows, is provided as a default value. Users can click the [...] button in InputBox to pick a folder from a tree view of folders.
2) Open a log file in the same directory this script is located in. Give it a name of form '<ScriptName>.log'.
3) Open the directory and start scaning for files. Also, start a progress display using ProgressBox.
4) For each file, create a VersionInfo object of MaxsUtil to gain access to the file's version resource.
5) Query the version number and file description by reading property VersionInfo.VersionString and calling method VersionInfo.QueryAttribute("FileDescription"), respectively.
6) Increment the progress position. Exit the loop if the Cancel button is hit.
7) Repeat steps 4-6 until all files are enumerated.
6) After the iteration, stop the progress display, alert the user, and open the log in notepad.

Note that the log file is opened for an append operation and immediately closed each time an event is written to it. That way, previous events can remain intact in case a fatal error happens before the program terminates normally.

To run the test, use WScript or CScript host. Type the name of the host followed by the name of the script at a command prompt.

> cscript <path\>TestMaxsUtil.js

You can just double click the file title in windows explorer, if .js is already associated with WScript or CScript.
*/

IB_BROWSERBUTTONTYPE_FILE_OPEN=1;
IB_BROWSERBUTTONTYPE_FILE_SAVEAS=2;
IB_BROWSERBUTTONTYPE_FOLDER=3;

var fso = new ActiveXObject("Scripting.FileSystemObject");
var wsh = new ActiveXObject("WScript.Shell");
var inputbox = new ActiveXObject("MaxsUtilLib.InputBox");
inputbox.InitialValue = wsh.ExpandEnvironmentStrings("%windir%");
inputbox.Note = "Note: directory you specify will be scanned for files with version resource.";
inputbox.Caption = WScript.ScriptName;
inputbox.BrowserButtonType = IB_BROWSERBUTTONTYPE_FOLDER;
if (inputbox.Show("Enter path of a directory to scan") <= 0)
  WScript.Quit(2);
var srcDir = inputbox.InputValue;

var log = new eventLog;
log.write(WScript.ScriptName + " Started");
log.write("DIRECTORY", srcDir);

var progress = new ActiveXObject("MaxsUtilLib.ProgressBox");

var fld = fso.GetFolder(srcDir);
var fc = fld.Files;
progress.Caption = WScript.ScriptName;
progress.Message = "Scanning "+fc.Count+" files...";
progress.UpperBound = fc.Count;
progress.Start();
var e = new Enumerator(fc);
for (; !e.atEnd(); e.moveNext()) {
	var f = e.item();
	progress.Note = "File "+(progress.ProgressPos+1)+": "+f.Name;
	if (progress.ProgressPos > 0.9*fc.Count)
    progress.BarColor = 0xFF00; // switch to green for last 10%
	var fv = new fileVersionInfo(f);
	if (fv.errCode == 0) {
	  log.write(f.Name, "["+fv.fileVersion+"] ("+fv.productVersion+")");
	  if (typeof fv.fileDescription == "string")
	    log.write(" DESCRIPTION", fv.fileDescription);
	} else {
		log.write(f.Name, "["+fv.errDesc+"]");
	}
	progress.Increment();
	if (progress.Canceled)
	{
		progress.Message = "Scanning canceled.";
		break;
	}
}

log.write(WScript.ScriptName + " Stopped")
WScript.Echo("The MaxsUtil test completed. The log file will open. Review the test result.");
wsh.run("notepad " + log._path);

progress.Stop();


function fileVersionInfo(f) {
	this._vi = new ActiveXObject("MaxsUtilLib.VersionInfo");
	this._vi.File = f.Path;
	this.errCode = 0;
  this.errDesc = "";
	try {
  	this.fileVersion = this._vi.VersionString;
	  this.productVersion = this._vi.QueryAttribute("ProductVersion");
	  this.fileDescription = this._vi.QueryAttribute("FileDescription");
	} catch(e) {
		this.errCode = e.number & 0xFFFF;
    if (this.errCode == 0x714)
      this.errDesc = "ERROR_RESOURCE_DATA_NOT_FOUND";
    else if (this.errCode == 0x715)
      this.errDesc = "ERROR_RESOURCE_TYPE_NOT_FOUND";
    else if (this.errCode == 0x716)
      this.errDesc = "ERROR_RESOURCE_NAME_NOT_FOUND";
    else if (this.errCode == 0x717)
      this.errDesc = "ERROR_RESOURCE_LANG_NOT_FOUND";
    else
      this.errDesc = "ERROR "+e.number.toString();
	}
}

function eventLog() {
	var tempDir = wsh.ExpandEnvironmentStrings("%TEMP%");
	this._path = tempDir + "\\" + WScript.ScriptName + ".log";
  var f = fso.CreateTextFile(this._path, true, true);
  f.Close();
	this.write = function(eventLabel, eventData) {
    var f = fso.OpenTextFile(this._path, 8, false, -1);
    if (typeof eventData == "string" || typeof eventData == "number") {
      f.Write(eventLabel+": ");
      f.WriteLine(eventData);
    } else {
			var now = new Date;
			var time = "["+now.getHours()+":"+now.getMinutes()+":"+now.getSeconds()+"] ";
      f.WriteLine(time+eventLabel);
	  }
    f.Close();
	};
}
