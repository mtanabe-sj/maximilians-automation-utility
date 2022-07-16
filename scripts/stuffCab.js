/*
  Copyright (c) 2022 Makoto Tanabe <mtanabe.sj@outlook.com>
  Licensed under the MIT License.
*/
/* stuffCab.js - has a user pick a directory and compresses it in a .cab file placed in the parent directory.

Usage: stuffCab.js [directory]

This script program uses MaxsUtil.dll, a COM server with utility functions, as well as the FileSystemObject and WScript.Shell automation objects of the system.

CAB format reference:
https://docs.microsoft.com/en-us/previous-versions/bb417343(v=msdn.10)
*/

// ButtonType constants for WScript.Shell.Popup(Message, TimeoutSeconds, Caption, ButtonType).
MB_OK=0;
MB_OKCANCEL=1;
MB_YESNO=4;
MB_YESNOCANCEL=3;
MB_RETRYCANCEL=5;
MB_ICONQUESTION=0x20;
MB_DEFBUTTON2=0x100;
MB_SYSTEMMODAL=0x1000;
MB_SETFOREGROUND=0x10000;
// possible return values from Shell.Popup.
IDOK=1;
IDCANCEL=2;
IDABORT=3;
IDRETRY=4;
IDIGNORE=5;
IDYES=6;
IDNO=7;

// options for MaxsUtilLib.ProgressBox.Start
PROGRESSBOXSTARTOPTION_DISABLE_CANCEL = 1;
PROGRESSBOXSTARTOPTION_SHOW_PROGRESSBAR = 2;
PROGRESSBOXSTARTOPTION_APPEND_TO_NOTE = 4;

// possible browser types for MaxsUtilLib.BrowserButtonType
IB_BROWSERBUTTONTYPE_FILE_OPEN = 1;
IB_BROWSERBUTTONTYPE_FILE_SAVEAS = 2;
IB_BROWSERBUTTONTYPE_FOLDER = 3;

// registry key where our input box initialization settings like initial directory path are stored.
APPREGKEY = "HKCU\\Software\\mtanabe\\MaxsUtilLib\\stuffCab\\";


var fso = new ActiveXObject("Scripting.FileSystemObject");
var wsh = new ActiveXObject("WScript.Shell");

var log = new eventLog;
log.write(WScript.ScriptName + " Started");

// collect configuration info from the user.
var params = new cabParams;
if (!params.configure(WScript.Arguments.length > 0? WScript.Arguments(0):""))
  WScript.Quit(2); // canceled. quit.

log.write("SOURCE DIRECTORY", params.srcDir);
log.write("CAB OUTPUT PATH", params.destCAB);

// create a ProgressBox and start a status display.
var progress = new ActiveXObject("MaxsUtilLib.ProgressBox");
progress.Caption = WScript.ScriptName;
progress.Message = "Scanning directory...";
progress.Start();
// the first task is to scan the source directory and count how many files there are in it.
var fileCount = scanDirectory(params.srcDir, params.includeNested);
if (progress.Canceled)
  WScript.Quit(1);
if (fileCount == 0) {
	WScript.Echo("No file exists in the directory.");
	WScript.Quit(2);
}
// use the tally to set the upper bound of the progress range.
progress.Message = "Generating DDF for "+fileCount+" files...";
progress.UpperBound = fileCount;
progress.Start(PROGRESSBOXSTARTOPTION_SHOW_PROGRESSBAR);

// create a DDF list of file paths and timestamps in ascii. makecab cannot handle unicode.
var ddfList = fso.CreateTextFile(params.tmpDDF, true, false);
ddfList.WriteLine(".Set DiskDirectory1=" + params.diskDir);
ddfList.WriteLine(".Set CabinetName1=" + params.tmpPrefix+params.srcName+".cab");
ddfList.WriteLine(".Set UniqueFiles=OFF");
ddfList.WriteLine(".Set Cabinet=ON");
ddfList.WriteLine(".Set Compress=ON");
ddfList.WriteLine(".Set MaxDiskSize=0");
// walk the files in the source directory and add them to the ddf. if the user has elected to include nested directories in the scan, iterate and pick up the files in all nested directories, too.
var wideNames = 0; // this counter is used to tally non-english filenames. makecab cannot handle multiple languages.
var unpackedSize = writeDirectoryToDDF(params.srcDir, ddfList, wideNames);
ddfList.Close();
// check cancelation by the user.
if (progress.Canceled)
  WScript.Quit(1);

// alert the user if we've encountered international filenames. the user will have to choose between cancel the job and ignore the name problem and march on.
if (wideNames > 0) {
	log.write(WScript.ScriptName+" Interrupted.");
	wsh.Run("Notepad \""+log._path+"\"");
  if (IDCANCEL == wsh.Popup("Your directory contains file(s) with wide character(s). Compression tool (MAKECAB) does not support a file name in Unicode. Refer to the log for the names. Either rename them, or move them out of the directory. Or, choose Retry to ignore and try to continue without the problem file(s).", 0, WScript.ScriptName, MB_RETRYCANCEL|MB_DEFBUTTON2))
    WScript.Quit(1);
	log.write(WScript.ScriptName+" Resuming.");
}

// hide the progress bar since we won't be able to meter the compression progress which is our next and final part of the job.
progress.HideProgressBar();
progress.Message = "Compressing "+fileCount+" files ("+getByteCountString(unpackedSize)+" total).";
progress.Note = "Waiting for the compression to complete. This may take a while.";
progress.Start(PROGRESSBOXSTARTOPTION_APPEND_TO_NOTE); // put ProgressBox in append mode. we want to keep all text output by makecab by appending new text to the old.

// generate a batch file in %TEMP%. it will prepare the staging area and run compression tool makecab on the ddf we've prepared.
var cabBatch = fso.CreateTextFile(params.tmpBAT, true);
cabBatch.WriteLine(params.tmpDrive);
cabBatch.WriteLine("CD "+params.diskDir);
cabBatch.WriteLine("DEL "+params.tmpCAB);
cabBatch.WriteLine("MAKECAB /F "+params.tmpDDF);
cabBatch.Close();

// start the batch file and wait until makecab finishes packing the files. we will monitor by reading stdout and echo the output on the progress box.
log.write("MAKECAB Batch Started.");
var i=0;
try {
  var proc = wsh.Exec(params.tmpBAT);
  while(!proc.StdOut.AtEndOfStream) {
	  var nextline = proc.StdOut.ReadLine();
 	  log.write("StdOut", nextline);
	  progress.Note = nextline;
    WScript.Sleep(0);
    // if the user has hit the cancel button, kill the batch process and exit.
	  if (progress.Canceled) {
      log.write("MAKECAB Canceled");
      proc.Terminate();
      log.write(WScript.ScriptName+" Aborted");
      WScript.Quit(1);
    }
    if (i++ == 0) {
			// this pushes the progress box to the front of the makecab console in z-order.
      progress.Visible = true;
    }
  }
} catch(e) {
	// canceling can cause wscript to throw a pipe-closing error 800700e8.
	log.write("MAKECAB Batch Aborted", "ErrorCode="+(e.number>0? e.number : (0x100000000+e.number)).toString(16)+"; Message='"+e.message+"'");
	WScript.Quit(1);
}
progress.Message = "Compression completed";

// we're done. tell the user that and show summary stats. open the log in notepad if the user wants it.
log.write(WScript.ScriptName+" Stopped.");

if(params.showStats) {
	var cab = fso.GetFile(params.tmpCAB);
	var compressionRatio = (100*cab.Size)/unpackedSize;
	if (IDYES == wsh.Popup(
		"Compression completed.\n\nResult Summary: "+fileCount+" files totalling "+getByteCountString(unpackedSize)+" bytes compressed to " + getByteCountString(cab.Size) + " ("+compressionRatio.toFixed(1) + "%)\n\nDo you want to view the log?",
			0,
			WScript.ScriptName,
			MB_YESNO)) {
		wsh.Run("Notepad \""+log._path+"\"");
	}
}
// clean up and exit.
progress.Stop();
params.clean();
// open the compressed cab in Windows Explorer.
wsh.Run(params.destCAB);


/////////////////////////////////////////////////////////
/* This is a directory scan configuration object. It manages these settings.

1) showStats : a boolean flag for showing a summary result to and prompting the user after the cab compression completes successfully.
2) includeNested : a boolean to indicate if all nested directories should be scanned as well.
3) srcDir : a fully qualified pathname (FQPN) of a source directory subject to the scan.
4) srcName : the name of the source directory.
5) excludedDirs : names of directories subject to exclusion from the scan.
6) excludedExts : extension names of files subject to exclusion from the scan.
7) destCAB : FQPN of the output cab.
8) diskDir : FQPN of a staging directory where the cab is generated. It's set to the value of %TEMP%. This is where a ddf and a bat that runs a cab compression tool (makecab.exe) are written and operated on.

Before starting a scan, call method configure(). It initializes parameters with saved settings from  the system registry, and collects the source directory and nested scan option from the user by running MaxsUtilLib.InputBox. It saves the new settings in the registry.

While scanning, call method canIncludeDirectory() to decide if a subdirectory can be included in the scan or not. Also, call method canIncludeFile() to decide if a particular file can be included in the ddf.

After the scan, call method clean(). It deletes work files used in generating a cab.
*/
function cabParams() {
	this.tmpPrefix = getFileTitle(WScript.ScriptName)+"_";
	this.showStats = false;
	this.includeNested = false;
	this.srcDir = "";
	this.srcName = "";

	this.configure = function(initialDirectoryPath) {
		var excludedDirectories = "";
		var excludedFileExtensions = "";
		var scanNested = false;
		try {
			if (initialDirectoryPath.length == 0)
			  initialDirectoryPath = wsh.RegRead(APPREGKEY+"InitialDirectoryPath");
			excludedDirectories = wsh.RegRead(APPREGKEY+"ExcludedDirectories");
			excludedFileExtensions = wsh.RegRead(APPREGKEY+"ExcludedFileExtensions");
			scanNested = wsh.RegRead(APPREGKEY+"ScanNested");
			this.showStats = wsh.RegRead(APPREGKEY+"ShowStats");
		} catch(e) {
			log.write("cabParams.config caught RegRead exception", (e.number & 0xFFFF));
		}

		var inputbox = new ActiveXObject("MaxsUtilLib.InputBox");
		inputbox.InitialValue = initialDirectoryPath;
		inputbox.Note = "Note: The directory you specify will be packed into a cab.";
		inputbox.Caption = WScript.ScriptName;
		inputbox.CheckBox = "Include nested directories";
		inputbox.Checked = scanNested;
		inputbox.BrowserButtonType = IB_BROWSERBUTTONTYPE_FOLDER;
		if (inputbox.Show("Enter path of a directory to compress") <= 0)
			return false;

		this.includeNested = inputbox.Checked;
		this.srcDir = inputbox.InputValue;
		this.srcName = fso.GetFolder(this.srcDir).Name;
		this.destCAB = this.srcDir+".cab";
		this.diskDir = wsh.ExpandEnvironmentStrings("%TEMP%");
		this.tmpDir = this.diskDir+"\\";
		this.tmpDDF = this.tmpDir+this.tmpPrefix+this.srcName+".ddf";
		this.tmpBAT = this.tmpDir+this.tmpPrefix+this.srcName+".bat";
		this.tmpCAB = this.tmpDir+this.tmpPrefix+this.srcName+".cab";
		this.tmpDrive = this.tmpDir.substr(0,2);
		this.destDrive = this.destCAB.substr(0,2);
		this.excludedDirs = excludedDirectories.split(",");
		this.excludedExts = excludedFileExtensions.split(",");

		try {
			wsh.RegWrite(APPREGKEY+"InitialDirectoryPath", this.srcDir, "REG_SZ");
			wsh.RegWrite(APPREGKEY+"ScanNested", this.includeNested, "REG_DWORD");
		} catch(e) {
			log.write("cabParams.config caught RegWrite exception", (e.number & 0xFFFF));
		}
		return true;
	}
	this.clean = function() {
    if (fso.FileExists(this.tmpDir+"setup.inf"))
      fso.DeleteFile(this.tmpDir+"setup.inf");
    if (fso.FileExists(this.tmpDir+"setup.rpt"))
      fso.DeleteFile(this.tmpDir+"setup.rpt");
		if (fso.FileExists(this.tmpDDF))
      fso.DeleteFile(this.tmpDDF);
    if (fso.FileExists(this.tmpBAT))
      fso.DeleteFile(this.tmpBAT);
		if (fso.FileExists(this.destCAB))
      fso.DeleteFile(this.destCAB);
    if (fso.FileExists(this.tmpCAB))
		  fso.MoveFile(this.tmpCAB, this.destCAB);
	}
	this.canIncludeDirectory = function(dir) {
		var p = dir.Path.toLowerCase();
    for(var i=0; i<this.excludedDirs.length; i++) {
	    var j1=p.indexOf(this.excludedDirs[i]);
			if(j1!=-1) {
        if (j1>0 && p.charAt(j1-1)=='\\') {
				  var j2=j1+this.excludedDirs[i].length;
				  if (j2 == p.length || p.charAt(j2)=='\\')
            return false;
        }
	    }
    }
		return true;
	}
	this.canIncludeFile = function(f) {
		var j = f.Path.lastIndexOf(".");
		if (j!=-1) {
			var ext=f.Path.substr(j+1).toLowerCase();
			for(var i=0; i<this.excludedExts.length; i++) {
				if(ext==this.excludedExts[i])
				  return false;
			}
		}
		return true;
	}
}

// returns the title part of an input filename. A title is the leading part of the name excluding the trailing part (extension name) separated from the former by a period. The function assumes filename does not contain a directory pathname.
function getFileTitle(filename) {
	var pos = filename.lastIndexOf(".");
	if (pos==-1)
	  return filename;
	return filename.substr(0, pos);
}

// finds how many files are in the specified directory. looks into all subdirectories if 'includeNested' is true.
function scanDirectory(dirPath) {
	if (progress.Canceled)
	  return 0;
  var c=0;
  var fld = fso.GetFolder(dirPath);
  if (params.includeNested) {
    var dc = fld.SubFolders;
    var ed = new Enumerator(dc);
    for (; !ed.atEnd(); ed.moveNext()) {
	    var subdir = ed.item();
	    if (params.canIncludeDirectory(subdir)) {
				progress.Note = subdir.Path;
	      c += scanDirectory(subdir.Path);
	    }
    }
  }
  var fc = fld.Files;
  c += fc.Count;
  return c;
}

// converts an integer number to a double-digit number as a character string. If the number is less than 10, a 0 is prepended. For example, if numVal=3, the function returns "03". If numVal>100, it truncates the digits higher than the 2nd place. If numVal=123, the function returns "23".
function getDoubleDigitString(numVal) {
  var s = "";
  if(numVal > 99)
	  numVal = numVal % 100;
  if(numVal < 10)
	  s = "0";
  s = s+numVal.toString();
  return s
}

// formats an input date-time value as a timestamp acceptable to ddf. This is the format: '/date=MM/DD/YY /time=hh:mm:ss?' where ? is a for am or p for pm.
function getDdfTimeStamp(dateTimeVal) {
  // DDF time stamp: '/date=02/15/07 /time=01:00:00p'
  var curAmOrPm = "a";
  var hour2 = dateTimeVal.getHours();
  if(hour2 > 12) {
    hour2 -= 12;
    curAmOrPm = "p";
  }
  if(hour2 == 0)
	  hour2 = 12;
  var curYear = getDoubleDigitString(dateTimeVal.getYear());
  var curMonth = getDoubleDigitString(dateTimeVal.getMonth()+1);
  var curDay = getDoubleDigitString(dateTimeVal.getDate());
  var curHour = getDoubleDigitString(hour2);
  var curMinute = getDoubleDigitString(dateTimeVal.getMinutes());
  var curSecond = "00";
  return " /date="+curMonth+"/"+curDay+"/"+curYear+" /time="+curHour+":"+curMinute+":"+curSecond+curAmOrPm;
}

// formats an input number as a byte count, a KB count, a MB count, or a GB count depending on the position of the number's most significant digit. For example, if the input number n has more than 6 digits, the method returns a string of form "x.y MB" where x is the quotient of operation n/1000000, and y the most significant digit of the remainder. If n=1000000, the method returns string "1.0MB".
function getByteCountString(n) {
	if(n<1000)
	  return n.toString()+"B";
  if(n<1000000)
    return ((n/1000).toFixed(1)).toString()+"KB";
  if(n<1000000000)
    return ((n/1000000).toFixed(1)).toString()+"MB";
  return ((n/1000000000).toFixed(1)).toString()+"GB";
}

// returns true if all characters of txt falls in the extended ASCII range of 0 to 0xff character codes. otherwise, returns false.
function isTextInAnsiRange(txt) {
	for(var i=0; i<txt.length; i++) {
		if(txt.charAt(i)>'ÿ')
		  return false;
	}
	return true;
}

/* enumerates files in the input directory at dirPath, and writes them as ddf entries to the input file ddf. Iterates into the subdirectories if the includeNested flag is true. Increments the progress indication and displays the name of a currently scanned file in the progress box.

Note that if it has an extension name on the exclusion list, a found file is not added to ddf. Also note that if subdirectories are iterated, and if a current subdirectory has a name that matches an item in the exclusion list, that subdirectory will not be scanned.

There is a caveat. Because it does not use a locale directive in the ddf header, the ddf our script generates does not support foreign characters. Only ASCII and extended ASCII characters are supported. A file with a name having a character whose Unicode node occurs outside the 8-bit (0-0xFF) range is automatically rejected. The file will not be added to the cab.

Consult the log for indications for files excluded from the cab. Look for these key words.
EXCLUDED_DIR : the directory has a name found in an exclusion list, and therefore, was excluded from the scan.
EXCLUDED_FILE : the file has an extension name found in an exclusion list, and therefore, was removed from the scan.
NON_ASCII_FILENAME : the file was rejected due to its wide character name.
*/
function writeDirectoryToDDF(dirPath, ddf) {
	if (progress.Canceled)
    return 0;
  var size = 0;
  var fld = fso.GetFolder(dirPath);
  if(params.includeNested) {
    var dc = fld.SubFolders;
    var ed = new Enumerator(dc);
    for (; !ed.atEnd(); ed.moveNext()) {
	    var subdir = ed.item();
	    if (!params.canIncludeDirectory(subdir)) {
        log.write("EXCLUDED_DIR", subdir.Path);
			  continue;
	    }
	    size += writeDirectoryToDDF(subdir.Path, ddf);
  	  if (progress.Canceled)
        return 0;
    }
  }
  var relPath = fld.Path.substr(params.srcDir.length+1);
  ddf.WriteLine(".Set DestinationDir=\""+relPath+"\"");
  var fc = fld.Files;
  var ef = new Enumerator(fc);
  for (; !ef.atEnd(); ef.moveNext()) {
	  var f = ef.item();
	  if(!params.canIncludeFile(f)) {
			log.write("EXCLUDED_FILE", f.Path);
			continue;
		}
	  if(!isTextInAnsiRange(f.Path)) {
			log.write("NON_ASCII_FILENAME", f.Path);
			wideNames++;
			continue;
		}
	  var fileSeq = progress.ProgressPos+1;
	  progress.Note = "File "+fileSeq+": "+f.Name;
    var tstamp = getDdfTimeStamp(new Date(f.DateLastModified));
    ddf.WriteLine("\""+f.Path+"\""+tstamp);
	  progress.Increment();
		if (progress.Canceled)
	    return 0;
    size += f.Size;
  }
  log.write("DIR="+dirPath, "SIZE="+size);
  return size;
}

/////////////////////////////////////////////////////////
/* This is an event log object. It creates a file with the same name as that of this script with extension .log in current user's %TEMP% directory.
To write an event, use a call statement of form eventLog.write(event). The time of the call and the event text are written to the log.
To write an event with associated data, call eventLog.write(event, data). Event is written to the log first. Data is written next, actually preceded by a separator colon. This time, no timestamp is written.
*/
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

