' TestVersionInfo.vbs
'
' this is a vbscript program to test the automation function of MaxsUtil.
' the program performs these operations.
' 1) obtain the Windows directory, e.g., c:\windows.
' 2) open a log file in the same directory this script is located in. give it a name of form '<ScriptName>.log'.
' 3) open the Windows directory and enumerate files in it.
' 4) for each file, create a VersionInfo object of MaxsUtil to gain access to the file's version resource.
' 5) query the version number and file description by reading property VersionInfo.VersionString and calling method VersionInfo.QueryAttribute("FileDescription"), respectively.
'
' note that the log file is opened for an append operation and immediately closed each time an event is written to it.
' that way, previous events can remain intact in case a fatal error happens before the program terminates normally.
'
' to run the test, simply type the name of the script at a command prompt. or, double click the file title in windows explorer.

srcDir = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%windir%")

dim fso : set fso = CreateObject("Scripting.FileSystemObject")
outPath = wscript.ScriptFullName & ".log"
set logFile = fso.CreateTextFile(outPath, true)
logFile.close
LogEvent wscript.scriptName & " Started", CStr(now)
LogEvent "DIRECTORY", srcDir

set fld = fso.GetFolder(srcDir)
set fc = fld.files
i=1
for each f in fc
  dim pv, fdesc
  fv = QueryFileVersion(f, pv, fdesc)
  if fv <> "" then
    ' fv has the file's version number string in a major.minor.revision.build format.
		' pv has a product version.
		' fdesc is a file description.
	  LogEvent i, f.name & " [" & fv & "]" & " (" & pv & ")"
	  if len(fdesc) > 0 then
	    LogEvent " DESCRIPTION", fdesc
	  end if
  else
    ' the file does not have a version info. fdesc says why.
    LogEvent i, f.name & " (" & fdesc & ")"
  end if
	i=i+1
next

LogEvent wscript.scriptName & " Stopped", CStr(now)
if vbOK = msgbox("The MaxsUtil test completed. Consult the log file and verify the test result." & vbCrLf & vbCrLf & outPath, vbOkCancel, wscript.scriptName) then
  dim wsh : set wsh = WScript.CreateObject("WScript.Shell")
  wsh.run "notepad " & outPath
end if

'***********************************************************
function QueryFileVersion(f, outProductVersion, outFileDesc)
  set vi = CreateVersionInfo(f)
  on error resume next
  ' retrieve a string of version number components (major.minor.revision.build).
  QueryFileVersion = vi.VersionString
  if err = 0 then
	  outProductVersion = vi.QueryAttribute("ProductVersion")
    ' retrieve the file description in the default language.
    outFileDesc = vi.QueryAttribute("FileDescription")
  else
      if err.number = &h80070714& then
        outFileDesc = "ERROR_RESOURCE_DATA_NOT_FOUND"
      elseif err.number = &h80070715& then
        outFileDesc = "ERROR_RESOURCE_TYPE_NOT_FOUND"
      elseif err.number = &h80070716& then
        outFileDesc = "ERROR_RESOURCE_NAME_NOT_FOUND"
      elseif err.number = &h80070717& then
        outFileDesc = "ERROR_RESOURCE_LANG_NOT_FOUND"
      else
        outFileDesc = "ERROR - 0x" & hex(err.number)
      end if
    	err.clear
    	QueryFileVersion = ""
  end if
end function

'***********************************************************
function CreateVersionInfo(f)
  on error resume next
  set vi = CreateObject("MaxsUtilLib.VersionInfo")
	if err <> 0 then
	  msgbox "Execution Stopped: MaxsUtilLib.VersionInfo Not Found",, wscript.scriptName
	  wscript.quit 2
	end if
  vi.File = f.path
  set CreateVersionInfo = vi
end function

'***********************************************************
sub LogEvent(eventId, eventData)
  set logFile = fso.OpenTextFile(outPath, 8)
  if eventData <> "" then
    logFile.writeline eventId & ": " & eventData
  else
    logFile.writeline eventId
  end if
  logFile.close
end sub
