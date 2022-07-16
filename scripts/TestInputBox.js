IB_BROWSERBUTTONTYPE_FILE_OPEN=1;
IB_BROWSERBUTTONTYPE_FILE_SAVEAS=2;
IB_BROWSERBUTTONTYPE_FOLDER=3;

var fso = new ActiveXObject("Scripting.FileSystemObject");
var wsh = new ActiveXObject("WScript.Shell");
var inputbox = new ActiveXObject("MaxsUtilLib.InputBox");
inputbox.Caption = WScript.ScriptName;
inputbox.Note = "Note: Selected image will be processed.";
inputbox.BrowserButtonType = IB_BROWSERBUTTONTYPE_FILE_OPEN;
inputbox.FileFilter = "Image files|*.jpg;*.png;*.bmp|All files|*.*";
if (inputbox.Show("Enter path of an image file") <= 0)
  WScript.Quit(2);
var pathname = inputbox.InputValue;

