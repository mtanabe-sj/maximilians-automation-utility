var pathname = "c:\\windows\\explorer.exe";

var fso = new ActiveXObject("Scripting.FileSystemObject");
var f = fso.GetFile(pathname);
var fv = new fileVersionInfo(f);
if(fv.errCode == 0)
  WScript.Echo("Path: "+pathname+"\n\nFile: "+fv.fileDescription+"; "+fv.fileVersion+"\nProduct: "+fv.productName+" "+fv.productVersion+"\nCompany: "+fv.companyName);
else
  WScript.Echo(fv.errDesc+"; Code="+fv.errCode);

function fileVersionInfo(f) {
	this._vi = new ActiveXObject("MaxsUtilLib.VersionInfo");
	this._vi.File = f.Path;
	this.errCode = 0;
	this.errDesc = "";
	try {
		this.fileVersion = this._vi.VersionString;
		this.fileDescription = this._vi.QueryAttribute("FileDescription");
		this.productName = this._vi.QueryAttribute("ProductName");
		this.productVersion = this._vi.QueryAttribute("ProductVersion");
		this.companyName = this._vi.QueryAttribute("CompanyName");
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
