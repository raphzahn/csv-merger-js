var electron = require('electron');
var XLSX = require('xlsx');
var app = electron.app;

var win = null;

function createWindow() {
	if (win) return;
	win = new electron.BrowserWindow({
		width: 1000, height: 440,
		webPreferences: {
			nodeIntegration: true,
			enableRemoteModule: true
		},
	});
	win.removeMenu();
	win.loadURL("file://" + __dirname + "/index.html");
	//win.webContents.openDevTools();
	win.on('closed', function () { win = null; });
}
if (app.setAboutPanelOptions) app.setAboutPanelOptions({ applicationName: 'sheetjs-electron', applicationVersion: "XLSX " + XLSX.version, copyright: "(C) 2017-present SheetJS LLC" });
app.on('open-file', function () { console.log(arguments); });
app.on('ready', createWindow);
app.on('activate', createWindow);
app.on('window-all-closed', function () { if (process.platform !== 'darwin') app.quit(); });