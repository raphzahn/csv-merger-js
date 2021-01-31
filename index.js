const XLSX = require('xlsx');
const electron = require('electron').remote;

const EXTENSIONS = "xls|xlsx|xlsm|xlsb|xml|csv|txt|dif|sylk|slk|prn|ods|fods|htm|html".split("|");

const data = [];
const dataTxt = [];
const dataXlsx = [];
const dataCsv = [];


const combine = function(file) {
	file.SheetNames.forEach(function(sheetName) {
		for(var i = 0; i < XLSX.utils.sheet_to_json(file.Sheets[sheetName]).length; i++){
			data.push(XLSX.utils.sheet_to_json(file.Sheets[sheetName])[i]);
		}
	});
};

const createTxt = function(data){
	for(var i = 0;i < data.length;i++){
		if(!dataTxt.includes(data[i].Vermittler)){
			dataTxt.push(data[i].Vermittler);
		}
	}
}

const createXlsx = function(data){
	for(var i = 0;i < data.length;i++){
		Json = [];
		Json["VM-Policeninfo"] = data[i]["VM-Policeninfo"];
		Json["Vermittler"] = data[i]["Vermittlers"];
		dataXlsx.push(Json);
		console.log(Json);
	}
}

const readFile = function(files) {
	
	for(var i = 0; i < files.length; i++){
		let f = files[i];
		let reader = new FileReader();
		reader.onload = function(e) {
			let data = e.target.result;
			data = new Uint8Array(data);
			combine(XLSX.read(data, {type: 'array'}));
		};
		reader.readAsArrayBuffer(f);
	}
	createTxt(data);
	console.log(data);

};

const handleReadBtn = async function() {
	const o = await electron.dialog.showOpenDialog({
		title: 'Select a file',
		filters: [{
			name: "Spreadsheets",
			extensions: 'csv'
		}],
		properties: ['openFile', 'multiSelections']
	});
	for (var i = 0; i < o.filePaths.length; i++) { combine(XLSX.readFile(o.filePaths[i]))};
	const XPORT = document.getElementById('exportBtn');
	XPORT.disabled = false;
	createTxt(data);
	createXlsx(data);
	var dataTxt1 = JSON.stringify(dataTxt);
	var dataTxt2 = dataTxt1.substring(1);
	var dataTxt3 = dataTxt2.slice(0, -1);
	console.log(dataTxt3);
	console.log(dataXlsx)
	console.log(data);
	
}
	

const exportXlsx = async function() {
	const o = await electron.dialog.showSaveDialog({
		title: 'Save file as',
		filters: [{
			name: "Spreadsheets",
			extensions: EXTENSIONS
		}]
	});
	console.log(o.filePath);
	var wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataXlsx))
	XLSX.writeFile(wb, o.filePath);
	electron.dialog.showMessageBox({ message: "Exported data to " + o.filePath, buttons: ["OK"] });
};

// add event listeners
const readBtn = document.getElementById('readBtn');
const readIn = document.getElementById('readIn');
const exportBtn = document.getElementById('exportBtn');
const drop = document.getElementById('drop');

readBtn.addEventListener('click', handleReadBtn, false);
readIn.addEventListener('change', (e) => { readFile(e.target.files); }, false);
exportBtn.addEventListener('click', exportXlsx, false);
drop.addEventListener('dragenter', (e) => {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
}, false);
drop.addEventListener('dragover', (e) => {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
}, false);
drop.addEventListener('drop', (e) => {
	e.stopPropagation();
	e.preventDefault();
	readFile(e.dataTransfer.files);
}, false);

exportBtn.disabled = true;