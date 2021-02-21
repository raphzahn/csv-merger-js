const XLSX = require('xlsx');
const electron = require('electron').remote;
const fs = require('fs') ;
const path = require('path');

const data = [];
const dataTxt = [];
const dataXlsx = [];
const dataCsv = [];
let dir = "";


const combine = function(file) {
	file.SheetNames.forEach(function(sheetName) {
		for(var i = 0; i < XLSX.utils.sheet_to_json(file.Sheets[sheetName]).length; i++){
			data.push(XLSX.utils.sheet_to_json(file.Sheets[sheetName])[i]);
		}
	});
};

const createTxt = function(data){
	for(var i = 0; i < data.length; i++){
		if(!dataTxt.includes(data[i]["Vermittler"])){
			dataTxt.push(data[i]["Vermittler"]);
		}
	}
	
}

const createXlsxAndCsv = function(data){
	for(var i = 0; i < data.length; i++){
		Json = [];
		Json["VM-Policeninfo"] = data[i]["VM-Policeninfo"];
		Json["Vermittler"] = data[i]["Vermittler"];
		Json["Quittungs-Nr."] = data[i]["Quittungs-Nr."];
		Json["Sprache"] = data[i]["Sprache"];
		Json["Anrede"] = data[i]["Anrede"];
		Json["Name"] = data[i]["Name"];
		Json["Vorname"] = data[i]["Vorname"];
		Json["Name2"] = data[i]["Name2"];
		Json["Strasse"] = data[i]["Strasse"];
		Json["Postfach"] = data[i]["Postfach"];
		Json["PLZ/Ort"] = data[i]["PLZ/Ort"];
		Json["Policen-Nr.1"] = data[i]["Policen-Nr.1"];
		Json["Produkt1"] = data[i]["Produkt1"];
		Json["Gewinn1"] = data[i]["Gewinn1"];
		Json["Policen-Nr.2"] = data[i]["Policen-Nr.2"];
		Json["Produkt2"] = data[i]["Produkt2"];
		Json["Gewinn2"] = data[i]["Gewinn2"];
		Json["Policen-Nr.3"] = data[i]["Policen-Nr.3"];
		Json["Produkt3"] = data[i]["Produkt3"];
		Json["Gewinn3"] = data[i]["Gewinn3"];
		Json["Policen-Nr.4"] = data[i]["Policen-Nr.4"];
		Json["Produkt4"] = data[i]["Produkt4"];
		Json["Gewinn4"] = data[i]["Gewinn4"];
		Json["Policen-Nr.5"] = data[i]["Policen-Nr.5"];
		Json["Produkt5"] = data[i]["Produkt5"];
		Json["Gewinn5"] = data[i]["Gewinn5"];
		Json["Policen-Nr.6"] = data[i]["Policen-Nr.6"];
		Json["Produkt6"] = data[i]["Produkt6"];
		Json["Gewinn6"] = data[i]["Gewinn6"];
		Json["Policen-Nr.7"] = data[i]["Policen-Nr.7"];
		Json["Produkt7"] = data[i]["Produkt7"];
		Json["Gewinn7"] = data[i]["Gewinn7"];
		Json["Policen-Nr.8"] = data[i]["Policen-Nr.8"];
		Json["Produkt8"] = data[i]["Produkt8"];
		Json["Gewinn8"] = data[i]["Gewinn8"];
		Json["Policen-Nr.9"] = data[i]["Policen-Nr.9"];
		Json["Produkt9"] = data[i]["Produkt9"];
		Json["Gewinn9"] = data[i]["Gewinn9"];
		Json["Policen-Nr.10"] = data[i]["Policen-Nr.10"];
		Json["Produkt20"] = data[i]["Produkt20"];
		Json["Gewinn20"] = data[i]["Gewinn20"];
		Json["Policen-Nr.11"] = data[i]["Policen-Nr.11"];
		Json["Produkt11"] = data[i]["Produkt11"];
		Json["Gewinn11"] = data[i]["Gewinn11"];
		Json["Policen-Nr.12"] = data[i]["Policen-Nr.12"];
		Json["Produkt12"] = data[i]["Produkt12"];
		Json["Gewinn12"] = data[i]["Gewinn12"];
		Json["Policen-Nr.13"] = data[i]["Policen-Nr.13"];
		Json["Produkt13"] = data[i]["Produkt13"];
		Json["Gewinn13"] = data[i]["Gewinn13"];
		Json["Policen-Nr.14"] = data[i]["Policen-Nr.14"];
		Json["Produkt14"] = data[i]["Produkt14"];
		Json["Gewinn14"] = data[i]["Gewinn14"];
		Json["Policen-Nr.15"] = data[i]["Policen-Nr.15"];
		Json["Produkt15"] = data[i]["Produkt15"];
		Json["Gewinn15"] = data[i]["Gewinn15"];
		Json["Policen-Nr.16"] = data[i]["Policen-Nr.16"];
		Json["Produkt16"] = data[i]["Produkt16"];
		Json["Gewinn16"] = data[i]["Gewinn16"];
		Json["Policen-Nr.17"] = data[i]["Policen-Nr.17"];
		Json["Produkt17"] = data[i]["Produkt17"];
		Json["Gewinn17"] = data[i]["Gewinn17"];
		Json["Policen-Nr.18"] = data[i]["Policen-Nr.18"];
		Json["Produkt18"] = data[i]["Produkt18"];
		Json["Gewinn18"] = data[i]["Gewinn18"];
		Json["Policen-Nr.19"] = data[i]["Policen-Nr.19"];
		Json["Produkt19"] = data[i]["Produkt19"];
		Json["Gewinn19"] = data[i]["Gewinn19"];
		Json["Policen-Nr.20"] = data[i]["Policen-Nr.20"];
		Json["Produkt20"] = data[i]["Produkt20"];
		Json["Gewinn20"] = data[i]["Gewinn20"];
		Json["Policen-Nr.22"] = data[i]["Policen-Nr.22"];
		Json["Produkt22"] = data[i]["Produkt22"];
		Json["Gewinn22"] = data[i]["Gewinn22"];
		Json["Policen-Nr.22"] = data[i]["Policen-Nr.22"];
		Json["Produkt22"] = data[i]["Produkt22"];
		Json["Gewinn22"] = data[i]["Gewinn22"];
		Json["Policen-Nr.23"] = data[i]["Policen-Nr.23"];
		Json["Produkt23"] = data[i]["Produkt23"];
		Json["Gewinn23"] = data[i]["Gewinn23"];
		Json["Policen-Nr.24"] = data[i]["Policen-Nr.24"];
		Json["Produkt24"] = data[i]["Produkt24"];
		Json["Gewinn24"] = data[i]["Gewinn24"];
		Json["Policen-Nr.25"] = data[i]["Policen-Nr.25"];
		Json["Produkt25"] = data[i]["Produkt25"];
		Json["Gewinn25"] = data[i]["Gewinn25"];
		Json["Summe Gewinn"] = data[i]["Summe Gewinn"];
		dataXlsx.push(Json);
		dataCsv.push(Json);
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
	const XPORT = document.getElementById('exportBtn');
	XPORT.disabled = false;
	createTxt(data);
	createXlsxAndCsv(data);	

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
	for (var i = 0; i < o.filePaths.length; i++) { 
		combine(XLSX.readFile(o.filePaths[i]))
	}
	const XPORT = document.getElementById('exportBtn');
	XPORT.disabled = false;
	createTxt(data);
	createXlsxAndCsv(data);	
}
	
const choosePath = async function() {
	const o = await electron.dialog.showOpenDialog({
		title: 'Select Export Directory',
		properties: ['openDirectory']
	});
	dir = o.filePaths;
};

const exportXlsx = function() {
	var filename = document.getElementById('filename');
	var dataTxtTemp1 = JSON.stringify(dataTxt);
	var dataTxtTemp2 = dataTxtTemp1.substring(1);
	var dataTxtTemp3 = dataTxtTemp2.slice(0, -1);

	fs.writeFile(
		path.join(dir.toString(), filename.value+".txt"), dataTxtTemp3, (err) => {
		if(err) throw err;
	});
	var wb_xslx = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb_xslx, XLSX.utils.json_to_sheet(dataXlsx))
	XLSX.writeFile(wb_xslx, path.join(dir.toString(), filename.value+".xlsx"));



	var wb_csv = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb_csv, XLSX.utils.json_to_sheet(dataCsv))
	XLSX.writeFile(wb_csv, path.join(dir.toString(), filename.value+".csv"));

	electron.dialog.showMessageBox(
		{
			type: 'info',
			buttons: ['OK'],
			defaultId: 2,
			title: 'Info',
			message: 'Export successful',
		}
	);
};

// add event listeners
const readBtn = document.getElementById('readBtn');
const chooseBtn = document.getElementById('pathBtn');
const exportBtn = document.getElementById('exportBtn');
const drop = document.getElementById('drop');

readBtn.addEventListener('click', handleReadBtn, false);
chooseBtn.addEventListener('click', choosePath, false);
exportBtn.addEventListener('click', exportXlsx, false);

drop.addEventListener('dragenter', (e) => {
	e.stopPropagation();
	e.preventDefault();
}, false);
drop.addEventListener('dragover', (e) => {
	e.stopPropagation();
	e.preventDefault();
}, false);
drop.addEventListener('drop', (e) => {
	e.stopPropagation();
	e.preventDefault();
	for (const f of e.dataTransfer.files) { 
		combine(XLSX.readFile(f.path))
	} 
	const XPORT = document.getElementById('exportBtn');
	XPORT.disabled = false;
	createTxt(data);
	createXlsxAndCsv(data);	
}, false);

exportBtn.disabled = true;