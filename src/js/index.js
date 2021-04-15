const XLSX = require('xlsx');
const csv = require('csvtojson')
const electron = require('electron').remote;
const iconv = require('iconv-lite');
const fs = require('fs') ;
const rimraf = require("rimraf");
const fsPromis = fs.promises;
const path = require('path');

let dataInput = [];
const dataTxt = [];
const dataXlsx = [];
const dataCsv = [];
let dir = "";

iconv.skipDecodeWarning = true;

function combine(file) {

	for (let row of file) {
		var newKey = 'Summe Gewinn'
		var oldKey = Object.keys(row)[Object.keys(row).length-1];
		delete Object.assign(row, {[newKey]: row[oldKey] })[oldKey];
		dataInput.push(row);
	}
};

function createTxt(data){
	for(var i in data){
		if(!dataTxt.includes(data[i]["Vermittler"])){
			dataTxt.push(data[i]["Vermittler"]);
		}
	}
	
}

function createXlsxAndCsv(data){
	for(i in data) {
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
		Json["Policen-Nr.1"] = data[i]["Policen-Nr1"];
		Json["Produkt1"] = data[i]["Produkt1"];
		Json["Gewinn1"] = data[i]["Gewinn1"];
		Json["Policen-Nr.2"] = data[i]["Policen-Nr2"];
		Json["Produkt2"] = data[i]["Produkt2"];
		Json["Gewinn2"] = data[i]["Gewinn2"];
		Json["Policen-Nr.3"] = data[i]["Policen-Nr3"];
		Json["Produkt3"] = data[i]["Produkt3"];
		Json["Gewinn3"] = data[i]["Gewinn3"];
		Json["Policen-Nr.4"] = data[i]["Policen-Nr4"];
		Json["Produkt4"] = data[i]["Produkt4"];
		Json["Gewinn4"] = data[i]["Gewinn4"];
		Json["Policen-Nr.5"] = data[i]["Policen-Nr5"];
		Json["Produkt5"] = data[i]["Produkt5"];
		Json["Gewinn5"] = data[i]["Gewinn5"];
		Json["Policen-Nr.6"] = data[i]["Policen-Nr6"];
		Json["Produkt6"] = data[i]["Produkt6"];
		Json["Gewinn6"] = data[i]["Gewinn6"];
		Json["Policen-Nr.7"] = data[i]["Policen-Nr7"];
		Json["Produkt7"] = data[i]["Produkt7"];
		Json["Gewinn7"] = data[i]["Gewinn7"];
		Json["Policen-Nr.8"] = data[i]["Policen-Nr8"];
		Json["Produkt8"] = data[i]["Produkt8"];
		Json["Gewinn8"] = data[i]["Gewinn8"];
		Json["Policen-Nr.9"] = data[i]["Policen-Nr9"];
		Json["Produkt9"] = data[i]["Produkt9"];
		Json["Gewinn9"] = data[i]["Gewinn9"];
		Json["Policen-Nr.10"] = data[i]["Policen-Nr10"];
		Json["Produkt10"] = data[i]["Produkt10"];
		Json["Gewinn10"] = data[i]["Gewinn10"];
		Json["Policen-Nr.11"] = data[i]["Policen-Nr11"];
		Json["Produkt11"] = data[i]["Produkt11"];
		Json["Gewinn11"] = data[i]["Gewinn11"];
		Json["Policen-Nr.12"] = data[i]["Policen-Nr12"];
		Json["Produkt12"] = data[i]["Produkt12"];
		Json["Gewinn12"] = data[i]["Gewinn12"];
		Json["Policen-Nr.13"] = data[i]["Policen-Nr13"];
		Json["Produkt13"] = data[i]["Produkt13"];
		Json["Gewinn13"] = data[i]["Gewinn13"];
		Json["Policen-Nr.14"] = data[i]["Policen-Nr14"];
		Json["Produkt14"] = data[i]["Produkt14"];
		Json["Gewinn14"] = data[i]["Gewinn14"];
		Json["Policen-Nr.15"] = data[i]["Policen-Nr15"];
		Json["Produkt15"] = data[i]["Produkt15"];
		Json["Gewinn15"] = data[i]["Gewinn15"];
		Json["Policen-Nr.16"] = data[i]["Policen-Nr16"];
		Json["Produkt16"] = data[i]["Produkt16"];
		Json["Gewinn16"] = data[i]["Gewinn16"];
		Json["Policen-Nr.17"] = data[i]["Policen-Nr17"];
		Json["Produkt17"] = data[i]["Produkt17"];
		Json["Gewinn17"] = data[i]["Gewinn17"];
		Json["Policen-Nr.18"] = data[i]["Policen-Nr18"];
		Json["Produkt18"] = data[i]["Produkt18"];
		Json["Gewinn18"] = data[i]["Gewinn18"];
		Json["Policen-Nr.19"] = data[i]["Policen-Nr19"];
		Json["Produkt19"] = data[i]["Produkt19"];
		Json["Gewinn19"] = data[i]["Gewinn19"];
		Json["Policen-Nr.20"] = data[i]["Policen-Nr20"];
		Json["Produkt20"] = data[i]["Produkt20"];
		Json["Gewinn20"] = data[i]["Gewinn20"];
		Json["Policen-Nr.21"] = data[i]["Policen-Nr21"];
		Json["Produkt21"] = data[i]["Produkt21"];
		Json["Gewinn21"] = data[i]["Gewinn21"];
		Json["Policen-Nr.22"] = data[i]["Policen-Nr22"];
		Json["Produkt22"] = data[i]["Produkt22"];
		Json["Gewinn22"] = data[i]["Gewinn22"];
		Json["Policen-Nr.23"] = data[i]["Policen-Nr23"];
		Json["Produkt23"] = data[i]["Produkt23"];
		Json["Gewinn23"] = data[i]["Gewinn23"];
		Json["Policen-Nr.24"] = data[i]["Policen-Nr24"];
		Json["Produkt24"] = data[i]["Produkt24"];
		Json["Gewinn24"] = data[i]["Gewinn24"];
		Json["Policen-Nr.25"] = data[i]["Policen-Nr25"];
		Json["Produkt25"] = data[i]["Produkt25"];
		Json["Gewinn25"] = data[i]["Gewinn25"];
		Json["Summe Gewinn"] = data[i]["Summe Gewinn"];
		dataXlsx.push(Json);
		dataCsv.push(Json);
	}
}

async function handleReadBtn() {
	dataInput = [];
	const o = await electron.dialog.showOpenDialog({
		title: 'Select a file',
		filters: [{
			name: "Spreadsheets",
			extensions: 'csv'
		}],
		properties: ['openFile', 'multiSelections']
	});
	if(o != null){
		for (var i = 0; i < o.filePaths.length; i++) { 
			var file = o.filePaths[i];
			var filename = path.basename(file, '.csv');
			var extname = path.extname(file);
			var dirname = path.dirname(file);
			var decodeFile = path.join(dirname, "decode",filename+"_decode"+extname);
			if (!fs.existsSync(path.join(dirname,"decode"))){
				fs.mkdirSync(path.join(dirname,"decode"));
			}
			var input = fs.readFileSync(file, {encoding: "binary"});
			var output = iconv.decode(input, "ISO-8859-1");
			fs.writeFileSync(decodeFile, output);
			await csv({
				delimiter:";",
				noheader: false,
				ignoreEmpty:true,
				headers:['VM-Policeninfo','Vermittler','Quittungs-Nr','Sprache','Anrede','Name','Vorname','Name2','Strasse','Postfach','PLZ/Ort','Policen-Nr1','Produkt1','Gewinn1','Policen-Nr2','Produkt2','Gewinn2','Policen-Nr3','Produkt3','Gewinn3','Policen-Nr4','Produkt4','Gewinn4','Policen-Nr5','Produkt5','Gewinn5','Policen-Nr6','Produkt6','Gewinn6','Policen-Nr7','Produkt7','Gewinn7','Policen-Nr8','Produkt8','Gewinn8','Policen-Nr9','Produkt9','Gewinn9','Policen-Nr10','Produkt10','Gewinn10','Policen-Nr11','Produkt11','Gewinn11','Policen-Nr12','Produkt12','Gewinn12','Policen-Nr13','Produkt13','Gewinn13','Policen-Nr14','Produkt14','Gewinn14','Policen-Nr15','Produkt15','Gewinn15','Policen-Nr16','Produkt16','Gewinn16','Policen-Nr17','Produkt17','Gewinn17','Policen-Nr18','Produkt18','Gewinn18','Policen-Nr19','Produkt19','Gewinn19','Policen-Nr20','Produkt20','Gewinn20','Policen-Nr21','Produkt21','Gewinn21','Policen-Nr22','Produkt22','Gewinn22','Policen-Nr23','Produkt23','Gewinn23','Policen-Nr24','Produkt24','Gewinn24','Policen-Nr25','Produkt25','Gewinn25','Summe Gewinn']
			}).fromFile(decodeFile).then(source => {
				combine(source);
			})
		}
		if (fs.existsSync(path.join(dirname,"decode"))){
			fs.rmdirSync(path.join(dirname,"decode"),{recursive: true});
		}
		const XPORT = document.getElementById('exportBtn');
		XPORT.disabled = false;
		createTxt(dataInput);
		createXlsxAndCsv(dataInput);
	}
	
	
}
	
async function choosePath() {
	const o = await electron.dialog.showOpenDialog({
		title: 'Select Export Directory',
		properties: ['openDirectory']
	});
	dir = o.filePaths;
};

function exportXlsx() {	
	var filename = document.getElementById('filename');
	var dataTxtTemp1 = JSON.stringify(dataTxt);
	var dataTxtTemp2 = dataTxtTemp1.substring(1);
	var dataTxtTemp3 = dataTxtTemp2.slice(0, -1);
	var dataTxtTemp4 = dataTxtTemp3.replace(/"/g, '');

	fs.writeFile(
		path.join(dir.toString(), filename.value+".txt"), dataTxtTemp4, (err) => {
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
drop.addEventListener('drop', async (e) => {
	dataInput = [];
	e.stopPropagation();
	e.preventDefault();
	for (const f of e.dataTransfer.files) {
		var file = f.path;
		var filename = path.basename(file, '.csv');
		var extname = path.extname(file);
		var dirname = path.dirname(file);
		var decodeFile = path.join(dirname, "decode",filename+"_decode"+extname);
		if (!fs.existsSync(path.join(dirname,"decode"))){
			fs.mkdirSync(path.join(dirname,"decode"));
		}
		var input = fs.readFileSync(file, {encoding: "binary"});
		var output = iconv.decode(input, "ISO-8859-1");
		fs.writeFileSync(decodeFile, output);
		await csv({
			delimiter:";",
			noheader: false,
			ignoreEmpty:true,
			headers:['VM-Policeninfo','Vermittler','Quittungs-Nr','Sprache','Anrede','Name','Vorname','Name2','Strasse','Postfach','PLZ/Ort','Policen-Nr1','Produkt1','Gewinn1','Policen-Nr2','Produkt2','Gewinn2','Policen-Nr3','Produkt3','Gewinn3','Policen-Nr4','Produkt4','Gewinn4','Policen-Nr5','Produkt5','Gewinn5','Policen-Nr6','Produkt6','Gewinn6','Policen-Nr7','Produkt7','Gewinn7','Policen-Nr8','Produkt8','Gewinn8','Policen-Nr9','Produkt9','Gewinn9','Policen-Nr10','Produkt10','Gewinn10','Policen-Nr11','Produkt11','Gewinn11','Policen-Nr12','Produkt12','Gewinn12','Policen-Nr13','Produkt13','Gewinn13','Policen-Nr14','Produkt14','Gewinn14','Policen-Nr15','Produkt15','Gewinn15','Policen-Nr16','Produkt16','Gewinn16','Policen-Nr17','Produkt17','Gewinn17','Policen-Nr18','Produkt18','Gewinn18','Policen-Nr19','Produkt19','Gewinn19','Policen-Nr20','Produkt20','Gewinn20','Policen-Nr21','Produkt21','Gewinn21','Policen-Nr22','Produkt22','Gewinn22','Policen-Nr23','Produkt23','Gewinn23','Policen-Nr24','Produkt24','Gewinn24','Policen-Nr25','Produkt25','Gewinn25','Summe Gewinn']
		}).fromFile(decodeFile).then(source => {
			combine(source);
		})
		
	}
	if (fs.existsSync(path.join(dirname,"decode"))){
		fs.rmdirSync(path.join(dirname,"decode"),{recursive: true});
	}
	const XPORT = document.getElementById('exportBtn');
	XPORT.disabled = false;
	createTxt(dataInput);
	createXlsxAndCsv(dataInput);	
}, false);

exportBtn.disabled = true;