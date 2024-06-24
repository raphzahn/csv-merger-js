const XLSX = require('xlsx');
const csv = require('csvtojson')
const electron = require('electron').remote;
const iconv = require('iconv-lite');
const fs = require('fs') ;
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
		Json["Quittungs-Nr."] = data[i]["Quittungs-Nr"];
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
		Json["Policen-Nr.26"] = data[i]["Policen-Nr26"];
		Json["Produkt26"] = data[i]["Produkt26"];
		Json["Gewinn26"] = data[i]["Gewinn26"];
		Json["Policen-Nr.27"] = data[i]["Policen-Nr27"];
		Json["Produkt27"] = data[i]["Produkt27"];
		Json["Gewinn27"] = data[i]["Gewinn27"];
		Json["Policen-Nr.28"] = data[i]["Policen-Nr28"];
		Json["Produkt28"] = data[i]["Produkt28"];
		Json["Gewinn28"] = data[i]["Gewinn28"];
		Json["Policen-Nr.29"] = data[i]["Policen-Nr29"];
		Json["Produkt29"] = data[i]["Produkt29"];
		Json["Gewinn29"] = data[i]["Gewinn29"];
		Json["Policen-Nr.30"] = data[i]["Policen-Nr30"];
		Json["Produkt30"] = data[i]["Produkt30"];
		Json["Gewinn30"] = data[i]["Gewinn30"];
		Json["Policen-Nr.31"] = data[i]["Policen-Nr31"];
		Json["Produkt31"] = data[i]["Produkt31"];
		Json["Gewinn31"] = data[i]["Gewinn31"];
		Json["Policen-Nr.32"] = data[i]["Policen-Nr32"];
		Json["Produkt32"] = data[i]["Produkt32"];
		Json["Gewinn32"] = data[i]["Gewinn32"];
		Json["Policen-Nr.33"] = data[i]["Policen-Nr33"];
		Json["Produkt33"] = data[i]["Produkt33"];
		Json["Gewinn33"] = data[i]["Gewinn33"];
		Json["Policen-Nr.34"] = data[i]["Policen-Nr34"];
		Json["Produkt34"] = data[i]["Produkt34"];
		Json["Gewinn34"] = data[i]["Gewinn34"];
		Json["Policen-Nr.35"] = data[i]["Policen-Nr35"];
		Json["Produkt35"] = data[i]["Produkt35"];
		Json["Gewinn35"] = data[i]["Gewinn35"];
		Json["Policen-Nr.36"] = data[i]["Policen-Nr36"];
		Json["Produkt36"] = data[i]["Produkt36"];
		Json["Gewinn36"] = data[i]["Gewinn36"];
		Json["Policen-Nr.37"] = data[i]["Policen-Nr37"];
		Json["Produkt37"] = data[i]["Produkt37"];
		Json["Gewinn37"] = data[i]["Gewinn37"];
		Json["Policen-Nr.38"] = data[i]["Policen-Nr38"];
		Json["Produkt38"] = data[i]["Produkt38"];
		Json["Gewinn38"] = data[i]["Gewinn38"];
		Json["Policen-Nr.39"] = data[i]["Policen-Nr39"];
		Json["Produkt39"] = data[i]["Produkt39"];
		Json["Gewinn39"] = data[i]["Gewinn39"];
		Json["Policen-Nr.40"] = data[i]["Policen-Nr40"];
		Json["Produkt40"] = data[i]["Produkt40"];
		Json["Gewinn40"] = data[i]["Gewinn40"];
		Json["Policen-Nr.41"] = data[i]["Policen-Nr41"];
		Json["Produkt41"] = data[i]["Produkt41"];
		Json["Gewinn41"] = data[i]["Gewinn41"];
		Json["Policen-Nr.42"] = data[i]["Policen-Nr42"];
		Json["Produkt42"] = data[i]["Produkt42"];
		Json["Gewinn42"] = data[i]["Gewinn42"];
		Json["Policen-Nr.43"] = data[i]["Policen-Nr43"];
		Json["Produkt43"] = data[i]["Produkt43"];
		Json["Gewinn43"] = data[i]["Gewinn43"];
		Json["Policen-Nr.44"] = data[i]["Policen-Nr44"];
		Json["Produkt44"] = data[i]["Produkt44"];
		Json["Gewinn44"] = data[i]["Gewinn44"];
		Json["Policen-Nr.45"] = data[i]["Policen-Nr45"];
		Json["Produkt45"] = data[i]["Produkt45"];
		Json["Gewinn45"] = data[i]["Gewinn45"];
		Json["Policen-Nr.46"] = data[i]["Policen-Nr46"];
		Json["Produkt46"] = data[i]["Produkt46"];
		Json["Gewinn46"] = data[i]["Gewinn46"];
		Json["Policen-Nr.47"] = data[i]["Policen-Nr47"];
		Json["Produkt47"] = data[i]["Produkt47"];
		Json["Gewinn47"] = data[i]["Gewinn47"];
		Json["Policen-Nr.48"] = data[i]["Policen-Nr48"];
		Json["Produkt48"] = data[i]["Produkt48"];
		Json["Gewinn48"] = data[i]["Gewinn48"];
		Json["Policen-Nr.49"] = data[i]["Policen-Nr49"];
		Json["Produkt49"] = data[i]["Produkt49"];
		Json["Gewinn49"] = data[i]["Gewinn49"];
		Json["Policen-Nr.50"] = data[i]["Policen-Nr50"];
		Json["Produkt50"] = data[i]["Produkt50"];
		Json["Gewinn50"] = data[i]["Gewinn50"];
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
				headers:['VM-Policeninfo','Vermittler','Quittungs-Nr','Sprache','Anrede','Name','Vorname','Name2','Strasse','Postfach','PLZ/Ort','Policen-Nr1','Produkt1','Gewinn1','Policen-Nr2','Produkt2','Gewinn2','Policen-Nr3','Produkt3','Gewinn3','Policen-Nr4','Produkt4','Gewinn4','Policen-Nr5','Produkt5','Gewinn5','Policen-Nr6','Produkt6','Gewinn6','Policen-Nr7','Produkt7','Gewinn7','Policen-Nr8','Produkt8','Gewinn8','Policen-Nr9','Produkt9','Gewinn9','Policen-Nr10','Produkt10','Gewinn10','Policen-Nr11','Produkt11','Gewinn11','Policen-Nr12','Produkt12','Gewinn12','Policen-Nr13','Produkt13','Gewinn13','Policen-Nr14','Produkt14','Gewinn14','Policen-Nr15','Produkt15','Gewinn15','Policen-Nr16','Produkt16','Gewinn16','Policen-Nr17','Produkt17','Gewinn17','Policen-Nr18','Produkt18','Gewinn18','Policen-Nr19','Produkt19','Gewinn19','Policen-Nr20','Produkt20','Gewinn20','Policen-Nr21','Produkt21','Gewinn21','Policen-Nr22','Produkt22','Gewinn22','Policen-Nr23','Produkt23','Gewinn23','Policen-Nr24','Produkt24','Gewinn24','Policen-Nr25','Produkt25','Gewinn25','Policen-Nr26','Produkt26','Gewinn26','Policen-Nr27','Produkt27','Gewinn27','Policen-Nr28','Produkt28','Gewinn28','Policen-Nr29','Produkt29','Gewinn29','Policen-Nr30','Produkt30','Gewinn30','Policen-Nr31','Produkt31','Gewinn31','Policen-Nr32','Produkt32','Gewinn32','Policen-Nr33','Produkt33','Gewinn33','Policen-Nr34','Produkt34','Gewinn34','Policen-Nr35','Produkt35','Gewinn35','Policen-Nr36','Produkt36','Gewinn36','Policen-Nr37','Produkt37','Gewinn37','Policen-Nr38','Produkt38','Gewinn38','Policen-Nr39','Produkt39','Gewinn39','Policen-Nr40','Produkt40','Gewinn40','Policen-Nr41','Produkt41','Gewinn41','Policen-Nr42','Produkt42','Gewinn42','Policen-Nr43','Produkt43','Gewinn43','Policen-Nr44','Produkt44','Gewinn44','Policen-Nr45','Produkt45','Gewinn45','Policen-Nr46','Produkt46','Gewinn46','Policen-Nr47','Produkt47','Gewinn47','Policen-Nr48','Produkt48','Gewinn48','Policen-Nr49','Produkt49','Gewinn49','Policen-Nr50','Produkt50','Gewinn50','Summe Gewinn']
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
			headers:['VM-Policeninfo','Vermittler','Quittungs-Nr','Sprache','Anrede','Name','Vorname','Name2','Strasse','Postfach','PLZ/Ort','Policen-Nr1','Produkt1','Gewinn1','Policen-Nr2','Produkt2','Gewinn2','Policen-Nr3','Produkt3','Gewinn3','Policen-Nr4','Produkt4','Gewinn4','Policen-Nr5','Produkt5','Gewinn5','Policen-Nr6','Produkt6','Gewinn6','Policen-Nr7','Produkt7','Gewinn7','Policen-Nr8','Produkt8','Gewinn8','Policen-Nr9','Produkt9','Gewinn9','Policen-Nr10','Produkt10','Gewinn10','Policen-Nr11','Produkt11','Gewinn11','Policen-Nr12','Produkt12','Gewinn12','Policen-Nr13','Produkt13','Gewinn13','Policen-Nr14','Produkt14','Gewinn14','Policen-Nr15','Produkt15','Gewinn15','Policen-Nr16','Produkt16','Gewinn16','Policen-Nr17','Produkt17','Gewinn17','Policen-Nr18','Produkt18','Gewinn18','Policen-Nr19','Produkt19','Gewinn19','Policen-Nr20','Produkt20','Gewinn20','Policen-Nr21','Produkt21','Gewinn21','Policen-Nr22','Produkt22','Gewinn22','Policen-Nr23','Produkt23','Gewinn23','Policen-Nr24','Produkt24','Gewinn24','Policen-Nr25','Produkt25','Gewinn25','Policen-Nr26','Produkt26','Gewinn26','Policen-Nr27','Produkt27','Gewinn27','Policen-Nr28','Produkt28','Gewinn28','Policen-Nr29','Produkt29','Gewinn29','Policen-Nr30','Produkt30','Gewinn30','Policen-Nr31','Produkt31','Gewinn31','Policen-Nr32','Produkt32','Gewinn32','Policen-Nr33','Produkt33','Gewinn33','Policen-Nr34','Produkt34','Gewinn34','Policen-Nr35','Produkt35','Gewinn35','Policen-Nr36','Produkt36','Gewinn36','Policen-Nr37','Produkt37','Gewinn37','Policen-Nr38','Produkt38','Gewinn38','Policen-Nr39','Produkt39','Gewinn39','Policen-Nr40','Produkt40','Gewinn40','Policen-Nr41','Produkt41','Gewinn41','Policen-Nr42','Produkt42','Gewinn42','Policen-Nr43','Produkt43','Gewinn43','Policen-Nr44','Produkt44','Gewinn44','Policen-Nr45','Produkt45','Gewinn45','Policen-Nr46','Produkt46','Gewinn46','Policen-Nr47','Produkt47','Gewinn47','Policen-Nr48','Produkt48','Gewinn48','Policen-Nr49','Produkt49','Gewinn49','Policen-Nr50','Produkt50','Gewinn50','Summe Gewinn']
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