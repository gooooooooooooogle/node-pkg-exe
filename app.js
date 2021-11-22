const xlsx = require('node-xlsx')
const fs = require('fs')


doTest();

function doTest (){
	const workSheetsFromFile = xlsx.parse('./resource/data.xls');

	const dataSource = workSheetsFromFile[0].data;

	delFiles('./result.xls');

	let saveData = [];
	let headData = ['表号', '信号A', '信号B'];

	saveData.push(headData);

	for (let index = 1; index < dataSource.length; index++) {
	  const row = dataSource[index];
	  if (row.length >= 7) {
		const meterAddr = row[1];
		const col = row[7];
		if (col.length >= 74) {
			let a = col.substring(58, 60);
			let b = col.substring(62, 64);
			let inta = parseInt(a, 16);
			let intb = parseInt(b, 16);
			let rowData = [];
			rowData[0] = meterAddr;
			rowData[1] = inta;
			rowData[2] = intb;
			
			saveData.push(rowData);
		}
	  }
	}

	writeXls(saveData);
}

function delFiles(filePath) {
	fs.unlink(filePath, () => {})
}

function writeXls(datas) {
	const options = {'!cols': [{ wch: 18 }, { wch: 10 }, { wch: 10 } ]};
    let buffer = xlsx.build([
        {
            name:'sheet1',
            data:datas   
        }
    ], options);
	fs.writeFile('./result.xls', buffer, { 'flag': 'w+' }, (err) => {
		if (err) return 
	})
}