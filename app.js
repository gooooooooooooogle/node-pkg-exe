const xlsx = require('node-xlsx')
const fs = require('fs')

const workSheetsFromFile = xlsx.parse('./resource/data.xls');

const dataSource = workSheetsFromFile[1].data;

for (let index = 1; index < dataSource.length; index++) {
  const row = dataSource[index];
  if (row.length >= 7) {
    const col = row[7] || '';
    let a = col.substring(58, 60);
    let b = col.substring(62, 64);
    a = parseInt(a, 16)
    b = parseInt(b, 16)
    if ((a > 160) || (a < 80)) {
      // console.log('信号A：', a);
      writeFiles('信号A：' +  a + '\n')
    }
    if ((b > 160) || (b < 80)) {
      // console.log('信号B：', b);
      writeFiles('信号B：' + b + '\n')
    }
  }
}

console.log('----------------处理结束！请查看当前目录下result.txt文件----------------')

function writeFiles(info) {
  fs.appendFile('./result.txt', info, () => {})
}