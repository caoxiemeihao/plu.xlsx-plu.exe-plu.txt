const fs = require('fs');
const path = require('path');
const os = require('os');
const iconv = require('iconv-lite');
const xlsx = require('xlsx');

const date = new Date();
const padLeft0 = (str = '') => String(str).length === 1 ? String(str).padStart(2, 0) : String(str);
const fieldDict = {
  A: 'ID', // 生鲜码
  B: 'ItemCode', // 货号
  C: 'DepartmentID', // 条码起始码
  D: 'Name1', // 名称1
  E: 'Price', // 单价
  F: 'UnitID', // 单位
  G: 'BarcodeType1', // 条码类型1
  H: 'BarcodeType2', // 条码类型2
  I: 'ProducedDate', // 生产日期
};

// 重量单位，7(500g)，4(kg)，10(PCS/kg)
const UnitID_type = { '500g': 7, kg: 4, PCS: 10, };
// 条码起始码，写 24
const DepartmentID = 24;
// 生产日期
const ProducedDate = `${
  date.getFullYear()
  }/${
  date.getMonth() + 1
  }/${
  date.getDate()
  } ${
  padLeft0(date.getHours())
  }:${
  padLeft0(date.getMinutes())
  }:${
  padLeft0(date.getSeconds())
  }`;

// const desktop = path.join(os.homedir(), 'Desktop');
const desktop = __dirname;
let error = '';
let sheetJsonData;

try {
  const workbook = xlsx.readFile(path.join(desktop, 'plu.xlsx'));
  const firstSheetName = workbook.SheetNames[0];
  sheetJsonData = workbook.Sheets[firstSheetName];
  // console.log(sheetJsonData);
} catch (e) {
  error = e;
}

/** 解析 xlsx */
function xlsx2plu(sheetJsonData) {
  let message = '';
  const arr = [];
  try {
    const sheetJsonData2 = {};
    Object.entries(sheetJsonData).forEach(([k, v]) => {
      const isData = /^[A-Z]\d+$/.test(k);
      if (!isData) { return; }
      const rowKey = k.substr(1);
      if (!(sheetJsonData2[rowKey] instanceof Object)) {
        sheetJsonData2[rowKey] = {};
      }
      const filed = fieldDict[k.charAt(0)];
      sheetJsonData2[rowKey][filed] = v.v;
    });
    // console.log(sheetJsonData2);

    Object.entries(sheetJsonData2).forEach(([k, v]) => {
      if (typeof v.ID === 'string') { return; } // 跳过标题行
      arr.push(Object.assign(v, { ProducedDate }));
    });
  } catch (e) {
    message = e;
  }
  return [message, arr];
}

/** 生成 xxx.txt */
function generatePlu(file_path, plu_arr) {
  let message = ''
  try {
    const data = `ID\tItemCode\tDepartmentID\tName1\tPrice\tUnitID\tBarcodeType1\tBarcodeType2\tProducedDate
${plu_arr.map(item => `${
      item.ID
      }\t${
      item.ItemCode
      }\t${
      item.DepartmentID
      }\t${
      item.Name1
      }\t${
      item.Price
      }\t${
      item.UnitID
      }\t${
      item.BarcodeType1
      }\t${
      item.BarcodeType2
      }\t${
      item.ProducedDate
      }`).join('\n')}`
    fs.writeFileSync(file_path, iconv.encode(data, 'GBK'))
  } catch (e) {
    message = e
  }
  return message
}

let msg, arr;
if (!error) {
  [msg, arr] = xlsx2plu(sheetJsonData);
  error = msg;
}
// console.log(msg);
// console.log(arr);
if (!error) {
  error = generatePlu(path.join(desktop, 'plu.txt'), arr);
}


console.log('--------------------');
const delay = error ? 20 : 4.999;
if (error) {
  console.log(error);
} else {
  console.log('生成成功', path.join(desktop, 'plu.txt'));
}
console.log(`[${ProducedDate}] ${~~delay} 秒后程序自动退出`);
setTimeout(() => { }, 1000 * delay);
