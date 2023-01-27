const Exceljs = require("exceljs");
const fs = require("fs");
const path = require("path");
const axios = require("axios");
const _ = require("lodash");

const oldExcel = new Exceljs.Workbook();
const newExcel = new Exceljs.Workbook();
const createExcelInstance = new Exceljs.Workbook();
// 将旧文件所有项写成JSON
const writeJSON = async () => {
  await oldExcel.xlsx.readFile(path.join(__dirname, "oldExcel.xlsx"));
  const sheetList = oldExcel.worksheets;
  sheetList.map((sheet) => {
    let obj = {};
    if (/^\d+$/g.test(sheet.name)) {
      sheet.eachRow((row, rowNum) => {
        let flag = row.getCell("O").value;
        if (sheet.name === "2017") {
          flag = row.getCell("P").value;
        }
        const hash = row.getCell("B").value;
        obj[hash] = flag;
      });
    }
    fs.writeFileSync(
      path.join(__dirname, `${sheet.name}.json`),
      JSON.stringify(obj, null, 2)
    );
  });
};
// 将JSON写入源文件
const writeOriginFile = async () => {
  await newExcel.xlsx.readFile(path.join(__dirname, "newExcel.xlsx"));
  const sheetList = newExcel.worksheets;
  sheetList.map((sheet) => {
    let count = 0;
    let haveHandleData = 0;
    if (/^\d+$/g.test(sheet.name)) {
      const createNewSheet = createExcelInstance.addWorksheet(sheet.name);
      const data = require(path.join(__dirname, `${sheet.name}.json`));
      sheet.eachRow((row, rowNum) => {
        const key = row.getCell("B").value;
        const flag = data[key];
        let arr = row.values.map((v) => String(v));
        if (flag) {
          console.log(`${sheet.name}中${key}命中`, ++count);
          createNewSheet.addRow(arr.concat(flag));
        } else {
          createNewSheet.addRow(row.values);
        }
        haveHandleData = rowNum;
      });
      console.log(`${sheet.name}处理完，共处理`, haveHandleData);
    }
  });
};
const sleep = (time) => {
  return new Promise((resolve, reject) => {
    setTimeout(resolve, time);
  });
};
const main = async () => {
  await writeJSON();
  // 由于写文件 sleep一会儿
  await sleep(500);
  await writeOriginFile();
  await createExcelInstance.xlsx.writeFile(
    path.join(__dirname, `nova-target.xlsx`)
  );
};

main();
