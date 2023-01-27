const Exceljs = require("exceljs");
const fs = require("fs");
const path = require("path");
const axios = require("axios");
const _ = require("lodash");
const bluebird = require("bluebird");
const getLanguage = (text) => {
  return new Promise((resolve) => {
    import("franc").then(({ franc }) => resolve(franc(text)));
  });
};
const excel = new Exceljs.Workbook();
const createExcelInstance = new Exceljs.Workbook();
const translateText = async (text, mode) => {
  try {
    const returnData = await axios.get(
      `https://translate.googleapis.com/translate_a/single?client=gtx&dt=t&sl=${mode}&tl=zh-CN&q=${text}`
    );
    return _.get(returnData, "data.0", [])
      ?.map((v) => v[0])
      ?.join("");
  } catch (error) {
    return text
  }
};
const main = async () => {
  await excel.xlsx.readFile(path.join(__dirname, "translate.xlsx"));
  const sheetList = excel.worksheets;
  sheetList.forEach((sheet) => {
    if (/^\d+$/g.test(sheet.name)) {
      const createNewSheet = createExcelInstance.addWorksheet(sheet.name);
      let textArr = [];
      sheet.eachRow((row, rowNum) => {
        let arr = row.values.map((v) => String(v));
        if (rowNum === 1) {
          createNewSheet.addRow(row.values);
        } else {
          const needTranslateText = row.getCell("I").value;
          textArr.push({ row, needTranslateText, arr });
        }
      });
      bluebird
        .map(
          textArr,
          async (v) => {
            const language = await getLanguage(v.needTranslateText);
            let mode = "en";
            if (language === "jpn") {
              mode = "ja";
              console.log("----->", v.needTranslateText);
            }

            const text = await translateText(v.needTranslateText, mode);
            createNewSheet.addRow(v.arr.concat(text));
            console.log(text);
            return text;
          },
          { concurrency: 3 }
        )
        .then(() => {
          createExcelInstance.xlsx.writeFile(`qtcreator-target.xlsx`)
        });
    }
  });
};
main();
