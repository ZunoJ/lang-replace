#! /usr/bin/env node

const fs = require("fs");
const glob = require("glob");
const path = require("path");
const ExcelJS = require("exceljs");

const projectPath = path.join(__dirname, "../../../"); // 前端项目目录
const excelFilePath = `./${process.env.NAME || "doc"}.xlsx`; // Excel文件的路径，根据你的实际路径进行修改
let keyIndex;
let newKeyIndex;
// 读取 Excel 文件
async function readExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheet = workbook.getWorksheet(1);
  const keyMapping = {};

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber !== 1) {
      const key = row.getCell(keyIndex).value;
      const newKey = row.getCell(newKeyIndex).value;
      keyMapping[key] = newKey;
    } else {
      keyIndex = row.values.indexOf("key");
      newKeyIndex = row.values.indexOf("newkey");
    }
  });
  return keyMapping;
}
// 定义匹配$T方法调用的正则表达式
const regex = /(\$[Tt])\(['"]([^'"]+)['"]\)/g;
// 使用glob模块找到所有Vue和JS文件
const files = glob.sync("**/*.+(vue|js)", {
  cwd: projectPath,
  ignore: "**/node_modules/**", // 忽略node_modules文件夹
});
// 获取 key 映射
readExcel(excelFilePath).then((res) => {
  // 遍历所有文件
  files.forEach((file) => {
    const content = fs.readFileSync(file, "utf-8");

    // 使用正则表达式匹配$T方法调用并进行替换
    const modifiedContent = content.replace(regex, (...args) => {
      const $0 = args[0];
      const funcName = args[1];
      const oldParam = args[2];
      // 如果有对应的 newKey，则替换
      if (res.hasOwnProperty(oldParam)) {
        const newKey = res[oldParam];
        return `${funcName}('${newKey}')`;
      }
      return $0; // 没有对应的 newKey，保持原样
    });

    // 将修改后的内容写回文件
    fs.writeFileSync(file, modifiedContent, "utf-8");
  });
  console.log("Replacement completed.");
});
