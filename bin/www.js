#! /usr/bin/env node

const fs = require("fs");
const glob = require("glob");
const path = require("path");
const ExcelJS = require("exceljs");

const projectPath = path.join(__dirname, "../../../"); // 前端项目目录
const excelFilePath = `./${process.env.NAME || "doc"}.xlsx`; // 新Excel文件的路径，根据你的实际路径进行修改
const oldExcelFilePath = `./${process.env.OLDNAME || "oldDoc"}.xlsx`; // 旧Excel文件的路径，根据你的实际路径进行修改
let sheetKeyIndex;
let oldSheetKeyIndex;
let sheetenUSIndex;
let oldSheetenUSIndex;
let sheetzhCNIndex;
let oldSheetzhCNIndex;
let sheetzhHKIndex;
let oldSheetzhHKIndex;
let sheetmsMYIndex;
let oldSheetmsMYIndex;
let oldKeyValueMapping = {
  "en-US": {},
  "zh-CN": {},
  "zh-HK": {},
  "ms-MY": {},
};
let newKeyValueMapping = {
  "en-US": {},
  "zh-CN": {},
  "zh-HK": {},
  "ms-MY": {},
};
// 读取 Excel 文件
async function readExcel(filePath, oldFilePath) {
  const workbook = new ExcelJS.Workbook();
  const oldWorkbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  await oldWorkbook.xlsx.readFile(oldFilePath);

  const sheet = workbook.getWorksheet(1);
  const oldSheet = oldWorkbook.getWorksheet(1);
  const keyMapping = {};

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber !== 1) {
      const key = row.getCell(sheetKeyIndex).value;
      const enUSText = row.getCell(sheetenUSIndex).value;
      const zhCNText = row.getCell(sheetzhCNIndex).value;
      const zhHKText = row.getCell(sheetzhHKIndex).value;
      const msMYText = row.getCell(sheetmsMYIndex).value;
      newKeyValueMapping["en-US"][key] = enUSText;
      newKeyValueMapping["zh-CN"][key] = zhCNText;
      newKeyValueMapping["zh-HK"][key] = zhHKText;
      newKeyValueMapping["ms-MY"][key] = msMYText;
      oldSheet.eachRow((oldRow, oldRowNumber) => {
        if (oldRowNumber !== 1) {
          const oldKey = oldRow.getCell(oldSheetKeyIndex).value;
          const oldenUSText = oldRow.getCell(oldSheetenUSIndex).value;
          const oldzhCNText = oldRow.getCell(oldSheetzhCNIndex).value;
          const oldzhHKText = oldRow.getCell(oldSheetzhHKIndex).value;
          const oldmsMYText = oldRow.getCell(oldSheetmsMYIndex).value;
          oldKeyValueMapping["en-US"][oldKey] = oldenUSText;
          oldKeyValueMapping["zh-CN"][oldKey] = oldzhCNText;
          oldKeyValueMapping["zh-HK"][oldKey] = oldzhHKText;
          oldKeyValueMapping["ms-MY"][oldKey] = oldmsMYText;
          if (enUSText === oldenUSText) {
            keyMapping[oldKey] = key;
          }
        } else {
          oldSheetKeyIndex = oldRow.values.indexOf("key");
          oldSheetenUSIndex = oldRow.values.indexOf("英文");
          oldSheetzhCNIndex = oldRow.values.indexOf("简中");
          oldSheetzhHKIndex = oldRow.values.indexOf("繁中");
          oldSheetmsMYIndex = oldRow.values.indexOf("马来语");
        }
      });
    } else {
      sheetKeyIndex = row.values.indexOf("key");
      sheetenUSIndex = row.values.indexOf("英文");
      sheetzhCNIndex = row.values.indexOf("简中");
      sheetzhHKIndex = row.values.indexOf("繁中");
      sheetmsMYIndex = row.values.indexOf("马来语");
    }
  });
  return keyMapping;
}

/**
 * @name 生成多语言文件
 * @param {*} area
 * @param {*} oldValue
 * @param {*} newValue
 */
function generatorI18nDoc(area, oldValue, newValue) {
  // 以新的为准，合并新旧两份文件的key-value对
  const mergedKeyMapping = { ...oldValue, ...newValue };
  // 删除与旧文件中属性值相同的属性
  for (const key in oldValue) {
    const oldValueText = oldValue[key];
    if (Object.values(newValue).includes(oldValueText)) {
      delete mergedKeyMapping[key];
    }
  }
  const duplicatedValuePairs = {};

  // 遍历 mergedKeyMapping，查找值相同但 key 不同的情况
  for (const key in mergedKeyMapping) {
    const value = mergedKeyMapping[key];
    const foundKey = Object.keys(mergedKeyMapping).find(
      (otherKey) => otherKey !== key && mergedKeyMapping[otherKey] === value
    );
    if (foundKey && key !== foundKey) {
      const sortedKeys = [key, foundKey].sort().join(",");
      duplicatedValuePairs[sortedKeys] = value;
    }
  }

  // 打印提醒信息
  Object.entries(duplicatedValuePairs).forEach(([keys, value]) => {
    const [key1, key2] = keys.split(",");
    console.log(
      `警告：值 "${value}" 在合并后的键映射中有多个键："${key1}" 和 "${key2}"。`
    );
  });
  // 将合并后的key-value对象转换成YAML格式字符串
  const yamlContent = Object.keys(mergedKeyMapping)
    .map((key) => `${key}: ${mergedKeyMapping[key]}`)
    .join("\n");

  // 创建输出文件夹
  const outputFolderPath = path.join(projectPath, "generatorI18n");
  if (!fs.existsSync(outputFolderPath)) {
    fs.mkdirSync(outputFolderPath);
  }
  // 确定生成 YAML 文件的路径
  const yamlFilePath = path.join(outputFolderPath, `${area}.yaml`);
  // 将YAML内容写入文件
  fs.writeFileSync(yamlFilePath, yamlContent, "utf-8");
}

// 定义匹配$T方法调用的正则表达式
const regex = /(\$?[Tt])\(['"]([^'"]+)['"]\)/g;
// 使用glob模块找到所有Vue和JS文件
const files = glob.sync("**/*.+(vue|js)", {
  cwd: projectPath,
  ignore: "**/node_modules/**", // 忽略node_modules文件夹
});
// 获取 key 映射
readExcel(excelFilePath, oldExcelFilePath).then((res) => {
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
  generatorI18nDoc(
    "en-US",
    oldKeyValueMapping["en-US"],
    newKeyValueMapping["en-US"]
  );
  generatorI18nDoc(
    "zh-CN",
    oldKeyValueMapping["zh-CN"],
    newKeyValueMapping["zh-CN"]
  );
  generatorI18nDoc(
    "zh-HK",
    oldKeyValueMapping["zh-HK"],
    newKeyValueMapping["zh-HK"]
  );
  generatorI18nDoc(
    "ms-MY",
    oldKeyValueMapping["ms-MY"],
    newKeyValueMapping["ms-MY"]
  );
  console.log("Replacement completed.");
});
