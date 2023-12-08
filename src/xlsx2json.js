import XLSX from "xlsx";
import { unflatten } from "flat";
import { glob } from "glob";
import { writeFile, readdir, unlink, mkdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import path, { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// 原文件目录
const srcFolderPath = resolve(__dirname, "../files");
// 存储目录
const folderPath = resolve(__dirname, "../dist");

// 表格key列，列名
const langKey = 'key'
// 表格value列，列名
const langValues = ['cn', 'en']

// 文件路径正斜杠或者反斜杠
function getFilePathDelimiter() {
  // 获取当前系统
  const currentSystem = process.platform;
  console.log("当前系统为：" + currentSystem);
  if (currentSystem === "win32") {
    // Windows系统
    return "\\";
  } else {
    // Linux/MacOS等其他系统
    return "/";
  }
}

// 生成前需要删除已生成的目录
async function deleteAllFilesInFolder(folderPath) {
  if (!existsSync(folderPath)) {
    await mkdir(folderPath);
  }
  const files = await readdir(folderPath);
  for (const file of files) {
    const filePath = path.join(folderPath, file);
    await unlink(filePath);
  }
}

// 转换sheet中的数据格式为json [{key: 'global.operate', cn: '操作', en: 'Operate'}]
function getSheetJsonData(sheetData, fileJsonObj) {
  sheetData = sheetData.filter((item) => item.key);
  langValues.forEach((lang) => {
    let jsonData = sheetData.reduce((obj, item) => {
      obj[item[langKey].trim()] = item[lang];
      return obj;
    }, {});
    jsonData = unflatten(jsonData);
    fileJsonObj[lang] = {...fileJsonObj[lang], ...jsonData};
  });
  return fileJsonObj;
}

// 转换主函数
async function transfer() {
  try {
    const delimiter = getFilePathDelimiter()
    await deleteAllFilesInFolder(folderPath);

    const xlsxFiles = await glob(`${srcFolderPath}/**/*.{xlsx,xls}`, {
      ignore: `${folderPath}/**`,
    });
    // console.log("xlsxFiles", xlsxFiles);
    xlsxFiles.forEach(async (file) => {
      const fileName = file
        .split(delimiter)
        .pop()
        .replace(".xlsx", "")
        .replace(".xls", "");

      const fileJsonObj = {}
      langValues.forEach(lang => {
        fileJsonObj[lang] = {}
      })

      const workbook = XLSX.readFile(file);
      workbook.SheetNames.forEach((sheet) => {
        const worksheet = workbook.Sheets[sheet];
        const sheetData = XLSX.utils.sheet_to_json(worksheet);
        getSheetJsonData(sheetData, fileJsonObj);
      });

      Object.keys(fileJsonObj).forEach(async lang => {
        // 将JSON数据转换为字符串
        let langJsonString = JSON.stringify(fileJsonObj[lang], "", "\t");
        if (!existsSync(folderPath)) {
          await mkdir(folderPath);
        }
        await writeFile(`${folderPath}/${fileName}.${lang}.json`, langJsonString);
      })
    });
  } catch (error) {
    console.log("[error]", error);
  }
}

transfer();
