const XLSX = require('xlsx');

// 文件路径定义
const tagFilePath = '20250228.xlsx'; 
const MSTagFilePath = '02_MSTag.xlsx';

// 读取 Excel 文件函数
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    
    return data;
}

// 读取两个 Excel 文件
const MSTag = readExcel(MSTagFilePath).slice(3); // 去除前三行
const tagData = readExcel(tagFilePath);

// 检查 Translate 是否包含 MSTag 的 Simp_TIMI，并添加 newTranslate 字段
function checkAndFormatTranslate(item, index) {
  let translate = item['newTranslate']; // 获取B列的值

  // 确保 translate 是一个字符串
  if (typeof translate !== 'string') {
    item['tagNewTranslate'] = ''; // 如果 translate 不是字符串，设置 tagNewTranslate 为空字符串
    return item;
  }

  // 复制原始的 translate 字符串，用于替换
  let newTranslate = translate;

  // 遍历 MSTag 数组
  let hasMatch = false;
  MSTag.forEach(tag => {
    const regex = new RegExp(tag.Simp_TIMI, 'g');
    if (regex.test(newTranslate)) {
      hasMatch = true;
      newTranslate = newTranslate.replace(regex, `{${tag.Simp_TIMI}}`);
    }
  });

  // 如果没有匹配，则将 tagNewTranslate 设置为空字符串
  if (!hasMatch) {
    item['tagNewTranslate'] = '';
  } else {
    item['tagNewTranslate'] = newTranslate;
  }

  return item;
}

// 遍历 tagData 数组并应用 checkAndFormatTranslate 函数
const updatedTagData = tagData.map(checkAndFormatTranslate);

// 将 JSON 数据转换为工作表
const worksheet = XLSX.utils.json_to_sheet(updatedTagData);

// 在B列后插入新的列（tagNewTranslate）
const headers = ['tagNewTranslate'];
headers.reverse().forEach(header => XLSX.utils.sheet_add_aoa(worksheet, [[header]], { origin: 'D1' }));

// 创建一个新的工作簿并添加工作表
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// 写入 Excel 文件
XLSX.writeFile(workbook, 'updatedTagTranslate.xlsx');