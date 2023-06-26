const { app, BrowserWindow, dialog } = require('electron');
const fs = require('fs');
const xlsx = require('xlsx');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
    },
  });

  mainWindow.loadFile('index.html');

  mainWindow.on('closed', function () {
    mainWindow = null;
  });
}

app.on('ready', createWindow);

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', function () {
  if (mainWindow === null) {
    createWindow();
  }
});

// 处理选择原始 Excel 文件按钮点击事件
function selectInputFile() {
  dialog.showOpenDialog(mainWindow, { properties: ['openFile'] }).then((result) => {
    if (!result.canceled) {
      const filePath = result.filePaths[0];
      readExcelFile(filePath);
    }
  });
}

// 读取 Excel 文件
function readExcelFile(filePath) {
  try {
    const workbook = xlsx.readFile(filePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    processExcelData(jsonData);
  } catch (error) {
    console.error('Error reading Excel file:', error);
  }
}

// 处理 Excel 数据
function processExcelData(data) {
  // 在这里进行数据处理和转换
  // 示例：将数据处理为要输出的格式

  const outputData = data.map((row) => row.join(','));

  // 将处理后的数据写入新的 Excel 文件
  const outputWorkbook = xlsx.utils.book_new();
  const outputWorksheet = xlsx.utils.aoa_to_sheet(outputData);
  xlsx.utils.book_append_sheet(outputWorkbook, outputWorksheet, 'Sheet1');
  const outputPath = dialog.showSaveDialogSync(mainWindow, { defaultPath: 'output.xlsx' });
  if (outputPath) {
    xlsx.writeFile(outputWorkbook, outputPath);
  }
}

module.exports = {
  selectInputFile,
};
