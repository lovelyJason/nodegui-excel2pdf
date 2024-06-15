import {
  QMainWindow,
  QWidget,
  QScrollArea,
  QFileDialog,
  QLabel,
  QPushButton,
  QIcon,
  QBoxLayout,
  Direction,
  WindowType,
  WidgetEventTypes,
  QDragMoveEvent,
  QDropEvent,
  QMimeData,
  QMessageBox,
  ButtonRole
} from "@nodegui/nodegui";
import * as path from "node:path";
import sourceMapSupport from "source-map-support";
import * as fs from "node:fs";
// import xlsx from 'node-xlsx'
import ExcelJS from "exceljs";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import os from 'node:os'

/**
 * 记录
 * ①新的 Docxtemplater 实例：在生成每个 Word 文件时，确保创建新的 PizZip 和 Docxtemplater 实例。这避免了共享模板状态的问题。

 */

sourceMapSupport.install();

// 读取Excel文件
async function readExcel(filePath: string) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.worksheets[0];

  const rows = [];
  worksheet.eachRow((row, rowNumber) => {
    rows.push(row.values);
  });

  return rows;
}

// 处理表头，去除空格和特殊字符
function processHeader(header: string): string {
  if (typeof header === 'string') {
    return header.replace(/\s+/g, '_').replace(/[^\w\s]/gi, '');
  } else {
    return String(header);
  }

}

// 生成多个Word文档
// 生成多个Word文档
async function generateWordDocuments(
  excelData: any[],
  wordTemplatePath: string,
  outputDir: string
) {
  try {
    // 读取模板内容
    const templateContent = fs.readFileSync(wordTemplatePath, "binary");

    // 提取从G列开始的所有数据
    const headers = excelData[0].slice(7); // 第7列开始
    const rows = excelData.slice(1).map((row) => row.slice(7));
    // console.log('headers', headers)
    for (let i = 0; i < rows.length; i++) {
      const data = rows[i];
      const templateData: { [key: string]: any } = {};

      headers.forEach((header: string, columnIndex: number) => {
        const columnLetter = String.fromCharCode(71 + columnIndex); // 从G开始
        console.log('column', columnLetter, header)
        templateData[`${columnLetter}-header`] = header;
        templateData[`${columnLetter}-content`] = data[columnIndex];
      });

      // 使用 PizZip 解压缩模板内容
      const zip = new PizZip(templateContent);

      // 使用 Docxtemplater 加载解压后的内容
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // 设置数据并渲染文档
      doc.setData(templateData);
      doc.render();

      // 生成新的Word文件
      const outputFilePath = path.join(outputDir, `result_${i + 1}.docx`);
      const generatedDocBuffer = doc.getZip().generate({ type: "nodebuffer" });
      fs.writeFileSync(outputFilePath, generatedDocBuffer);

      console.log(`Word 文件已成功生成: ${outputFilePath}`);

    }
  } catch (error) {
    // console.error("发生错误:", error);
    throw error
  }
}

function main(): void {
  const win = new QMainWindow();
  win.setWindowTitle("Excel2PDF By JasonHuang QQ315945659");
  win.setMinimumSize(360, 200);
  win.setWindowFlag(WindowType.WindowMaximizeButtonHint, false);

  const centralWidget = new QWidget();
  centralWidget.setAcceptDrops(true);

  centralWidget.addEventListener(WidgetEventTypes.DragEnter, (event) => {
    const dragEvent = new QDragMoveEvent(event);
    const mimeData = dragEvent.mimeData();
    if (mimeData.hasUrls()) {
      const urls = mimeData.urls();
      if (urls.length > 0 && urls[0].toString().endsWith(".xlsx")) {
        dragEvent.acceptProposedAction();
      }
    }
  });

  centralWidget.addEventListener(WidgetEventTypes.Drop, async (event) => {
    const dropEvent = new QDropEvent(event);
    const mimeData = dropEvent.mimeData();
    if (mimeData.hasUrls()) {
      const urls = mimeData.urls();
      if (urls.length > 0 && urls[0].toString().endsWith(".xlsx")) {
        const excelFilePath = urls[0].toString().replace("file://", "");
        const excelData = await readExcel(excelFilePath);
        const wordTemplatePath = path.resolve(process.cwd(), "src/template.docx");
        const outputDir = path.resolve(os.homedir(), 'Desktop/ayesha');
        console.log('outputDir', outputDir);
        if (!fs.existsSync(outputDir)) {
          fs.mkdirSync(outputDir);
        }
        await generateWordDocuments(excelData, wordTemplatePath, outputDir);
        const outputFiles = fs
          .readdirSync(outputDir)
          .map((file) => path.join(outputDir, file));
        const fileContent = outputFiles.join("\n");
        fileContentLabel.setText(fileContent);

        console.log("Excel 文件已成功读取:", excelFilePath);
        console.log("Word 文件已成功生成，文件列表:\n", outputFiles);

        // 弹出通知窗口
        const messageBox = new QMessageBox();
        messageBox.setText('老板，已为您处理完成\n已经放到了桌面的ayesha目录');
        const accept = new QPushButton();
        accept.setText('好的');
        messageBox.addButton(accept, ButtonRole.AcceptRole);
        messageBox.exec();

      }
    }
  });

  const rootLayout = new QBoxLayout(Direction.TopToBottom);
  centralWidget.setObjectName("myroot");
  centralWidget.setLayout(rootLayout);

  const label = new QLabel();
  label.setObjectName("mylabel");
  label.setText("请选择文件");

  const scrollArea = new QScrollArea();
  const fileContentLabel = new QLabel();
  fileContentLabel.setText("Excel Content: None");
  fileContentLabel.setWordWrap(true);
  fileContentLabel.setInlineStyle("width: 800px;"); // 设置宽度
  scrollArea.setWidget(fileContentLabel);
  scrollArea.setWidgetResizable(true);
  scrollArea.setInlineStyle("max-height: 600px;"); // 设置最大高度

  const projectRoot = process.cwd();
  console.log("项目根目录:", projectRoot);

  const button = new QPushButton();
  button.setIcon(new QIcon(path.join(__dirname, "../assets/logo.jpeg")));
  button.setInlineStyle("height: 30px");
  button.setText("选择Excel(支持拖拽excel进来)");
  button.addEventListener("clicked", async () => {
    const fileDialog = new QFileDialog();
    fileDialog.setFileMode(0); // 0 means QFileDialog.FileMode.AnyFile
    fileDialog.setNameFilter("Excel Files (*.xlsx)");

    if (fileDialog.exec()) {
      const selectedFiles = fileDialog.selectedFiles();
      const excelFilePath = selectedFiles[0];
      try {
        const excelData = await readExcel(excelFilePath);
        const wordTemplatePath = path.resolve(projectRoot, "src/template.docx");
        const outputDir = path.resolve(os.homedir(), 'Desktop/ayesha');
        console.log('outputDir', outputDir)
        if (!fs.existsSync(outputDir)) {
          fs.mkdirSync(outputDir);
        }
        // console.log('excelData', excelData)
        await generateWordDocuments(excelData, wordTemplatePath, outputDir);
        const outputFiles = fs
          .readdirSync(outputDir)
          .map((file) => path.join(outputDir, file));
        const fileContent = outputFiles.join("\n");
        fileContentLabel.setText(fileContent);

        console.log("Excel 文件已成功读取:", excelFilePath);
        console.log("Word 文件已成功生成，文件列表:\n", outputFiles);

        const messageBox = new QMessageBox();
        messageBox.setText('老板，已为您处理完成\n已经放到了桌面的ayesha目录');
        const accept = new QPushButton();
        accept.setText('好的');
        messageBox.addButton(accept, ButtonRole.AcceptRole);
        messageBox.exec();
      } catch (error) {
        console.error("发生错误:", error);
        // 弹出错误信息窗口
        const messageBox = new QMessageBox();
        messageBox.setText(error.message);
        const accept = new QPushButton();
        accept.setText('确认');
        messageBox.addButton(accept, ButtonRole.AcceptRole);
        messageBox.exec();
      }
    }
  });

  rootLayout.addWidget(button);
  rootLayout.addWidget(scrollArea);
  // rootLayout.addWidget(label);
  win.setCentralWidget(centralWidget);
  win.setStyleSheet(
    `
    #myroot {
      background-color: #009688;
      height: '100%';
      align-items: 'center';
      justify-content: 'center';
    }
    #mylabel {
      font-size: 16px;
      font-weight: bold;
      padding: 1;
    }
  `
  );
  win.show();

  (global as any).win = win;
}
main();
