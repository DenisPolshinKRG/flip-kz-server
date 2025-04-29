require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const PDFDocument = require('pdfkit');
const bwipjs = require('bwip-js');
const fs = require('fs');
const path = require('path');
const app = express();
const port = 3000;

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();
  next();
});

app.use(express.json());
app.use('/pdfs', express.static(path.join(__dirname, 'pdfs')));

const pdfDir = path.join(__dirname, 'pdfs');
if (!fs.existsSync(pdfDir)) fs.mkdirSync(pdfDir);

const credentials = require('./credentials.json');
console.log('Loaded credentials:', credentials);

const spreadsheetId = '1_Tj2MdjT0H4AhVTQiuG50HgQmugNQ_pAV7SsT9pJ5Is';
const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;

const auth = new google.auth.GoogleAuth({
  credentials: credentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });

function formatDate(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

function sanitizeString(str) {
  if (!str) return '';
  return str
    .replace(/\n/g, ' ')
    .replace(/[«»"]/g, '')
    .replace(/&/g, 'and')
    .trim();
}

async function generateBarcode(text, labelSize) {
  const height = labelSize === '30x20' ? 10 : 20;
  return new Promise((resolve, reject) => {
    bwipjs.toBuffer(
      {
        bcid: 'code128',
        text: text,
        scale: 2,
        height: height,
        includetext: false,
        textalign: 'center',
      },
      (err, png) => {
        if (err) reject(err);
        else resolve(png);
      }
    );
  });
}

app.post('/export', async (req, res) => {
  try {
    const orders = req.body.orders;
    if (!orders || !Array.isArray(orders)) {
      return res.status(400).json({ error: 'Invalid orders data' });
    }

    const table2Response = await sheets.spreadsheets.values.get({
      spreadsheetId: spreadsheetId,
      range: 'Лист2!A:C',
    });

    const table2Rows = table2Response.data.values || [];
    const supplierCodeToTextMap = new Map();
    const supplierCodeToExpirationMap = new Map();
    for (let i = 1; i < table2Rows.length; i++) {
      const supplierCode = table2Rows[i][0]?.trim();
      const replacementText = table2Rows[i][1]?.trim();
      const expirationValue = table2Rows[i][2]?.trim();
      if (supplierCode && replacementText) {
        supplierCodeToTextMap.set(supplierCode, replacementText);
      }
      if (supplierCode && expirationValue) {
        supplierCodeToExpirationMap.set(supplierCode, expirationValue);
      }
    }

    await sheets.spreadsheets.values.clear({
      spreadsheetId: spreadsheetId,
      range: 'Лист1!A1:ZZ10000',
    });

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId,
      resource: {
        requests: [{
          repeatCell: {
            range: {
              sheetId: 0,
              startRowIndex: 0,
              endRowIndex: 10000,
              startColumnIndex: 0,
              endColumnIndex: 702,
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 1, green: 1, blue: 1 },
                textFormat: { bold: false, italic: false, underline: false, strikethrough: false, foregroundColor: { red: 0, green: 0, blue: 0 } },
                horizontalAlignment: 'LEFT',
                verticalAlignment: 'BOTTOM',
                borders: {},
              },
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,borders)',
          },
        }],
      },
    });

    const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: spreadsheetId });
    const sheet = spreadsheet.data.sheets.find(s => s.properties.sheetId === 0);
    const conditionalFormatRules = sheet.conditionalFormats || [];
    for (let i = 0; i < conditionalFormatRules.length; i++) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheetId,
        resource: {
          requests: [{ deleteConditionalFormatRule: { sheetId: 0, index: 0 } }],
        },
      });
    }

    // Добавляем текст "ПОСТАВКА (К)" в C1
    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId,
      range: 'Лист1!C1',
      valueInputOption: 'RAW',
      resource: { values: [['ПОСТАВКА (К)']] },
    });

    // Форматируем ячейку C1: жирный шрифт, размер 20, выравнивание по центру
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId,
      resource: {
        requests: [
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 2, endColumnIndex: 3 },
              cell: {
                userEnteredFormat: {
                  textFormat: { bold: true, fontSize: 20 },
                  horizontalAlignment: 'CENTER',
                },
              },
              fields: 'userEnteredFormat(textFormat,horizontalAlignment)',
            },
          },
        ],
      },
    });

    const headers = ['Код поставщика', 'КОД FLIP', 'Наименование + ссылка на товар', 'Кол-во', 'Цена', 'Сумма', 'Срок годности', 'Дата заказа'];
    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId,
      range: 'Лист1!A3:H3',
      valueInputOption: 'RAW',
      resource: { values: [headers] },
    });

    const newRows = orders.map(order => {
      const quantity = parseInt(order.quantity, 10) || 0;
      const price = parseInt(order.price.replace(/[^0-9]/g, ''), 10) || 0;
      const amount = quantity * price;
      const date = formatDate(new Date());
      const productUrl = `https://www.flip.kz/catalog?prod=${order.flipCode}`;
      let displayText = sanitizeString(order.productName);
      const replacementText = supplierCodeToTextMap.get(order.supplierCode?.trim());
      if (replacementText) displayText = sanitizeString(replacementText);
      const productCell = '=ГИПЕРССЫЛКА(' + JSON.stringify(productUrl) + '; ' + JSON.stringify(displayText) + ')';
      const expiration = supplierCodeToExpirationMap.get(order.supplierCode?.trim()) || '';
      return [order.supplierCode, order.flipCode, productCell, quantity, price, amount, expiration, date];
    });

    const rangeStart = 4;
    const rangeEnd = rangeStart + newRows.length - 1;
    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId,
      range: `Лист1!A${rangeStart}:H${rangeEnd}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: newRows },
    });

    let totalQuantity = 0;
    let totalAmount = 0;
    for (const order of orders) {
      const qty = parseInt(order.quantity, 10) || 0;
      const prc = parseInt(order.price.replace(/[^0-9]/g, ''), 10) || 0;
      totalQuantity += qty;
      totalAmount += qty * prc;
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId,
      range: `Лист1!A${rangeEnd + 1}:H${rangeEnd + 1}`,
      valueInputOption: 'RAW',
      resource: { values: [['', '', `ИТОГО:`, totalQuantity, '', totalAmount, '', '']] },
    });

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId,
      resource: {
        requests: [
          {
            updateBorders: {
              range: { sheetId: 0, startRowIndex: 2, endRowIndex: rangeEnd + 1, startColumnIndex: 0, endColumnIndex: 8 },
              top: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              bottom: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              left: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              right: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              innerHorizontal: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
              innerVertical: { style: 'SOLID', width: 1, color: { red: 0, green: 0, blue: 0 } },
            },
          },
          {
            updateBorders: {
              range: { sheetId: 0, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 8 },
              top: { style: 'SOLID_MEDIUM', width: 2, color: { red: 0, green: 0, blue: 0 } },
              bottom: { style: 'SOLID_MEDIUM', width: 2, color: { red: 0, green: 0, blue: 0 } },
              left: { style: 'SOLID_MEDIUM', width: 2, color: { red: 0, green: 0, blue: 0 } },
              right: { style: 'SOLID_MEDIUM', width: 2, color: { red: 0, green: 0, blue: 0 } },
              innerHorizontal: { style: 'SOLID_MEDIUM', width: 2, color: { red: 0, green: 0, blue: 0 } },
              innerVertical: { style: 'SOLID_MEDIUM', width: 2, color: { red: 0, green: 0, blue: 0 } },
            },
          },
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: 2, endRowIndex: 3, startColumnIndex: 0, endColumnIndex: 8 },
              cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 0 }, textFormat: { foregroundColor: { red: 0, green: 0, blue: 0 }, bold: true }, horizontalAlignment: 'CENTER' } },
              fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
            },
          },
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: rangeEnd, endRowIndex: rangeEnd + 1, startColumnIndex: 0, endColumnIndex: 8 },
              cell: { userEnteredFormat: { backgroundColor: { red: 0, green: 1, blue: 0 }, textFormat: { bold: true }, horizontalAlignment: 'RIGHT' } },
              fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
            },
          },
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: 3, endRowIndex: rangeEnd, startColumnIndex: 3, endColumnIndex: 4 },
              cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
              fields: 'userEnteredFormat(horizontalAlignment)',
            },
          },
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: 3, endRowIndex: rangeEnd, startColumnIndex: 4, endColumnIndex: 6 },
              cell: { userEnteredFormat: { horizontalAlignment: 'RIGHT' } },
              fields: 'userEnteredFormat(horizontalAlignment)',
            },
          },
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: 3, endRowIndex: rangeEnd, startColumnIndex: 6, endColumnIndex: 7 },
              cell: { userEnteredFormat: { horizontalAlignment: 'CENTER' } },
              fields: 'userEnteredFormat(horizontalAlignment)',
            },
          },
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: 3, endRowIndex: rangeEnd, startColumnIndex: 7, endColumnIndex: 8 },
              cell: { userEnteredFormat: { horizontalAlignment: 'RIGHT' } },
              fields: 'userEnteredFormat(horizontalAlignment)',
            },
          },
          {
            repeatCell: {
              range: { sheetId: 0, startRowIndex: 3, endRowIndex: rangeEnd, startColumnIndex: 2, endColumnIndex: 3 },
              cell: { userEnteredFormat: { horizontalAlignment: 'LEFT' } },
              fields: 'userEnteredFormat(horizontalAlignment)',
            },
          },
          { updateDimensionProperties: { range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 120 }, fields: 'pixelSize' } },
          { updateDimensionProperties: { range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 80 }, fields: 'pixelSize' } },
          { updateDimensionProperties: { range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 3, endIndex: 7 }, properties: { pixelSize: 100 }, fields: 'pixelSize' } },
          { updateDimensionProperties: { range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 7, endIndex: 8 }, properties: { pixelSize: 100 }, fields: 'pixelSize' } },
          { autoResizeDimensions: { dimensions: { sheetId: 0, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 } } },
          {
            addConditionalFormatRule: {
              rule: {
                ranges: [{ sheetId: 0, startRowIndex: 3, endRowIndex: rangeEnd, startColumnIndex: 0, endColumnIndex: 8 }],
                booleanRule: { condition: { type: 'NUMBER_EQ', values: [{ userEnteredValue: '0' }] }, format: { backgroundColor: { red: 1, green: 0, blue: 0 } } },
              },
              index: 0,
            },
          },
        ],
      },
    });

    res.json({ success: true, message: 'Orders exported successfully', spreadsheetUrl });
  } catch (error) {
    console.error('Error exporting to Google Sheets:', error);
    res.status(500).json({ error: 'Failed to export orders' });
  }
});

app.post('/print-barcodes', async (req, res) => {
  try {
    const { orders, labelSize } = req.body;
    if (!orders || !Array.isArray(orders) || orders.length === 0) {
      return res.status(400).json({ error: 'Invalid orders data' });
    }

    let labelWidth, labelHeight, barcodeHeight, barcodeTextFontSize, flipCodeFontSize, productNameFontSize, barcodeY, barcodeTextY, flipCodeY, productNameY;
    if (labelSize === '30x20') {
      labelWidth = 30 * 2.83464567;
      labelHeight = 20 * 2.83464567;
      barcodeHeight = 25;
      barcodeTextFontSize = 6;
      flipCodeFontSize = 6;
      productNameFontSize = 4;
      barcodeY = 2;
      barcodeTextY = barcodeHeight + 2;
      flipCodeY = barcodeHeight + 10;
      productNameY = barcodeHeight + 17;
    } else {
      labelWidth = 58 * 2.83464567;
      labelHeight = 40 * 2.83464567;
      barcodeHeight = 50;
      barcodeTextFontSize = 8;
      flipCodeFontSize = 10;
      productNameFontSize = 8;
      barcodeY = 5;
      barcodeTextY = barcodeHeight + 5;
      flipCodeY = barcodeHeight + 15;
      productNameY = barcodeHeight + 27;
    }

    const pageSize = [labelWidth, labelHeight];
    let totalLabels = 0;
    for (const order of orders) {
      const quantity = parseInt(order.quantity, 10) || 0;
      if (quantity > 0) totalLabels += quantity;
    }
    if (totalLabels === 0) return res.status(400).json({ error: 'No labels to print' });

    const doc = new PDFDocument({ size: pageSize, margin: 0 });
    const filename = `barcodes_${Date.now()}.pdf`;
    const filePath = path.join(pdfDir, filename);
    const stream = fs.createWriteStream(filePath);
    doc.pipe(stream);

    doc.registerFont('Arial', path.join(__dirname, 'arial.ttf'));

    let firstPage = true;
    for (const order of orders) {
      const quantity = parseInt(order.quantity, 10) || 0;
      if (quantity === 0) continue;

      const barcodeText = order.supplierCode || 'N/A';
      const flipCodeText = `КОД FLIP - ${order.flipCode || 'N/A'}`;
      const productName = sanitizeString(order.productName).substring(0, labelSize === '30x20' ? 50 : 50);

      const barcodeBuffer = await generateBarcode(barcodeText, labelSize);

      for (let i = 0; i < quantity; i++) {
        if (!firstPage) {
          doc.addPage({ size: pageSize, margin: 0 });
        }
        firstPage = false;

        // Убрали черную рамку, удалив doc.rect(0, 0, labelWidth, labelHeight).stroke()
        const barcodeWidth = labelWidth - (labelSize === '30x20' ? 10 : 20); // Увеличили отступы, чтобы сузить штрихкод
        doc.image(barcodeBuffer, (labelWidth - barcodeWidth) / 2, barcodeY, { width: barcodeWidth, height: barcodeHeight });
        doc.font('Arial').fontSize(barcodeTextFontSize).text(barcodeText, 2, barcodeTextY, { width: labelWidth - 4, align: 'center' });
        doc.font('Arial').fontSize(flipCodeFontSize).text(flipCodeText, 2, flipCodeY, { width: labelWidth - 4, align: 'center' });
        doc.font('Arial').fontSize(productNameFontSize).text(productName, 2, productNameY, { width: labelWidth - 4, align: 'center', lineBreak: true });
      }
    }

    doc.end();
    await new Promise(resolve => stream.on('finish', resolve));

    const pdfUrl = `${req.protocol}://${req.get('host')}/pdfs/${filename}`;
    res.json({ success: true, pdfUrl });
  } catch (error) {
    console.error('Error generating barcode PDF:', error);
    res.status(500).json({ error: 'Failed to generate barcode PDF' });
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});